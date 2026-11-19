//! The scan behind `read_xlsx` and `COPY ... FROM`: reads the rows of a sheet as text and lets DuckDB cast them
//! to the column types.

use std::{
    fmt,
    ops::Range,
    sync::atomic::{AtomicU64, Ordering},
};

use duckdb_neo::{
    DuckDBType, Error, Result,
    connection::Context,
    data_chunk::{DataChunk, DataChunkRef},
    file::FileSystem,
    logical_type::{LogicalType, LogicalTypeID},
    types::{DateValue, TimeValue, TimestampValue},
    vector::{Unknown, Vector, WritableNativeElement},
};

use crate::{
    error::{internal, invalid_input, raw_message},
    xlsx::{
        archive::Archive,
        cell::{CellPos, CellRange, CellType},
        options::ReadOptions,
        parsers::{
            row_collector::{CellValue, RowBatch, RowReader},
            shared_strings::SharedStringsParser,
        },
        resolve::{SHARED_STRINGS_PATH, SheetLayout},
        serial::{DateSystem, SerialError, serial_to_date, serial_to_time, serial_to_timestamp},
        strings::StringTable,
    },
};

/// How the text of the cells of a column becomes values of the column type: a text column takes it as it is,
/// the numbers of a date or time column are serial numbers, and DuckDB casts everything else. A serial number
/// that is no value of the column type is a cell that does not convert.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Conversion {
    Text,
    Cast,
    SerialDate,
    SerialTime,
    SerialTimestamp,
}

/// A column the scan returns.
#[derive(Debug)]
pub struct ScanColumn {
    pub logical_type: LogicalType,
    conversion: Conversion,
}

/// What a scan reads and returns, decided when the statement is bound.
///
/// A scan of a range the caller chose that does not stop at empty rows returns rows of NULLs up to the end of
/// the range.
#[derive(Debug)]
pub struct ScanPlan {
    file_path: String,
    sheet_path: String,
    range: CellRange,
    stop_at_empty: bool,
    fill_range: bool,
    ignore_errors: bool,
    date_system: DateSystem,
    pub columns: Vec<ScanColumn>,
}

impl ScanPlan {
    /// Plans to return the columns of `layout` as `types`, one type for each column.
    pub fn new(
        file_path: String,
        layout: SheetLayout,
        options: &ReadOptions,
        types: Vec<LogicalType>,
    ) -> Result<ScanPlan> {
        if types.len() != layout.columns.len()
            || u64::try_from(types.len()).ok() != Some(layout.range.width())
        {
            return Err(invalid_input(
                "read_xlsx: the columns of the sheet do not match its range (is the file corrupted?)",
            ));
        }
        let columns = types
            .into_iter()
            .zip(&layout.columns)
            .map(|(logical_type, column)| {
                let from_number = column.source_type == CellType::Number;
                let conversion = match logical_type.id()? {
                    LogicalTypeID::Varchar => Conversion::Text,
                    LogicalTypeID::Date if from_number => Conversion::SerialDate,
                    LogicalTypeID::Time if from_number => Conversion::SerialTime,
                    LogicalTypeID::Timestamp if from_number => Conversion::SerialTimestamp,
                    _ => Conversion::Cast,
                };
                Ok(ScanColumn {
                    logical_type,
                    conversion,
                })
            })
            .collect::<Result<Vec<_>>>()?;
        Ok(ScanPlan {
            file_path,
            sheet_path: layout.sheet_path,
            range: layout.range,
            stop_at_empty: options.stop_at_empty,
            fill_range: options.has_explicit_range && !options.stop_at_empty,
            ignore_errors: options.ignore_errors,
            date_system: layout.date_system,
            columns,
        })
    }
}

/// How far a scan has read into its sheet; shared with the thread that reports progress.
#[derive(Debug, Default)]
pub struct ScanProgress {
    position: AtomicU64,
    length: AtomicU64,
}

impl ScanProgress {
    /// The part of the sheet that has been read, between 0 and 1.
    pub fn fraction(&self) -> f64 {
        let position = self.position.load(Ordering::Relaxed) as f64;
        let length = self.length.load(Ordering::Relaxed) as f64;
        if position == 0.0 || length == 0.0 {
            0.0
        } else {
            position / length
        }
    }
}

/// A scan in progress: the open sheet and the rows read ahead.
#[derive(Debug)]
pub struct SheetScan {
    rows: RowReader<Archive>,
    strings: StringTable,
}

impl SheetScan {
    /// Opens the workbook, loads its shared strings and opens the sheet; `progress` is the state all workers
    /// of the scan share, as DuckDB hands it over.
    pub fn open(
        context: &Context,
        plan: &ScanPlan,
        progress: Option<&ScanProgress>,
    ) -> Result<Option<SheetScan>> {
        let progress = progress.ok_or_else(missing_state)?;
        let mut archive = Archive::open(&FileSystem::new(context)?, &plan.file_path)?;
        let mut strings = StringTable::new();
        archive.parse_entry(
            SHARED_STRINGS_PATH,
            &mut SharedStringsParser::new(&mut strings),
        )?;
        if !archive.try_open_entry(&plan.sheet_path)? {
            return Err(invalid_input(format!(
                "Sheet '{}' not found in xlsx file",
                plan.sheet_path
            )));
        }
        progress
            .length
            .store(archive.entry_len(), Ordering::Relaxed);
        progress.position.store(0, Ordering::Relaxed);
        let rows = RowReader::new(
            archive,
            plan.range,
            plan.stop_at_empty,
            plan.fill_range,
            strings.len(),
        );
        Ok(Some(SheetScan { rows, strings }))
    }

    /// Writes the next rows of `scan` into `output`, as many as it holds; none when the sheet is exhausted.
    pub fn exec(
        context: &Context,
        plan: &ScanPlan,
        scan: Option<&mut SheetScan>,
        progress: Option<&ScanProgress>,
        output: DataChunkRef<'_>,
    ) -> Result<()> {
        let (Some(scan), Some(progress)) = (scan, progress) else {
            return Err(missing_state());
        };
        scan.rows.read_batch(output.capacity()?)?;
        progress
            .position
            .store(scan.rows.bytes_read(), Ordering::Relaxed);
        let emitter = Emitter {
            context,
            plan,
            batch: scan.rows.batch(),
            strings: &scan.strings,
        };
        emitter.emit(output)
    }
}

fn missing_state() -> Error {
    internal("xlsx: the scan state is missing")
}

/// A chunk whose vectors a cast can write to and whose cast results can be read back.
trait CastTarget {
    fn target(&mut self, index: usize) -> Result<Vector<'_, Unknown>>;

    fn result(&self, index: usize) -> Result<&Vector<'_, Unknown>>;
}

impl CastTarget for DataChunkRef<'_> {
    fn target(&mut self, index: usize) -> Result<Vector<'_, Unknown>> {
        self.get_vector_at_mut(index)
    }

    fn result(&self, index: usize) -> Result<&Vector<'_, Unknown>> {
        self.get_vector_at(index)
    }
}

impl CastTarget for DataChunk<'_> {
    fn target(&mut self, index: usize) -> Result<Vector<'_, Unknown>> {
        self.get_vector_at_mut(index)
    }

    fn result(&self, index: usize) -> Result<&Vector<'_, Unknown>> {
        self.get_vector_at(index)
    }
}

/// Where the text of the cells of a column goes: into the output, or into a vector that is cast afterwards.
#[derive(Debug, Clone, Copy)]
enum TextTarget {
    Output(usize),
    Scratch(usize),
}

/// The cells a block of rows has at most when a batch is written to its vectors, and the rows it has at least.
const BLOCK_CELLS: usize = 32 * 1024;
const MIN_BLOCK_ROWS: usize = 32;

/// Turns a batch of rows into the vectors of an output chunk.
///
/// A cell that does not convert becomes NULL when it is empty or errors are ignored, and fails the scan
/// otherwise, with the reason that this cell does not convert. Columns the output leaves out are converted
/// all the same, so that their cells are checked.
struct Emitter<'a> {
    context: &'a Context,
    plan: &'a ScanPlan,
    batch: &'a RowBatch,
    strings: &'a StringTable,
}

impl Emitter<'_> {
    fn emit(self, mut output: DataChunkRef<'_>) -> Result<()> {
        let rows = self.batch.len();
        // A scan that is only asked for its row count gets a chunk with the first column alone.
        let present = output.col_count().min(self.plan.columns.len());
        if rows == 0 {
            if present > 0 {
                output.get_vector_at_mut::<Unknown>(0)?.set_size(0)?;
            }
            return Ok(());
        }

        let columns = &self.plan.columns;
        let mut targets = Vec::with_capacity(columns.len());
        let mut casts = 0;
        for (index, column) in columns.iter().enumerate() {
            if column.conversion != Conversion::Text {
                targets.push((index, TextTarget::Scratch(casts)));
                casts += 1;
            } else if index < present {
                targets.push((index, TextTarget::Output(index)));
            }
        }
        if casts == 0 {
            return self.write_texts(&targets, &mut output, None);
        }
        let varchar = String::logical_type(self.context)?;
        let mut texts = DataChunk::create(self.context, vec![varchar; casts], true)?;
        let mut doubles =
            DataChunk::create(self.context, [f64::logical_type(self.context)?], true)?;
        let unseen_types: Vec<LogicalType> = columns
            .iter()
            .skip(present)
            .filter(|column| column.conversion == Conversion::Cast)
            .map(|column| column.logical_type.clone())
            .collect();
        let mut unseen = if unseen_types.is_empty() {
            None
        } else {
            Some(DataChunk::create(self.context, unseen_types, true)?)
        };
        self.write_texts(&targets, &mut output, Some(&mut texts))?;

        let mut next_text = 0;
        let mut next_unseen = 0;
        for (index, column) in columns.iter().enumerate() {
            if column.conversion == Conversion::Text {
                continue;
            }
            let is_present = index < present;
            let text = texts.get_vector_at::<String>(next_text)?;
            next_text += 1;
            match column.conversion {
                Conversion::Cast if is_present => self.cast(text, &mut output, index, index)?,
                Conversion::Cast => {
                    if let Some(unseen) = unseen.as_mut() {
                        self.cast(text, unseen, next_unseen, index)?;
                        next_unseen += 1;
                    }
                }
                Conversion::SerialDate => {
                    let system = self.plan.date_system;
                    let date = |serial| serial_to_date(serial, system).map(DateValue);
                    self.cast(text, &mut doubles, 0, index)?;
                    let target = is_present.then_some(&mut output);
                    self.write_serials(&doubles, target, index, "DATE", date)?;
                }
                Conversion::SerialTime => {
                    let time = |serial| serial_to_time(serial).map(TimeValue);
                    self.cast(text, &mut doubles, 0, index)?;
                    let target = is_present.then_some(&mut output);
                    self.write_serials(&doubles, target, index, "TIME", time)?;
                }
                Conversion::SerialTimestamp => {
                    let system = self.plan.date_system;
                    let timestamp =
                        |serial| serial_to_timestamp(serial, system).map(TimestampValue);
                    self.cast(text, &mut doubles, 0, index)?;
                    let target = is_present.then_some(&mut output);
                    self.write_serials(&doubles, target, index, "TIMESTAMP", timestamp)?;
                }
                Conversion::Text => {}
            }
        }
        Ok(())
    }

    fn write_texts(
        &self,
        targets: &[(usize, TextTarget)],
        output: &mut DataChunkRef<'_>,
        mut scratch: Option<&mut DataChunk<'_>>,
    ) -> Result<()> {
        let rows = self.batch.len();
        // The cells of a wide sheet are written a block of rows at a time, so that what a block reads and
        // writes stays in the cache: the batch holds them row by row, the vectors column by column.
        let block_rows = (BLOCK_CELLS / self.plan.columns.len().max(1)).max(MIN_BLOCK_ROWS);
        let mut first = 0;
        while first < rows {
            let block = first..rows.min(first.saturating_add(block_rows));
            for (column, target) in targets {
                let vector = match (target, scratch.as_mut()) {
                    (TextTarget::Output(index), _) => output.get_vector_at_mut(*index)?,
                    (TextTarget::Scratch(index), Some(scratch)) => {
                        scratch.get_vector_at_mut(*index)?
                    }
                    (TextTarget::Scratch(_), None) => return Err(missing_state()),
                };
                self.write_text(vector, *column, block.clone())?;
            }
            first = block.end;
        }
        Ok(())
    }

    fn write_text(
        &self,
        mut vector: Vector<'_, String>,
        column: usize,
        rows: Range<usize>,
    ) -> Result<()> {
        vector.set_size(self.batch.len())?;
        let mut writer = vector.writer()?;
        for row in rows {
            let text = match self.batch.cell(row, column) {
                CellValue::Null => None,
                CellValue::Text(text) => Some(text),
                CellValue::Shared(index) => self.strings.get(index),
            };
            writer.set_option(row, text)?;
        }
        Ok(())
    }

    fn cast(
        &self,
        text: &Vector<'_, String>,
        target: &mut impl CastTarget,
        index: usize,
        column: usize,
    ) -> Result<()> {
        let failure = self.try_cast(text, target.target(index)?)?;
        let Some(failure) = failure else {
            return Ok(());
        };
        if self.plan.ignore_errors {
            return Ok(());
        }
        let values = target.result(index)?;
        let is_null = values.reader()?;
        for row in 0..self.batch.len() {
            let text = match self.batch.cell(row, column) {
                CellValue::Null => continue,
                CellValue::Text(text) => text,
                CellValue::Shared(index) => self.strings.get(index).unwrap_or_default(),
            };
            if !text.is_empty() && is_null.is_null(row)? {
                // The failure of the batch may be that of another cell, so this one is cast on its own.
                let own = self.cast_alone(text, values.logical_type())?;
                let reason = own.as_ref().unwrap_or(&failure);
                return Err(self.cell_error(row, column, raw_message(reason)));
            }
        }
        Ok(())
    }

    fn try_cast(
        &self,
        text: &Vector<'_, String>,
        target: Vector<'_, Unknown>,
    ) -> Result<Option<Error>> {
        // SAFETY: a text vector is filled before its cast and not written again, and the target is only read
        // once the cast is done.
        unsafe { target.try_cast_rows(text, self.context) }
    }

    fn cast_alone(&self, text: &str, target_type: &LogicalType) -> Result<Option<Error>> {
        let varchar = String::logical_type(self.context)?;
        let mut source = DataChunk::create(self.context, [varchar], true)?;
        let mut vector = source.get_vector_at_mut::<String>(0)?;
        vector.set_size(1)?;
        vector.writer()?.set_option(0, Some(text))?;
        let mut target = DataChunk::create(self.context, [target_type.clone()], true)?;
        self.try_cast(
            source.get_vector_at::<String>(0)?,
            target.get_vector_at_mut(0)?,
        )
    }

    fn cell_error(&self, row: usize, column: usize, reason: impl fmt::Display) -> Error {
        let sheet_row = self.batch.sheet_row(row).unwrap_or_default();
        let sheet_col = self.plan.range.beg.col.wrapping_add(column as u64);
        let cell = CellPos::new(sheet_row, sheet_col);
        invalid_input(format!(
            "read_xlsx: Failed to parse cell '{cell}': {reason}"
        ))
    }

    fn write_serials<T: WritableNativeElement<Write = T>>(
        &self,
        doubles: &DataChunk<'_>,
        output: Option<&mut DataChunkRef<'_>>,
        column: usize,
        type_name: &str,
        convert: impl Fn(f64) -> std::result::Result<T, SerialError>,
    ) -> Result<()> {
        let serials = doubles.get_vector_at::<f64>(0)?;
        let rows = serials.len();
        let serials = serials.reader()?;
        let mut vector = match output {
            Some(output) => Some(output.get_vector_at_mut::<T>(column)?),
            None => None,
        };
        if let Some(vector) = vector.as_mut() {
            vector.set_size(rows)?;
        }
        let mut writer = vector.as_mut().map(Vector::writer).transpose()?;
        for row in 0..rows {
            let value = match serials.get(row)?.map(|serial| (*serial, convert(*serial))) {
                None => None,
                Some((_, Ok(value))) => Some(value),
                Some((_, Err(_))) if self.plan.ignore_errors => None,
                Some((serial, Err(reason))) => {
                    let reason = format!(
                        "Could not convert the serial number {serial:?} to {type_name}: {reason}"
                    );
                    return Err(self.cell_error(row, column, reason));
                }
            };
            if let Some(writer) = writer.as_mut() {
                writer.set_option(row, value.as_ref())?;
            }
        }
        Ok(())
    }
}
