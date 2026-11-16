//! The scan behind `read_xlsx` and `COPY ... FROM`: reads the rows of a sheet as text and lets DuckDB cast them
//! to the column types.

use std::sync::atomic::{AtomicU64, Ordering};

use duckdb_neo::{
    DuckDBType, Result,
    connection::Context,
    data_chunk::{DataChunk, DataChunkRef},
    file::FileSystem,
    logical_type::{LogicalType, LogicalTypeID},
    types::{DateValue, TimeValue, TimestampValue},
    vector::{Unknown, Vector},
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
        serial::{epoch_micros_to_date, epoch_micros_to_time, serial_to_epoch_micros},
        strings::StringTable,
    },
};

/// How the text of the cells of a column becomes values of the column type: a text column takes it as it is,
/// the numbers of a date or time column are serial numbers, and DuckDB casts everything else.
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
///
/// A cell that does not convert is reported with the message of the first cast that failed in the scan, which
/// may have been the cast of another cell.
#[derive(Debug)]
pub struct SheetScan {
    rows: RowReader<Archive>,
    strings: StringTable,
    cast_error: Option<String>,
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
        Ok(Some(SheetScan {
            rows,
            strings,
            cast_error: None,
        }))
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
            cast_error: &mut scan.cast_error,
        };
        emitter.emit(output)
    }
}

fn missing_state() -> duckdb_neo::Error {
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

/// Turns a batch of rows into the vectors of an output chunk.
///
/// A cell that does not convert becomes NULL when it is empty or errors are ignored, and fails the scan
/// otherwise. Columns the output leaves out are converted all the same, so that their cells are checked.
struct Emitter<'a> {
    context: &'a Context,
    plan: &'a ScanPlan,
    batch: &'a RowBatch,
    strings: &'a StringTable,
    cast_error: &'a mut Option<String>,
}

impl Emitter<'_> {
    fn emit(mut self, mut output: DataChunkRef<'_>) -> Result<()> {
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
        let casts = columns
            .iter()
            .filter(|column| column.conversion != Conversion::Text)
            .count();
        if casts == 0 {
            for index in 0..present {
                self.write_text(output.get_vector_at_mut(index)?, index)?;
            }
            return Ok(());
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

        let mut next_text = 0;
        let mut next_unseen = 0;
        for (index, column) in columns.iter().enumerate() {
            let is_present = index < present;
            if column.conversion == Conversion::Text {
                if is_present {
                    self.write_text(output.get_vector_at_mut(index)?, index)?;
                }
                continue;
            }
            let text_index = next_text;
            next_text += 1;
            self.write_text(texts.get_vector_at_mut(text_index)?, index)?;
            let text = texts.get_vector_at::<String>(text_index)?;
            match column.conversion {
                Conversion::Cast if is_present => self.cast(text, &mut output, index, index)?,
                Conversion::Cast => {
                    if let Some(unseen) = unseen.as_mut() {
                        self.cast(text, unseen, next_unseen, index)?;
                        next_unseen += 1;
                    }
                }
                conversion => {
                    self.cast(text, &mut doubles, 0, index)?;
                    let serials = doubles.get_vector_at::<f64>(0)?;
                    let target = is_present.then_some(&mut output);
                    write_serials(serials, conversion, target, index)?;
                }
            }
        }
        Ok(())
    }

    fn write_text(&self, mut vector: Vector<'_, String>, column: usize) -> Result<()> {
        let rows = self.batch.len();
        vector.set_size(rows)?;
        let mut writer = vector.writer()?;
        for row in 0..rows {
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
        &mut self,
        text: &Vector<'_, String>,
        target: &mut impl CastTarget,
        index: usize,
        column: usize,
    ) -> Result<()> {
        // SAFETY: a text vector is filled before its cast and not written again, and the target is only read
        // once the cast is done.
        let failure = unsafe { target.target(index)?.try_cast_rows(text, self.context)? };
        let Some(failure) = failure else {
            return Ok(());
        };
        let message = self
            .cast_error
            .get_or_insert_with(|| raw_message(&failure).to_owned());
        if self.plan.ignore_errors {
            return Ok(());
        }
        let values = target.result(index)?.reader()?;
        for row in 0..self.batch.len() {
            let text = match self.batch.cell(row, column) {
                CellValue::Null => continue,
                CellValue::Text(text) => text,
                CellValue::Shared(index) => self.strings.get(index).unwrap_or_default(),
            };
            if !text.is_empty() && values.is_null(row)? {
                let sheet_row = self.batch.sheet_row(row).unwrap_or_default();
                let sheet_col = self.plan.range.beg.col.wrapping_add(column as u64);
                let cell = CellPos::new(sheet_row, sheet_col);
                return Err(invalid_input(format!(
                    "read_xlsx: Failed to parse cell '{cell}': {message}"
                )));
            }
        }
        Ok(())
    }
}

fn write_serials(
    serials: &Vector<'_, f64>,
    conversion: Conversion,
    output: Option<&mut DataChunkRef<'_>>,
    index: usize,
) -> Result<()> {
    let rows = serials.len();
    let serials = serials.reader()?;
    let micros = |row: usize| -> Result<Option<i64>> {
        Ok(serials
            .get(row)?
            .map(|serial| serial_to_epoch_micros(*serial)))
    };
    match (conversion, output) {
        (Conversion::SerialTime, None) => {
            for row in 0..rows {
                micros(row)?.map(epoch_micros_to_time).transpose()?;
            }
        }
        (Conversion::SerialTime, Some(output)) => {
            let mut vector = output.get_vector_at_mut::<TimeValue>(index)?;
            vector.set_size(rows)?;
            let mut writer = vector.writer()?;
            for row in 0..rows {
                let time = micros(row)?.map(epoch_micros_to_time).transpose()?;
                writer.set_option(row, time.map(TimeValue).as_ref())?;
            }
        }
        (Conversion::SerialDate, Some(output)) => {
            let mut vector = output.get_vector_at_mut::<DateValue>(index)?;
            vector.set_size(rows)?;
            let mut writer = vector.writer()?;
            for row in 0..rows {
                let date = micros(row)?.map(epoch_micros_to_date);
                writer.set_option(row, date.map(DateValue).as_ref())?;
            }
        }
        (Conversion::SerialTimestamp, Some(output)) => {
            let mut vector = output.get_vector_at_mut::<TimestampValue>(index)?;
            vector.set_size(rows)?;
            let mut writer = vector.writer()?;
            for row in 0..rows {
                writer.set_option(row, micros(row)?.map(TimestampValue).as_ref())?;
            }
        }
        _ => {}
    }
    Ok(())
}
