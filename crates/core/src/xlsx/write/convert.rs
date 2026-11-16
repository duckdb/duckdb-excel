//! Turns the values of a chunk into the text of their cells.
//!
//! DuckDB casts every value to text. Booleans become integers first, and dates, times and timestamps the
//! serial numbers spreadsheets count time in.

use duckdb_neo::{
    DuckDBType, Result,
    connection::Context,
    data_chunk::{DataChunk, DataChunkRef},
    logical_type::{LogicalType, LogicalTypeID},
    types::{DateValue, TimeValue, TimestampValue},
    vector::{NativeElement, Unknown, Vector, VectorReader},
};

use crate::{
    error::{internal, reraise},
    xlsx::{
        serial::{epoch_days_to_serial, epoch_micros_to_serial, time_micros_to_serial},
        write::sheet::CellKind,
    },
};

/// How the values of a column become text; the numbers are the scratch vectors the conversion goes through.
///
/// For most types the cast to text is all it takes. A time with a zone offset is cast to a time, which drops
/// the offset, and a timestamp of another precision or with a time zone to a plain timestamp.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Conversion {
    Cast,
    Boolean { integer: usize },
    Date { serial: usize },
    Time { serial: usize },
    ZonedTime { time: usize, serial: usize },
    Timestamp { serial: usize },
    OtherTimestamp { timestamp: usize, serial: usize },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum ScratchType {
    Integer,
    Time,
    Timestamp,
    Serial,
}

/// The conversions of the columns of a sheet, decided when the statement is bound.
#[derive(Debug, Default)]
pub struct ConversionPlan {
    conversions: Vec<Conversion>,
    scratch: Vec<ScratchType>,
}

impl ConversionPlan {
    /// Adds a column of the type `type_id` and returns the kind of cell its values are written as.
    pub fn add_column(&mut self, type_id: LogicalTypeID) -> CellKind {
        use LogicalTypeID as Id;
        let mut scratch = |scratch_type: ScratchType| {
            self.scratch.push(scratch_type);
            self.scratch.len() - 1
        };
        let (conversion, kind) = match type_id {
            Id::Boolean => {
                let integer = scratch(ScratchType::Integer);
                (Conversion::Boolean { integer }, CellKind::Boolean)
            }
            Id::Date => {
                let serial = scratch(ScratchType::Serial);
                (Conversion::Date { serial }, CellKind::Date)
            }
            Id::Time => {
                let serial = scratch(ScratchType::Serial);
                (Conversion::Time { serial }, CellKind::Time)
            }
            Id::TimeTz => {
                let time = scratch(ScratchType::Time);
                let serial = scratch(ScratchType::Serial);
                (Conversion::ZonedTime { time, serial }, CellKind::Time)
            }
            Id::Timestamp => {
                let serial = scratch(ScratchType::Serial);
                (Conversion::Timestamp { serial }, CellKind::Timestamp)
            }
            Id::TimestampTz | Id::TimestampMs | Id::TimestampNs | Id::TimestampSec => {
                let timestamp = scratch(ScratchType::Timestamp);
                let serial = scratch(ScratchType::Serial);
                let kind = if type_id == Id::TimestampSec {
                    CellKind::TimestampSeconds
                } else {
                    CellKind::Timestamp
                };
                (Conversion::OtherTimestamp { timestamp, serial }, kind)
            }
            Id::TinyInt
            | Id::SmallInt
            | Id::Integer
            | Id::BigInt
            | Id::HugeInt
            | Id::UTinyInt
            | Id::USmallInt
            | Id::UInteger
            | Id::UBigInt
            | Id::UHugeInt
            | Id::Float
            | Id::Double
            | Id::Decimal => (Conversion::Cast, CellKind::Number),
            _ => (Conversion::Cast, CellKind::Text),
        };
        self.conversions.push(conversion);
        kind
    }

    pub fn column_count(&self) -> usize {
        self.conversions.len()
    }
}

/// Converts the chunks of one batch, which must have a column at least.
///
/// Every vector it casts into is written once for a chunk and read only after that, and all of them are reset
/// before the next chunk.
pub struct TextConverter<'a> {
    context: &'a Context,
    plan: &'a ConversionPlan,
    texts: DataChunk<'a>,
    scratch: Option<DataChunk<'a>>,
}

impl<'a> TextConverter<'a> {
    pub fn new(context: &'a Context, plan: &'a ConversionPlan) -> Result<TextConverter<'a>> {
        let text = String::logical_type(context)?;
        let texts = DataChunk::create(context, vec![text; plan.conversions.len()], true)?;
        let scratch_types = plan
            .scratch
            .iter()
            .map(|scratch_type| match scratch_type {
                ScratchType::Integer => i32::logical_type(context),
                ScratchType::Time => TimeValue::logical_type(context),
                ScratchType::Timestamp => TimestampValue::logical_type(context),
                ScratchType::Serial => f64::logical_type(context),
            })
            .collect::<Result<Vec<LogicalType>>>()?;
        let scratch = if scratch_types.is_empty() {
            None
        } else {
            Some(DataChunk::create(context, scratch_types, true)?)
        };
        Ok(TextConverter {
            context,
            plan,
            texts,
            scratch,
        })
    }

    /// Converts the columns of `chunk`; [`Self::texts`] has the result until the next chunk is converted.
    pub fn convert(&mut self, chunk: &DataChunkRef<'_>) -> Result<()> {
        self.texts.clear()?;
        if let Some(scratch) = self.scratch.as_mut() {
            scratch.clear()?;
        }
        for (column, conversion) in self.plan.conversions.iter().enumerate() {
            let source = chunk.get_vector_at::<Unknown>(column)?;
            let last = match *conversion {
                Conversion::Cast => {
                    cast(self.context, source, self.texts.get_vector_at_mut(column)?)?;
                    continue;
                }
                Conversion::Boolean { integer } => {
                    cast(
                        self.context,
                        source,
                        self.scratch()?.get_vector_at_mut(integer)?,
                    )?;
                    integer
                }
                Conversion::Date { serial } => {
                    let dates = source.cast_ref::<DateValue>()?.reader()?;
                    let serials = self.scratch()?.get_vector_at_mut(serial)?;
                    write_serials(serials, dates.len(), |row| {
                        Ok(dates.get(row)?.map(|date| epoch_days_to_serial(date.0)))
                    })?;
                    serial
                }
                Conversion::Time { serial } => {
                    let times = source.cast_ref::<TimeValue>()?.reader()?;
                    let serials = self.scratch()?.get_vector_at_mut(serial)?;
                    write_serials(serials, times.len(), |row| {
                        Ok(times.get(row)?.map(|time| time_micros_to_serial(time.0)))
                    })?;
                    serial
                }
                Conversion::ZonedTime { time, serial } => {
                    cast(
                        self.context,
                        source,
                        self.scratch()?.get_vector_at_mut(time)?,
                    )?;
                    let [times, serials] = self.scratch()?.get_vector_array_mut([time, serial])?;
                    let times = times.cast::<TimeValue>()?;
                    let times = times.reader()?;
                    write_serials(serials.cast()?, times.len(), |row| {
                        Ok(times.get(row)?.map(|time| time_micros_to_serial(time.0)))
                    })?;
                    serial
                }
                Conversion::Timestamp { serial } => {
                    let timestamps = source.cast_ref::<TimestampValue>()?.reader()?;
                    let serials = self.scratch()?.get_vector_at_mut(serial)?;
                    write_serials(serials, timestamps.len(), |row| {
                        let timestamp = timestamps.get(row)?;
                        Ok(timestamp.map(|timestamp| epoch_micros_to_serial(timestamp.0)))
                    })?;
                    serial
                }
                Conversion::OtherTimestamp { timestamp, serial } => {
                    cast(
                        self.context,
                        source,
                        self.scratch()?.get_vector_at_mut(timestamp)?,
                    )?;
                    let [timestamps, serials] =
                        self.scratch()?.get_vector_array_mut([timestamp, serial])?;
                    let timestamps = timestamps.cast::<TimestampValue>()?;
                    let timestamps = timestamps.reader()?;
                    write_serials(serials.cast()?, timestamps.len(), |row| {
                        let timestamp = timestamps.get(row)?;
                        Ok(timestamp.map(|timestamp| epoch_micros_to_serial(timestamp.0)))
                    })?;
                    serial
                }
            };
            let text = self.texts.get_vector_at_mut(column)?;
            let Some(scratch) = self.scratch.as_ref() else {
                return Err(no_scratch());
            };
            cast(self.context, scratch.get_vector_at::<Unknown>(last)?, text)?;
        }
        Ok(())
    }

    /// The text of the cells of the chunk converted last, a reader for each column.
    pub fn texts(&self) -> Result<Vec<VectorReader<'_, String>>> {
        (0..self.plan.conversions.len())
            .map(|column| self.texts.get_vector_at::<String>(column)?.reader())
            .collect()
    }

    fn scratch(&mut self) -> Result<&mut DataChunk<'a>> {
        self.scratch.as_mut().ok_or_else(no_scratch)
    }
}

fn no_scratch() -> duckdb_neo::Error {
    internal("COPY TO xlsx: a conversion has no scratch vector")
}

fn cast<T: NativeElement>(
    context: &Context,
    source: &Vector<'_, T>,
    target: Vector<'_, Unknown>,
) -> Result<()> {
    // SAFETY: the source is a vector of the chunk being converted, which nothing writes to, or a scratch vector
    // that was filled before; the target is read once the cast is done and not written until both are reset.
    unsafe { target.cast_rows(source, context) }.map_err(reraise)
}

fn write_serials(
    mut serials: Vector<'_, f64>,
    rows: usize,
    serial_of_row: impl Fn(usize) -> Result<Option<f64>>,
) -> Result<()> {
    serials.set_size(rows)?;
    let mut writer = serials.writer()?;
    for row in 0..rows {
        writer.set_option(row, serial_of_row(row)?.as_ref())?;
    }
    Ok(())
}
