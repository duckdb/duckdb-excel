//! The `read_xlsx` table function.

use duckdb_neo::{
    DuckDBType, Result,
    bind_arguments::BindArgument,
    connection::Context,
    data_chunk::DataChunkRef,
    file::FileSystem,
    logical_type::{LogicalType, LogicalTypeID},
    signature::{FunctionParameter, FunctionParameterKind, SignatureBuilder},
    table_function::{
        BindFunctionHandle, ExecColumnInfo, InitColumnData, TableFunctionBuilder,
        TableFunctionCallbacks,
    },
    types::{Any, DateValue, TimeValue, TimestampValue},
    value::Value,
};

use crate::{
    duckdb_compat::deduplicate_columns,
    error::{binder, invalid_input, raw_message},
    file::first_match,
    xlsx::{
        archive::Archive,
        names::{normalize_column_name, trim_whitespace},
        options::{NamedOptions, ReadOptions},
        resolve::{ColumnType, resolve_sheet},
        scan::{ScanPlan, ScanProgress, SheetScan},
    },
};

pub const FUNCTION_NAME: &str = "read_xlsx";

/// Where a named parameter of `read_xlsx` goes once it has the type of its option.
enum Setter {
    Boolean(fn(&mut NamedOptions) -> &mut Option<bool>),
    Varchar(fn(&mut NamedOptions) -> &mut Option<String>),
}

/// The named parameters in the order they are declared, which is the order `bind` gets them in.
const NAMED_PARAMETERS: [(&str, Setter); 8] = [
    ("header", Setter::Boolean(|options| &mut options.header)),
    (
        "all_varchar",
        Setter::Boolean(|options| &mut options.all_varchar),
    ),
    (
        "ignore_errors",
        Setter::Boolean(|options| &mut options.ignore_errors),
    ),
    ("range", Setter::Varchar(|options| &mut options.range)),
    ("sheet", Setter::Varchar(|options| &mut options.sheet)),
    (
        "stop_at_empty",
        Setter::Boolean(|options| &mut options.stop_at_empty),
    ),
    (
        "empty_as_varchar",
        Setter::Boolean(|options| &mut options.empty_as_varchar),
    ),
    (
        "normalize_names",
        Setter::Boolean(|options| &mut options.normalize_names),
    ),
];

fn sql_literal(value: &Value) -> Result<String> {
    let text = value.dbg_string()?;
    let logical_type = value.logical_type()?;
    Ok(match logical_type.id()? {
        LogicalTypeID::Varchar | LogicalTypeID::Enum => format!("'{}'", text.replace('\'', "''")),
        LogicalTypeID::Uuid
        | LogicalTypeID::Date
        | LogicalTypeID::Time
        | LogicalTypeID::TimeNs
        | LogicalTypeID::TimeTz
        | LogicalTypeID::Timestamp
        | LogicalTypeID::TimestampSec
        | LogicalTypeID::TimestampMs
        | LogicalTypeID::TimestampNs
        | LogicalTypeID::TimestampTz
        | LogicalTypeID::TimestampTzNs
        | LogicalTypeID::Interval
        | LogicalTypeID::Blob
        | LogicalTypeID::Bit
        | LogicalTypeID::Geometry => format!("'{text}'::{logical_type}"),
        _ => text,
    })
}

fn cast_parameter(
    context: &Context,
    name: &str,
    value: &Value,
    target: &LogicalType,
) -> Result<Value> {
    value.cast(context, target).map_err(|error| {
        let reason = raw_message(&error);
        let reason = reason.strip_prefix("Failed to cast value: ").unwrap_or(reason);
        let literal = sql_literal(value).unwrap_or_default();
        invalid_input(format!(
            "Could not cast value {literal} to named parameter \"{name}\" of type {target}: {reason}"
        ))
    })
}

fn named_options(
    context: &Context,
    arguments: impl Iterator<Item = BindArgument>,
) -> Result<NamedOptions> {
    let boolean = bool::logical_type(context)?;
    let varchar = String::logical_type(context)?;
    let mut options = NamedOptions::default();
    for ((name, setter), argument) in NAMED_PARAMETERS.iter().zip(arguments) {
        let value = argument.value.ok_or_else(|| {
            binder(format!(
                "{FUNCTION_NAME}: named parameter \"{name}\" must be a constant"
            ))
        })?;
        if value.is_null()? {
            continue;
        }
        match setter {
            Setter::Boolean(field) => {
                *field(&mut options) = cast_parameter(context, name, &value, &boolean)?.get()?;
            }
            Setter::Varchar(field) => {
                *field(&mut options) = cast_parameter(context, name, &value, &varchar)?.get()?;
            }
        }
    }
    Ok(options)
}

fn logical_type(context: &Context, column_type: ColumnType) -> Result<LogicalType> {
    match column_type {
        ColumnType::Varchar => String::logical_type(context),
        ColumnType::Boolean => bool::logical_type(context),
        ColumnType::Double => f64::logical_type(context),
        ColumnType::Date => DateValue::logical_type(context),
        ColumnType::Time => TimeValue::logical_type(context),
        ColumnType::Timestamp => TimestampValue::logical_type(context),
    }
}

/// Reads one sheet of an XLSX file; the scan is single-threaded and returns rows in sheet order.
///
/// The C API has no optional parameter without a default, so the named parameters default to NULL and a NULL
/// is an option that was not given. They accept any type and are cast to the type of their option when the
/// function is bound, the way an explicit cast does.
#[derive(Debug)]
pub struct ReadXlsx;

impl TableFunctionCallbacks for ReadXlsx {
    type BindData = ScanPlan;
    type LocalState = SheetScan;
    type GlobalState = ScanProgress;

    fn bind(
        &self,
        context: &Context,
        arguments: Vec<BindArgument>,
        bind_handle: BindFunctionHandle<'_>,
    ) -> Result<ScanPlan> {
        let mut arguments = arguments.into_iter();
        let pattern = arguments
            .next()
            .and_then(|argument| argument.value)
            .map(|value| value.get::<String>())
            .transpose()?
            .flatten()
            .ok_or_else(|| binder(format!("{FUNCTION_NAME}: the file path cannot be NULL")))?;
        let named = named_options(context, arguments)?;

        // Globbing also loads the extension that serves the file system of the path.
        let file_system = FileSystem::new(context)?;
        let file_path = first_match(&file_system, &pattern)?;
        let mut archive = Archive::open(&file_system, &file_path)?;
        let options = ReadOptions::new(&named)?;
        let layout = resolve_sheet(&mut archive, &file_path, &options)?;

        let mut names: Vec<String> = layout
            .columns
            .iter()
            .map(|column| {
                if options.normalize_names {
                    normalize_column_name(&column.name)
                } else {
                    trim_whitespace(&column.name).to_owned()
                }
            })
            .collect();
        deduplicate_columns(&mut names);
        let types = layout
            .columns
            .iter()
            .map(|column| logical_type(context, column.column_type))
            .collect::<Result<Vec<_>>>()?;
        for (name, logical_type) in names.iter().zip(&types) {
            bind_handle.add_result_column(name, logical_type)?;
        }
        ScanPlan::new(file_path, layout, &options, types)
    }

    fn init_global_state(
        &self,
        _context: &Context,
        _bind_data: &ScanPlan,
        _column_data: InitColumnData<'_>,
    ) -> Result<Option<ScanProgress>> {
        Ok(Some(ScanProgress::default()))
    }

    fn init_local_state(
        &self,
        context: &Context,
        bind_data: &ScanPlan,
        global_state: Option<&ScanProgress>,
        _column_data: InitColumnData<'_>,
    ) -> Result<Option<SheetScan>> {
        SheetScan::open(context, bind_data, global_state)
    }

    fn exec(
        &self,
        context: &Context,
        bind_data: &ScanPlan,
        local_state: Option<&mut SheetScan>,
        global_state: Option<&ScanProgress>,
        output: DataChunkRef<'_>,
        _column_info: ExecColumnInfo<'_>,
    ) -> Result<()> {
        SheetScan::exec(context, bind_data, local_state, global_state, output)
    }

    fn progress(
        &self,
        _context: &Context,
        _bind_data: &ScanPlan,
        global_state: Option<&ScanProgress>,
    ) -> Result<Option<f64>> {
        Ok(global_state.map(ScanProgress::fraction))
    }
}

/// `read_xlsx(file_path, header, all_varchar, ignore_errors, range, sheet, stop_at_empty, empty_as_varchar,
/// normalize_names)`; everything but the path is optional and passed by name.
pub fn table_function() -> TableFunctionBuilder<ReadXlsx> {
    let mut signature =
        SignatureBuilder::new().param(FunctionParameter::new::<String>("file_path"));
    for (name, setter) in &NAMED_PARAMETERS {
        let parameter = FunctionParameter::new::<Any>(*name).kind(FunctionParameterKind::Named);
        signature = signature.param(match setter {
            Setter::Boolean(_) => parameter.default_value(None::<bool>),
            Setter::Varchar(_) => parameter.default_value(None::<String>),
        });
    }
    TableFunctionBuilder::new(FUNCTION_NAME, signature, ReadXlsx).with_progress()
}
