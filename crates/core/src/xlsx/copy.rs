//! The `xlsx` format of `COPY`: `COPY tbl FROM 'file.xlsx' (FORMAT xlsx)` fills a table from a sheet, and
//! `COPY ... TO 'file.xlsx' (FORMAT xlsx)` writes rows into a sheet.

use duckdb_neo::{
    DuckDBType, Result,
    column_data_collection::ColumnDataCollection,
    connection::Context,
    copy_function::{
        CopyFromBindInfo, CopyFromFunctionCallbacks, CopyFunctionBuilder, CopyToBindInfo,
        CopyToFunctionCallbacks,
    },
    data_chunk::DataChunkRef,
    file::FileSystem,
    logical_type::LogicalTypeID,
    value::Value,
};

use crate::{
    error::{binder, conversion, internal, reraise},
    xlsx::{
        archive::Archive,
        options::{NamedOptions, ReadOptions},
        resolve::resolve_sheet,
        scan::{ScanPlan, ScanProgress, SheetScan},
        write::export::{WorkbookOutput, WritePlan},
    },
};

pub const FORMAT_NAME: &str = "xlsx";

/// The copy function of the `xlsx` format.
///
/// `COPY ... FROM` runs the scan of `read_xlsx` with the column types of the table. Its options are those of
/// `read_xlsx` but for `normalize_names`: a flag given without a value is set, and options the format does not
/// know are ignored.
///
/// `COPY ... TO` writes one sheet. `HEADER` puts the column names into the first row, `SHEET` names the sheet,
/// `SHEET_ROW_LIMIT` is the number of rows it may have, and `MODE 'append'` or `MODE 'replace'` add the sheet
/// to a workbook that exists, where `MODE 'create'` writes the workbook anew.
///
/// DuckDB may prepare the batches of a `COPY ... TO` on several threads at once and for a file other than the
/// one they end up in, but it flushes the batches of a file one after the other and in the order of their
/// rows. All the work is therefore done when a batch is flushed.
#[derive(Debug)]
pub struct XlsxCopy;

fn quoted(name: &str) -> String {
    format!("\"{}\"", name.replace('"', "\"\""))
}

fn boolean_option(context: &Context, name: &str, value: &Value) -> Result<bool> {
    let name = quoted(name);
    let not_boolean = || {
        binder(format!(
            "'{name}' option must be standalone or a BOOLEAN value"
        ))
    };
    // An option with several values arrives as one tuple.
    let type_id = value.logical_type()?.id()?;
    if type_id == LogicalTypeID::Tuple || value.is_null()? {
        return Err(not_boolean());
    }
    if matches!(
        type_id,
        LogicalTypeID::Float | LogicalTypeID::Double | LogicalTypeID::Decimal
    ) {
        return Err(binder(format!(
            "'{name}' options expects a boolean value (e.g. TRUE or 1)"
        )));
    }
    value
        .cast(context, &bool::logical_type(context)?)
        .map_err(reraise)?
        .get()?
        .ok_or_else(not_boolean)
}

fn varchar_option(name: &str, value: &Value) -> Result<String> {
    let not_varchar = || {
        binder(format!(
            "'{}' option must be a single VARCHAR value",
            quoted(name)
        ))
    };
    if value.logical_type()?.id()? != LogicalTypeID::Varchar {
        return Err(not_varchar());
    }
    value.get()?.ok_or_else(not_varchar)
}

fn copy_from_options(context: &Context, bind_info: &CopyFromBindInfo) -> Result<NamedOptions> {
    let mut options = NamedOptions::default();
    for index in 0..bind_info.option_count()? {
        let (name, value) = bind_info.get_option(index)?;
        let flag = match name.to_ascii_lowercase().as_str() {
            "sheet" => {
                options.sheet = Some(varchar_option(name, &value)?);
                continue;
            }
            "range" => {
                options.range = Some(varchar_option(name, &value)?);
                continue;
            }
            "header" => &mut options.header,
            "all_varchar" => &mut options.all_varchar,
            "ignore_errors" => &mut options.ignore_errors,
            "stop_at_empty" => &mut options.stop_at_empty,
            "empty_as_varchar" => &mut options.empty_as_varchar,
            _ => continue,
        };
        *flag = Some(boolean_option(context, name, &value)?);
    }
    Ok(options)
}

impl CopyFromFunctionCallbacks for XlsxCopy {
    type BindData = ScanPlan;
    type LocalState = SheetScan;
    type GlobalState = ScanProgress;

    fn bind(
        &self,
        context: &Context,
        file_path: &str,
        bind_info: &CopyFromBindInfo,
    ) -> Result<ScanPlan> {
        let options = ReadOptions::new(&copy_from_options(context, bind_info)?)?;
        let mut archive = Archive::open(&FileSystem::new(context)?, file_path)?;
        let layout = resolve_sheet(&mut archive, file_path, &options)?;

        let expected = (0..bind_info.column_count()?)
            .map(|index| {
                let (name, logical_type) = bind_info.get_column(index)?;
                Ok((name.to_owned(), logical_type))
            })
            .collect::<Result<Vec<_>>>()?;
        if expected.len() != layout.columns.len() {
            let table_schema: Vec<String> = expected
                .iter()
                .map(|(name, logical_type)| format!("{name} {logical_type}"))
                .collect();
            let xlsx_schema: Vec<String> = layout
                .columns
                .iter()
                .map(|column| format!("{} {}", column.name, column.column_type.name()))
                .collect();
            return Err(conversion(format!(
                "Failed to read file(s) \"{file_path}\" - column count mismatch: expected {} columns but found {}\n\
                 Table schema: {}\n\
                 XLSX schema: {}\n\n\
                 Possible solutions:\n\
                 * Manually specify which columns to insert using \"INSERT INTO tbl SELECT ... FROM read_xlsx(...)\"\n\
                 * Provide an explicit range option with the same width as the table schema using e.g.\
                 \"COPY tbl FROM ... (FORMAT 'xlsx', range 'A1:Z10')\"",
                expected.len(),
                layout.columns.len(),
                table_schema.join(", "),
                xlsx_schema.join(", "),
            )));
        }
        let types = expected
            .into_iter()
            .map(|(_, logical_type)| logical_type)
            .collect();
        ScanPlan::new(file_path.to_owned(), layout, &options, types)
    }

    fn init_global_state(
        &self,
        _context: &Context,
        _bind_data: &ScanPlan,
    ) -> Result<Option<ScanProgress>> {
        Ok(Some(ScanProgress::default()))
    }

    fn init_local_state(
        &self,
        context: &Context,
        bind_data: &ScanPlan,
        global_state: Option<&ScanProgress>,
    ) -> Result<Option<SheetScan>> {
        let progress = global_state.ok_or_else(missing_state)?;
        SheetScan::open(context, bind_data, progress).map(Some)
    }

    fn exec(
        &self,
        context: &Context,
        bind_data: &ScanPlan,
        local_state: Option<&mut SheetScan>,
        global_state: Option<&ScanProgress>,
        output: DataChunkRef<'_>,
    ) -> Result<()> {
        let (Some(scan), Some(progress)) = (local_state, global_state) else {
            return Err(missing_state());
        };
        scan.exec(context, bind_data, progress, output)
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

impl CopyToFunctionCallbacks for XlsxCopy {
    type InitData = WorkbookOutput;
    type BindData = WritePlan;
    type BatchData = ();

    fn bind(&self, context: &Context, bind_info: &CopyToBindInfo) -> Result<WritePlan> {
        WritePlan::new(context, bind_info)
    }

    fn init(
        &self,
        context: &Context,
        bind_data: &WritePlan,
        file_path: &str,
    ) -> Result<WorkbookOutput> {
        WorkbookOutput::begin(context, bind_data, file_path)
    }

    fn batch(
        &self,
        _context: &Context,
        _bind_data: &WritePlan,
        _init_data: &WorkbookOutput,
        _input: &ColumnDataCollection<'_>,
    ) -> Result<()> {
        Ok(())
    }

    fn flush(
        &self,
        context: &Context,
        bind_data: &WritePlan,
        init_data: &WorkbookOutput,
        input: &ColumnDataCollection<'_>,
        _batch_data: &mut (),
    ) -> Result<()> {
        init_data.write_rows(context, bind_data, input)
    }

    fn finalize(
        &self,
        context: &Context,
        bind_data: &WritePlan,
        init_data: &WorkbookOutput,
    ) -> Result<()> {
        init_data.finish(context, bind_data)
    }
}

fn missing_state() -> duckdb_neo::Error {
    internal("COPY FROM xlsx: the scan state is missing")
}

/// The `xlsx` copy function; files DuckDB names itself get the extension `.xlsx`.
pub fn copy_function() -> CopyFunctionBuilder<XlsxCopy> {
    CopyFunctionBuilder::new(FORMAT_NAME, XlsxCopy)
        .with_fallback_extension(FORMAT_NAME)
        .with_progress()
}
