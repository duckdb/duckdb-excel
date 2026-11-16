//! `SELECT * FROM 'file.xlsx'`: table names that end in `.xlsx` are read with `read_xlsx`.

use duckdb_neo::{
    Result, ToValue,
    connection::Context,
    qualified_name::QualifiedName,
    replacement_scan::{
        ReplacementHandle, ReplacementScanBuilder, ReplacementScanCallbacks, ReplacementType,
    },
};

use crate::xlsx::read;

/// Replaces a table name that DuckDB cannot find and that ends in `.xlsx`, in any letter case, with a call of
/// `read_xlsx`. The path is the whole name: the parts of a qualified name joined with dots.
#[derive(Debug)]
pub struct XlsxReplacementScan;

impl ReplacementScanCallbacks for XlsxReplacementScan {
    fn scan(
        &self,
        context: &Context,
        name: &QualifiedName,
        handle: ReplacementHandle<'_>,
    ) -> Result<()> {
        let path = name.parts()?.join(".");
        if !path.to_ascii_lowercase().ends_with(".xlsx") {
            return Ok(());
        }
        handle.set_reference(ReplacementType::Table(QualifiedName::from_parts(&[
            read::FUNCTION_NAME,
        ])?))?;
        handle.add_parameter(path.value(context)?)
    }
}

pub fn replacement_scan() -> ReplacementScanBuilder<'static, XlsxReplacementScan> {
    ReplacementScanBuilder::new(XlsxReplacementScan)
}
