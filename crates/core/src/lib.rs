//! The DuckDB Excel extension: `text`, `excel_text`, `read_xlsx` and `COPY ... (FORMAT xlsx)`.

use duckdb_neo::{Result, extension::Extension};

pub mod duckdb_compat;
pub mod error;
pub mod file;
pub mod text;
pub mod xlsx;

/// Registers every function of the extension through `$link`, an `Extension` or a `Connection`.
#[macro_export]
macro_rules! register_functions {
    ($link:expr) => {{
        let link = $link;
        $crate::text::text_function(link, "text")?.register(link)?;
        $crate::text::text_function(link, "excel_text")?.register(link)?;
        $crate::xlsx::read::table_function().register(link)?;
        $crate::xlsx::replacement::replacement_scan().register(link)?;
        $crate::xlsx::copy::copy_function().register_to_and_from(link)?;
        duckdb_neo::Result::Ok(())
    }};
}

/// Registers everything the extension provides.
pub fn register(extension: &Extension) -> Result<()> {
    register_functions!(extension)
}
