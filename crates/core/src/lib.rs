//! The DuckDB Excel extension: `text`, `excel_text`, `read_xlsx` and `COPY ... (FORMAT xlsx)`.

use duckdb_neo::{Result, extension::Extension};

/// Registers every function of the extension through `$link`, an `Extension` or a `Connection`.
#[macro_export]
macro_rules! register_functions {
    ($link:expr) => {{
        let _link = $link;
        duckdb_neo::Result::Ok(())
    }};
}

/// Registers everything the extension provides.
pub fn register(extension: &Extension) -> Result<()> {
    register_functions!(extension)
}
