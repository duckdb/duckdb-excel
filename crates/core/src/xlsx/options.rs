//! The options of `read_xlsx` and `COPY ... FROM ... (FORMAT xlsx)`.

use duckdb_neo::Result;

use crate::{
    error::binder,
    xlsx::{
        cell::{CellRange, CellType, MAX_CELL_COLS},
        escape::escape_xml,
    },
};

/// Whether the first row of the range holds the column names.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum HeaderMode {
    Never,
    /// It does when all its cells are non-empty strings.
    Maybe,
    Force,
}

/// The options as the caller gave them; `None` is an option that was not given.
#[derive(Debug, Default, Clone, PartialEq, Eq)]
pub struct NamedOptions {
    pub header: Option<bool>,
    pub all_varchar: Option<bool>,
    pub ignore_errors: Option<bool>,
    pub range: Option<String>,
    pub sheet: Option<String>,
    pub stop_at_empty: Option<bool>,
    pub empty_as_varchar: Option<bool>,
    pub normalize_names: Option<bool>,
}

/// The options a sheet is read with, defaults filled in.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ReadOptions {
    /// The name of the sheet with its XML special characters escaped; empty for the first sheet.
    pub sheet: String,
    pub header_mode: HeaderMode,
    pub all_varchar: bool,
    pub ignore_errors: bool,
    pub stop_at_empty: bool,
    pub has_explicit_range: bool,
    pub normalize_names: bool,
    /// The type of the cells a row leaves out.
    pub default_cell_type: CellType,
    pub range: CellRange,
}

impl ReadOptions {
    pub fn new(named: &NamedOptions) -> Result<ReadOptions> {
        let mut options = ReadOptions {
            sheet: String::new(),
            header_mode: HeaderMode::Maybe,
            all_varchar: false,
            ignore_errors: false,
            stop_at_empty: true,
            has_explicit_range: false,
            normalize_names: false,
            default_cell_type: CellType::Number,
            range: CellRange::WHOLE_SHEET,
        };
        if let Some(sheet) = &named.sheet {
            options.sheet = escape_xml(sheet);
        }
        if let Some(header) = named.header {
            options.header_mode = if header {
                HeaderMode::Force
            } else {
                HeaderMode::Never
            };
        }
        options.all_varchar = named.all_varchar.unwrap_or(false);
        options.ignore_errors = named.ignore_errors.unwrap_or(false);
        options.normalize_names = named.normalize_names.unwrap_or(false);
        if let Some(text) = &named.range {
            let invalid = || binder(format!("Invalid range '{text}' specified"));
            let mut range = CellRange::parse(text).ok_or_else(invalid)?;
            // More rows than a sheet can have are fine, more columns are not: every column costs a vector.
            if !range.is_valid() || range.width() > MAX_CELL_COLS {
                return Err(invalid());
            }
            range.end.col = range.end.col.wrapping_add(1);
            range.end.row = range.end.row.wrapping_add(1);
            options.range = range;
            options.has_explicit_range = true;
            options.stop_at_empty = false;
        }
        if let Some(stop_at_empty) = named.stop_at_empty {
            options.stop_at_empty = stop_at_empty;
        }
        if named.empty_as_varchar == Some(true) {
            options.default_cell_type = CellType::InlineString;
        }
        Ok(options)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xlsx::cell::{CellPos, MAX_CELL_ROWS};
    use duckdb_neo::error::DuckDBError;

    fn with_range(range: &str) -> Result<ReadOptions> {
        ReadOptions::new(&NamedOptions {
            range: Some(range.to_owned()),
            ..NamedOptions::default()
        })
    }

    #[test]
    fn defaults() {
        let options = ReadOptions::new(&NamedOptions::default()).unwrap();
        assert_eq!(options.header_mode, HeaderMode::Maybe);
        assert!(options.stop_at_empty && !options.has_explicit_range);
        assert!(!options.all_varchar && !options.ignore_errors && !options.normalize_names);
        assert_eq!(options.default_cell_type, CellType::Number);
        assert_eq!(options.range, CellRange::WHOLE_SHEET);
        assert_eq!(options.sheet, "");
    }

    #[test]
    fn header_is_forced_or_never() {
        for (header, mode) in [(true, HeaderMode::Force), (false, HeaderMode::Never)] {
            let options = ReadOptions::new(&NamedOptions {
                header: Some(header),
                ..NamedOptions::default()
            });
            assert_eq!(options.unwrap().header_mode, mode);
        }
    }

    #[test]
    fn a_range_includes_its_last_cell_and_reads_empty_rows() {
        let options = with_range("B2:D7").unwrap();
        assert_eq!(options.range.beg, CellPos::new(2, 2));
        assert_eq!(options.range.end, CellPos::new(8, 5));
        assert!(options.has_explicit_range && !options.stop_at_empty);

        let options = with_range("B:D").unwrap();
        assert_eq!(options.range.beg, CellPos::new(1, 2));
        assert_eq!(options.range.end, CellPos::new(MAX_CELL_ROWS + 1, 5));
    }

    #[test]
    fn stop_at_empty_overrides_the_default_of_a_range() {
        let options = ReadOptions::new(&NamedOptions {
            range: Some("A1:B2".to_owned()),
            stop_at_empty: Some(true),
            ..NamedOptions::default()
        });
        assert!(options.unwrap().stop_at_empty);
    }

    #[test]
    fn invalid_ranges() {
        for range in [
            "",
            "A1",
            "a1:b2",
            "Z1:A5",
            "A5:A1",
            "A1:ZZZZZZZZ5",
            "A1:XFF1",
        ] {
            let error = with_range(range).unwrap_err();
            assert_eq!(error.code, DuckDBError::QueryBinder, "{range}");
            assert_eq!(error.message, format!("Invalid range '{range}' specified"));
        }
        assert_eq!(with_range("A1:XFD1").unwrap().range.width(), MAX_CELL_COLS);
        assert_eq!(
            with_range("A1:XFE1").unwrap().range.width(),
            MAX_CELL_COLS + 1
        );
        assert_eq!(with_range("1:2").unwrap().range.width(), MAX_CELL_COLS);
    }

    #[test]
    fn empty_as_varchar_changes_the_default_cell_type() {
        let options = ReadOptions::new(&NamedOptions {
            empty_as_varchar: Some(true),
            ..NamedOptions::default()
        });
        assert_eq!(options.unwrap().default_cell_type, CellType::InlineString);
    }

    #[test]
    fn sheet_names_are_escaped() {
        let options = ReadOptions::new(&NamedOptions {
            sheet: Some("A&B <1>\0".to_owned()),
            ..NamedOptions::default()
        });
        assert_eq!(options.unwrap().sheet, "A&amp;B &lt;1&gt;");
    }
}
