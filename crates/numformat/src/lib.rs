//! Excel number formats, as the `TEXT` function of Microsoft Excel applies them.
//!
//! [`format_number`] formats one value; [`NumberFormat`] parses a format code once to apply it to
//! many values.
//!
//! Format codes are processed as bytes and never decoded, so text in any script passes through
//! unchanged. Only the en-US locale is available: `.` is the decimal separator, `,` the group
//! separator, and month and day names are English.
//!
//! This crate is a derivative work of the number formatter of Apache OpenOffice (`svl`), which is
//! licensed under the Apache License, Version 2.0. See the `LICENSE` file of this crate.

mod calendar;
mod condition;
mod format;
mod locale;
mod math;
mod scanner;
mod symbol;
mod text;

use std::fmt;

pub use format::NumberFormat;

/// A format code, or its combination with a value, that the formatter cannot process.
///
/// This is not the same as a format that shows nothing: that one formats to an empty string.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct FormatError;

impl fmt::Display for FormatError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("Unexpected result for number format")
    }
}

impl std::error::Error for FormatError {}

/// Formats `value` with the Excel number format `format`; a format that cannot be applied gives an empty string.
pub fn format_number(format: &str, value: f64) -> String {
    NumberFormat::new(format)
        .and_then(|format| format.format(value))
        .unwrap_or_default()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn format_number_formats() {
        assert_eq!(format_number("$#,##0.00", 1234.567), "$1,234.57");
        assert_eq!(format_number("yyyy-mm-dd", 45000.0), "2023-03-15");
        assert_eq!(format_number("", 1.0), "");
    }

    #[test]
    fn format_number_gives_an_empty_string_instead_of_an_error() {
        assert_eq!(format_number("[$", 1.0), "");
        assert_eq!(format_number("0E", 1.0), "");
        assert_eq!(format_number("# ?/?", f64::NAN), "");
    }

    #[test]
    fn format_error_has_the_message_of_the_sql_function() {
        assert_eq!(
            FormatError.to_string(),
            "Unexpected result for number format"
        );
    }
}
