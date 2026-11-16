//! Excel number formats, as the `TEXT` function of Microsoft Excel applies them.

/// Formats `value` with the Excel number format `format`; a format that cannot be applied gives an empty string.
pub fn format_number(format: &str, value: f64) -> String {
    let _ = (format, value);
    String::new()
}
