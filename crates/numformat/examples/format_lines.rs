//! Formats `format<TAB>value` lines from standard input and prints one result line per case:
//! `ok<TAB>result`, or `throw` where the format cannot be applied.
//!
//! Formats, values and results are encoded as in `tests/data/README.md`. The line format suits
//! comparing the formatter with another implementation on large generated inputs.

use std::io::{self, BufRead, BufWriter, Write};

use excel_numformat::NumberFormat;

#[path = "../tests/support/mod.rs"]
mod support;

fn main() -> io::Result<()> {
    let mut output = BufWriter::new(io::stdout().lock());
    for line in io::stdin().lock().lines() {
        let line = line?;
        let (format, value) = line.split_once('\t').unwrap_or((&line, "0"));
        let value = support::parse_value(value).unwrap_or_default();
        let Ok(format) = String::from_utf8(support::unescape(format)) else {
            writeln!(output, "skip")?;
            continue;
        };
        match NumberFormat::new(&format).and_then(|format| format.format_bytes(value)) {
            Ok(bytes) => writeln!(output, "ok\t{}", support::escape(&bytes))?,
            Err(_) => writeln!(output, "throw")?,
        }
    }
    Ok(())
}
