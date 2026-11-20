//! Checks the regression corpora in `tests/data` (see the README there).

use std::fs;
use std::path::Path;

use excel_numformat::NumberFormat;

// Marks a case the formatter gives up on.
const ERROR: &str = "\\E";

fn unescape(text: &str) -> Vec<u8> {
    let bytes = text.as_bytes();
    let mut result = Vec::with_capacity(bytes.len());
    let mut index = 0;
    while index < bytes.len() {
        let byte = bytes[index];
        index += 1;
        if byte != b'\\' || index >= bytes.len() {
            result.push(byte);
            continue;
        }
        let escaped = bytes[index];
        index += 1;
        match escaped {
            b't' => result.push(b'\t'),
            b'n' => result.push(b'\n'),
            b'r' => result.push(b'\r'),
            b'x' => {
                let hex = text.get(index..index + 2);
                let value = hex.and_then(|hex| u8::from_str_radix(hex, 16).ok());
                result.push(value.expect("\\x is followed by two hex digits"));
                index += 2;
            }
            other => result.push(other),
        }
    }
    result
}

fn escape(bytes: &[u8]) -> String {
    let mut result = String::with_capacity(bytes.len());
    for &byte in bytes {
        match byte {
            b'\\' => result.push_str("\\\\"),
            b'\t' => result.push_str("\\t"),
            b'\n' => result.push_str("\\n"),
            b'\r' => result.push_str("\\r"),
            0x20..=0x7e => result.push(char::from(byte)),
            _ => result.push_str(&format!("\\x{byte:02X}")),
        }
    }
    result
}

fn parse_value(text: &str) -> Option<f64> {
    match text.strip_prefix("0x") {
        Some(bits) if bits.len() == 16 => u64::from_str_radix(bits, 16).ok().map(f64::from_bits),
        _ => text.parse().ok(),
    }
}

fn check_corpus(name: &str) {
    let path = Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("tests/data")
        .join(name);
    let corpus = fs::read_to_string(&path).expect("corpus file is readable");
    let mut failures = Vec::new();
    let mut count = 0;
    for (number, line) in corpus.lines().enumerate() {
        count += 1;
        let mut fields = line.splitn(3, '\t');
        let (Some(format), Some(value), Some(expected)) =
            (fields.next(), fields.next(), fields.next())
        else {
            panic!("{name}:{}: expected three fields", number + 1);
        };
        let format = String::from_utf8(unescape(format)).expect("formats are UTF-8");
        let value = parse_value(value).expect("values are numbers");
        let actual = match NumberFormat::new(&format).and_then(|format| format.format_bytes(value))
        {
            Ok(bytes) => escape(&bytes),
            Err(_) => ERROR.to_string(),
        };
        if actual != expected {
            failures.push(format!(
                "{name}:{}: expected {expected:?}, got {actual:?}",
                number + 1
            ));
        }
    }
    assert!(count > 0, "{name} has no cases");
    assert!(
        failures.is_empty(),
        "{} of {count} cases differ, first ones:\n{}",
        failures.len(),
        failures[..failures.len().min(20)].join("\n")
    );
}

#[test]
fn golden() {
    check_corpus("golden.tsv");
}

#[test]
#[cfg(all(target_os = "linux", target_env = "gnu", target_arch = "x86_64"))]
fn golden_glibc() {
    check_corpus("golden_glibc.tsv");
}
