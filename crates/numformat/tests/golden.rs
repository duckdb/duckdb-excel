//! Checks the regression corpora in `tests/data` (see the README there).

use std::fs;
use std::path::Path;

use excel_numformat::NumberFormat;

mod support;

/// Marks a case the formatter gives up on.
const ERROR: &str = "\\E";

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
        let format = String::from_utf8(support::unescape(format)).expect("formats are UTF-8");
        let value = support::parse_value(value).expect("values are numbers");
        let actual = match NumberFormat::new(&format).and_then(|format| format.format_bytes(value))
        {
            Ok(bytes) => support::escape(&bytes),
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
