//! Turns the text of header cells into column names.

use unicode_normalization::UnicodeNormalization;

use crate::duckdb_compat::is_reserved_keyword;

fn is_space_separator(character: char) -> bool {
    // The Unicode category Zs.
    matches!(
        character,
        '\u{0020}' | '\u{00A0}' | '\u{1680}' | '\u{2000}'
            ..='\u{200A}' | '\u{202F}' | '\u{205F}' | '\u{3000}'
    )
}

/// Removes the space separators around `name`; tabs and line breaks are not space separators and stay.
pub fn trim_whitespace(name: &str) -> &str {
    name.trim_matches(is_space_separator)
}

/// Makes `name` an identifier that needs no quotes: letters without accents and digits in lower case, words
/// joined by `_`, and a leading `_` where the name would otherwise start with a digit or be a reserved keyword.
pub fn normalize_column_name(name: &str) -> String {
    let mut ascii = String::with_capacity(name.len());
    for character in name.nfkd() {
        match character {
            '_' | '0'..='9' | 'A'..='Z' | 'a'..='z' => ascii.push(character),
            ' ' | '\t' | '\n' | '\x0b' | '\x0c' | '\r' => ascii.push(' '),
            _ => {}
        }
    }
    let mut cleaned = String::with_capacity(ascii.len() + 1);
    for word in ascii.split(' ').filter(|word| !word.is_empty()) {
        if !cleaned.is_empty() {
            cleaned.push('_');
        }
        cleaned.push_str(word);
    }
    if cleaned.is_empty() {
        return "_".to_owned();
    }
    cleaned.make_ascii_lowercase();
    if is_reserved_keyword(&cleaned) || cleaned.starts_with(|first: char| first.is_ascii_digit()) {
        cleaned.insert(0, '_');
    }
    cleaned
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn trims_space_separators_only() {
        assert_eq!(trim_whitespace("  a b  "), "a b");
        assert_eq!(trim_whitespace("\u{00A0}\u{3000}a\u{2003}\u{202F}"), "a");
        assert_eq!(trim_whitespace("\ta\n"), "\ta\n");
        assert_eq!(trim_whitespace(" \u{200B}a "), "\u{200B}a");
        assert_eq!(trim_whitespace("   "), "");
        assert_eq!(trim_whitespace(""), "");
    }

    #[test]
    fn normalizes_names() {
        for (name, normalized) in [
            ("Hello World", "hello_world"),
            ("  Hello   World  ", "hello_world"),
            ("A", "a"),
            ("Ünïcödé Nämé", "unicode_name"),
            ("ｆｕｌｌ　ｗｉｄｔｈ", "full_width"),
            ("ﬁrst", "first"),
            ("a\tb\nc", "a_b_c"),
            ("a-b.c", "abc"),
            ("snake_case", "snake_case"),
            ("1st", "_1st"),
            ("①", "_1"),
            ("Select", "_select"),
            ("TRY_CAST", "_try_cast"),
            ("selected", "selected"),
            ("", "_"),
            ("???", "_"),
            ("   ", "_"),
            ("_", "_"),
            ("日本語", "_"),
            ("a\u{00A0}b", "a_b"),
        ] {
            assert_eq!(normalize_column_name(name), normalized, "{name:?}");
        }
    }
}
