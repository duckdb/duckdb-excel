//! Escaping of text for XML and its reverse, and the escapes of the text of a workbook.

/// The length of an `_xHHHH_` escape.
const ESCAPE_LEN: usize = 7;

/// Escapes the five characters XML reserves and drops NUL, which XML cannot hold.
pub fn escape_xml(text: &str) -> String {
    let mut escaped = String::with_capacity(text.len());
    for character in text.chars() {
        match character {
            '\0' => {}
            '&' => escaped.push_str("&amp;"),
            '<' => escaped.push_str("&lt;"),
            '>' => escaped.push_str("&gt;"),
            '"' => escaped.push_str("&quot;"),
            '\'' => escaped.push_str("&apos;"),
            _ => escaped.push(character),
        }
    }
    escaped
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum ByteClass {
    Plain,
    Reserved,
    Control,
    Underscore,
    /// The first byte of U+F000 to U+FFFF, of which the last two are not characters of XML.
    HighLead,
}

const fn class_of(byte: u8) -> ByteClass {
    match byte {
        b'&' | b'<' | b'>' | b'"' | b'\'' => ByteClass::Reserved,
        b'\t' | b'\n' | b'\r' => ByteClass::Plain,
        0..b' ' => ByteClass::Control,
        b'_' => ByteClass::Underscore,
        0xEF => ByteClass::HighLead,
        _ => ByteClass::Plain,
    }
}

const BYTE_CLASSES: [ByteClass; 256] = {
    let mut classes = [ByteClass::Plain; 256];
    let mut byte = 0;
    while byte < classes.len() {
        classes[byte] = class_of(byte as u8);
        byte += 1;
    }
    classes
};

fn byte_class(byte: u8) -> ByteClass {
    BYTE_CLASSES
        .get(usize::from(byte))
        .copied()
        .unwrap_or(ByteClass::Plain)
}

fn push_escaped_unit(unit: u16, escaped: &mut Vec<u8>) {
    const DIGITS: &[u8; 16] = b"0123456789ABCDEF";
    escaped.extend_from_slice(b"_x");
    for shift in [12, 8, 4, 0] {
        let digit = DIGITS.get(usize::from(unit >> shift & 0xF));
        escaped.push(digit.copied().unwrap_or(b'0'));
    }
    escaped.push(b'_');
}

fn is_written_as_escape(text: &[u8]) -> bool {
    match text {
        [first, ..] if byte_class(*first) == ByteClass::Control => true,
        [0xEF, 0xBF, 0xBE | 0xBF, ..] => true,
        _ => false,
    }
}

fn would_read_as_escape(text: &[u8]) -> bool {
    let [b'_', b'x', digits @ .., next] = text.get(..ESCAPE_LEN).unwrap_or_default() else {
        return false;
    };
    // What follows the digits is written with an underscore in front when it is escaped itself.
    let ends_in_underscore =
        *next == b'_' || is_written_as_escape(text.get(ESCAPE_LEN - 1..).unwrap_or_default());
    ends_in_underscore && digits.iter().all(u8::is_ascii_hexdigit)
}

/// Appends `text` to `escaped` as the text of a cell: the five characters XML reserves as entities, and the
/// characters XML cannot hold, NUL among them, as the `_xHHHH_` escapes [`decode_escaped_units`] resolves. An
/// underscore that would read as the start of such an escape is escaped itself.
pub fn escape_text_into(text: &[u8], escaped: &mut Vec<u8>) {
    let mut rest = text;
    while let Some(index) = rest
        .iter()
        .position(|byte| byte_class(*byte) != ByteClass::Plain)
    {
        let (plain, special) = rest.split_at(index);
        escaped.extend_from_slice(plain);
        let taken = match special {
            [b'&', ..] => {
                escaped.extend_from_slice(b"&amp;");
                1
            }
            [b'<', ..] => {
                escaped.extend_from_slice(b"&lt;");
                1
            }
            [b'>', ..] => {
                escaped.extend_from_slice(b"&gt;");
                1
            }
            [b'"', ..] => {
                escaped.extend_from_slice(b"&quot;");
                1
            }
            [b'\'', ..] => {
                escaped.extend_from_slice(b"&apos;");
                1
            }
            [0xEF, 0xBF, last @ (0xBE | 0xBF), ..] => {
                push_escaped_unit(0xFFC0 | u16::from(*last & 0x3F), escaped);
                3
            }
            [control, ..] if byte_class(*control) == ByteClass::Control => {
                push_escaped_unit(u16::from(*control), escaped);
                1
            }
            [b'_', ..] if would_read_as_escape(special) => {
                escaped.extend_from_slice(b"_x005F_");
                1
            }
            [other, ..] => {
                escaped.push(*other);
                1
            }
            [] => 0,
        };
        rest = special.get(taken..).unwrap_or_default();
    }
    escaped.extend_from_slice(rest);
}

/// Replaces the five predefined entities; any other `&` stays as it is.
pub fn unescape_xml_entities(text: &str) -> String {
    let mut unescaped = String::with_capacity(text.len());
    let mut rest = text;
    while let Some(start) = rest.find('&') {
        let (before, entity_start) = rest.split_at(start);
        unescaped.push_str(before);
        let replacement = entity_start.find(';').and_then(|end| {
            let (entity, after) = entity_start.split_at(end + 1);
            let character = match entity {
                "&lt;" => '<',
                "&gt;" => '>',
                "&amp;" => '&',
                "&quot;" => '"',
                "&apos;" => '\'',
                _ => return None,
            };
            Some((character, after))
        });
        match replacement {
            Some((character, after)) => {
                unescaped.push(character);
                rest = after;
            }
            None => {
                unescaped.push('&');
                rest = entity_start.get(1..).unwrap_or_default();
            }
        }
    }
    unescaped.push_str(rest);
    unescaped
}

fn escaped_unit(text: &[u8]) -> Option<u32> {
    let [b'_', b'x', digits @ .., b'_'] = text.get(..ESCAPE_LEN)? else {
        return None;
    };
    digits.iter().try_fold(0, |unit, digit| {
        Some(unit << 4 | char::from(*digit).to_digit(16)?)
    })
}

fn escaped_character(text: &[u8]) -> Option<(Option<char>, usize)> {
    let unit = escaped_unit(text)?;
    if !(0xD800..0xDC00).contains(&unit) {
        return Some((char::from_u32(unit), ESCAPE_LEN));
    }
    let low = text
        .get(ESCAPE_LEN..)
        .and_then(escaped_unit)
        .filter(|low| (0xDC00..0xE000).contains(low));
    Some(match low {
        Some(low) => {
            let code = 0x10000 + ((unit - 0xD800) << 10) + (low - 0xDC00);
            (char::from_u32(code), 2 * ESCAPE_LEN)
        }
        None => (None, ESCAPE_LEN),
    })
}

/// Replaces the `_xHHHH_` escapes in `text` from byte `start` on with the UTF-16 code units they name;
/// `_x005F_` is an underscore that starts no escape, and half a surrogate pair stays as it is written.
pub fn decode_escaped_units(text: &mut String, start: usize) {
    // Most text has no underscore at all, which is the quickest to find out.
    let has_underscore = text
        .as_bytes()
        .get(start..)
        .is_some_and(|tail| tail.contains(&b'_'));
    if has_underscore {
        decode_escapes_after_underscore(text, start);
    }
}

fn decode_escapes_after_underscore(text: &mut String, start: usize) {
    let Some(escaped) = text.get(start..).filter(|tail| tail.contains("_x")) else {
        return;
    };
    let mut decoded = String::with_capacity(escaped.len());
    let mut rest = escaped;
    while let Some(index) = rest.find("_x") {
        let (plain, candidate) = rest.split_at(index);
        decoded.push_str(plain);
        let (character, len) = escaped_character(candidate.as_bytes()).unwrap_or((None, 1));
        let (written, after) = candidate.split_at_checked(len).unwrap_or((candidate, ""));
        match character {
            Some(character) => decoded.push(character),
            None => decoded.push_str(written),
        }
        rest = after;
    }
    decoded.push_str(rest);
    text.truncate(start);
    text.push_str(&decoded);
}

#[cfg(test)]
mod tests {
    use super::*;

    fn decoded(text: &str) -> String {
        let mut text = text.to_owned();
        decode_escaped_units(&mut text, 0);
        text
    }

    #[test]
    fn escapes_become_the_characters_they_name() {
        for (escaped, text) in [
            ("a_x000D_b", "a\rb"),
            ("a_x000d__x000A_", "a\r\n"),
            ("_x0009_", "\t"),
            ("_x0008__x001F_", "\u{8}\u{1f}"),
            ("_x00e9_ _x00E9_", "é é"),
            ("_x0000_", "\0"),
            ("__x0041__", "_A_"),
            ("é_x0041_é", "éAé"),
            ("_x0041__x0042__x0043_", "ABC"),
            ("_xFFFF_", "\u{ffff}"),
        ] {
            assert_eq!(decoded(escaped), text, "{escaped}");
        }
    }

    #[test]
    fn an_escaped_underscore_starts_no_escape() {
        assert_eq!(decoded("_x005F_"), "_");
        assert_eq!(decoded("_x005F_x000D_"), "_x000D_");
        assert_eq!(decoded("_x005f__x000D_"), "_\r");
        assert_eq!(decoded("a_x005F_b"), "a_b");
    }

    #[test]
    fn surrogate_pairs_are_one_character_and_halves_stay() {
        assert_eq!(decoded("_xD83D__xDE00_"), "\u{1F600}");
        assert_eq!(decoded("_xd83d__xde00_!"), "\u{1F600}!");
        assert_eq!(decoded("_xD83D_"), "_xD83D_");
        assert_eq!(decoded("_xDE00_"), "_xDE00_");
        assert_eq!(decoded("_xDE00__xD83D_"), "_xDE00__xD83D_");
        assert_eq!(decoded("_xD83D__x0041_"), "_xD83D_A");
        assert_eq!(decoded("_xD83D__xD83D__xDE00_"), "_xD83D_\u{1F600}");
        assert_eq!(decoded("_xD83D_x0041_"), "_xD83D_x0041_");
    }

    #[test]
    fn text_that_only_looks_like_an_escape_stays() {
        for text in [
            "",
            "plain",
            "a_xb",
            "_x",
            "_x000D",
            "_x12G4_",
            "_x 041_",
            "_x+041_",
            "_X000D_",
            "x000D_",
            "_x000D",
            "_x00D_",
            "_x0000D_",
            "_xé00_",
            "snake_xcase_name",
        ] {
            assert_eq!(decoded(text), text, "{text}");
        }
    }

    #[test]
    fn only_the_text_from_the_start_on_is_decoded() {
        let mut text = "_x0041_ kept _x0042_ and _x0043_".to_owned();
        decode_escaped_units(&mut text, 13);
        assert_eq!(text, "_x0041_ kept B and C");
        let end = text.len();
        decode_escaped_units(&mut text, end);
        decode_escaped_units(&mut text, 500);
        assert_eq!(text, "_x0041_ kept B and C");
        // An escape that starts before the start is not one.
        let mut text = "ab_x0041_".to_owned();
        decode_escaped_units(&mut text, 3);
        assert_eq!(text, "ab_x0041_");
    }

    #[test]
    fn unescaping_undoes_escaping() {
        for text in ["", "plain", "a&b<c>d\"e'f", "&amp;", "&&;;", "&unknown; &"] {
            assert_eq!(unescape_xml_entities(&escape_xml(text)), text);
        }
    }

    fn escaped(text: &str) -> String {
        let mut escaped = Vec::new();
        escape_text_into(text.as_bytes(), &mut escaped);
        String::from_utf8(escaped).expect("escaped text in UTF-8")
    }

    #[test]
    fn cell_text_keeps_what_xml_can_hold() {
        for (text, written) in [
            ("", ""),
            ("plain", "plain"),
            ("a&b<c>d\"e'f", "a&amp;b&lt;c&gt;d&quot;e&apos;f"),
            ("<<>>", "&lt;&lt;&gt;&gt;"),
            ("é&ü 漢字 \u{1F600}", "é&amp;ü 漢字 \u{1F600}"),
            ("tab\there\nand\rthere", "tab\there\nand\rthere"),
            (
                "\u{FFFD}\u{FF01}\u{F000}\u{EFFF}",
                "\u{FFFD}\u{FF01}\u{F000}\u{EFFF}",
            ),
            ("snake_case_x_name_", "snake_case_x_name_"),
        ] {
            assert_eq!(escaped(text), written, "{text:?}");
        }
        let mut target = b"kept ".to_vec();
        escape_text_into(b"\xff<\xfe\xef\xef\xbf", &mut target);
        assert_eq!(target, b"kept \xff&lt;\xfe\xef\xef\xbf");
    }

    #[test]
    fn characters_xml_cannot_hold_are_written_as_escapes() {
        assert_eq!(escaped("\0a\0\0b\0"), "_x0000_a_x0000__x0000_b_x0000_");
        assert_eq!(
            escaped("\u{1}\u{8}\u{b}\u{c}\u{e}\u{1f}"),
            "_x0001__x0008__x000B__x000C__x000E__x001F_"
        );
        assert_eq!(escaped("a\u{FFFE}b\u{FFFF}"), "a_xFFFE_b_xFFFF_");
        for code in (0..0x20).filter(|code| ![9, 10, 13].contains(code)) {
            let text = char::from_u32(code).unwrap().to_string();
            assert_eq!(escaped(&text), format!("_x{code:04X}_"));
        }
        assert_eq!(escaped("\u{7f}\u{80}\u{9f}"), "\u{7f}\u{80}\u{9f}");
    }

    #[test]
    fn an_underscore_that_would_start_an_escape_is_escaped() {
        for (text, written) in [
            ("_x000D_", "_x005F_x000D_"),
            ("_x005F_", "_x005F_x005F_"),
            ("a_x0041_b", "a_x005F_x0041_b"),
            ("__x0041__", "__x005F_x0041__"),
            ("_x0041_x0042_", "_x005F_x0041_x005F_x0042_"),
            ("_xD83D_", "_x005F_xD83D_"),
            ("_xabcd_", "_x005F_xabcd_"),
            // What follows the digits is written with an underscore of its own.
            ("_x0041\u{1}", "_x005F_x0041_x0001_"),
            ("_x0041\0", "_x005F_x0041_x0000_"),
            ("_x0041\u{FFFF}", "_x005F_x0041_xFFFF_"),
            ("_\u{1}", "__x0001_"),
            // No escape, so nothing to keep apart.
            ("_x0041", "_x0041"),
            ("_x0041 ", "_x0041 "),
            ("_x0041\t", "_x0041\t"),
            ("_x0041\u{FFFD}", "_x0041\u{FFFD}"),
            ("_x004_", "_x004_"),
            ("_x00411_", "_x00411_"),
            ("_X0041_", "_X0041_"),
            ("_x004g_", "_x004g_"),
            ("_y0041_", "_y0041_"),
            ("x0041_", "x0041_"),
        ] {
            assert_eq!(escaped(text), written, "{text:?}");
        }
    }

    /// Numbers that look random and are the same in every run.
    fn sequence(mut state: u64) -> impl FnMut() -> u64 {
        move || {
            state ^= state << 13;
            state ^= state >> 7;
            state ^= state << 17;
            state
        }
    }

    #[test]
    fn escaped_text_reads_back_as_it_was() {
        let pieces = [
            "_",
            "x",
            "_x",
            "0",
            "00",
            "005F",
            "000D",
            "D83D",
            "DE00",
            "d83d",
            "FFFF",
            "41",
            "g",
            "\0",
            "\u{1}",
            "\u{8}",
            "\u{b}",
            "\u{1f}",
            "\t",
            "\n",
            " ",
            "&",
            "<",
            "é",
            "漢",
            "\u{1F600}",
            "\u{FFFE}",
            "\u{FFFF}",
            "\u{FFFD}",
            "\u{F000}",
            "_x005F_",
            "_x000D_",
            "_xD83D_",
            "_xDE00_",
            "_x0041",
            "x0041_",
        ];
        let mut next = sequence(0x9E37_79B9_7F4A_7C15);
        for _ in 0..200_000 {
            let text: String = (0..next() % 12)
                .map(|_| pieces[(next() % pieces.len() as u64) as usize])
                .collect();
            let written = escaped(&text);
            assert!(
                !written
                    .bytes()
                    .any(|byte| byte_class(byte) == ByteClass::Control),
                "{text:?} is written as {written:?}"
            );
            // An XML parser hands the text on with its entities resolved.
            let mut read = unescape_xml_entities(&written);
            decode_escaped_units(&mut read, 0);
            assert_eq!(read, text, "written as {written:?}");
        }
        for code in 0..=0xFFFF {
            let Some(character) = char::from_u32(code) else {
                continue;
            };
            let text = format!("_x0041{character}_x{code:04X}_{character}");
            let mut read = unescape_xml_entities(&escaped(&text));
            decode_escaped_units(&mut read, 0);
            assert_eq!(read, text, "U+{code:04X}");
        }
    }

    #[test]
    fn unescaping_keeps_what_is_not_an_entity() {
        assert_eq!(unescape_xml_entities("a &amp; b"), "a & b");
        assert_eq!(unescape_xml_entities("&lt;&gt;&quot;&apos;"), "<>\"'");
        assert_eq!(unescape_xml_entities("a & b; c"), "a & b; c");
        assert_eq!(unescape_xml_entities("&nbsp;&amp;"), "&nbsp;&");
        assert_eq!(unescape_xml_entities("fish & chips"), "fish & chips");
        assert_eq!(unescape_xml_entities("&amp;amp;"), "&amp;");
    }
}
