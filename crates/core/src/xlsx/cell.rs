//! Cell positions, ranges and types of a worksheet.

use std::fmt;

/// The longest text Excel stores in a cell, in characters.
pub const MAX_CELL_SIZE: usize = 32767;
pub const MAX_CELL_ROWS: u64 = 1_048_576;
pub const MAX_CELL_COLS: u64 = 16384;

/// A cell of a worksheet; rows and columns count from 1.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CellPos {
    pub row: u64,
    pub col: u64,
}

impl CellPos {
    pub const fn new(row: u64, col: u64) -> CellPos {
        CellPos { row, col }
    }

    /// Reads a reference like `B7`, `B` or `7` from the start of `text` into `self`, keeping what the reference
    /// leaves out; returns the rest of `text`, or `None` when `text` starts with neither letters nor digits.
    pub fn parse_prefix<'a>(&mut self, text: &'a [u8]) -> Option<&'a [u8]> {
        let letters = text
            .iter()
            .take_while(|byte| byte.is_ascii_uppercase())
            .count();
        let (letters, rest) = text.split_at(letters);
        let digits = match rest.first() {
            Some(b'1'..=b'9') => rest.iter().take_while(|byte| byte.is_ascii_digit()).count(),
            _ => 0,
        };
        let (digits, rest) = rest.split_at(digits);
        if letters.is_empty() && digits.is_empty() {
            return None;
        }
        if !letters.is_empty() {
            self.col = letters.iter().fold(0, |col: u64, letter| {
                col.wrapping_mul(26)
                    .wrapping_add(u64::from(letter - b'A') + 1)
            });
        }
        if !digits.is_empty() {
            self.row = digits.iter().fold(0, |row: u64, digit| {
                row.wrapping_mul(10).wrapping_add(u64::from(digit - b'0'))
            });
        }
        Some(rest)
    }

    /// The letters of the column, `A` for column 1.
    pub fn column_name(&self) -> String {
        let mut letters = Vec::new();
        let mut col = self.col;
        while col > 0 {
            letters.push(char::from(b'A' + ((col - 1) % 26) as u8));
            col = (col - 1) / 26;
        }
        letters.iter().rev().collect()
    }
}

impl fmt::Display for CellPos {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        write!(formatter, "{}{}", self.column_name(), self.row)
    }
}

/// A rectangle of cells, from `beg` up to but not including `end`.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CellRange {
    pub beg: CellPos,
    pub end: CellPos,
}

impl CellRange {
    /// Every row there can be, in every column but the last one Excel has.
    pub const WHOLE_SHEET: CellRange = CellRange {
        beg: CellPos::new(1, 1),
        end: CellPos::new(u64::MAX, MAX_CELL_COLS),
    };

    /// Parses `A1:B2`, where either corner may leave out its row or its column; what follows the range is ignored.
    /// The corners are returned as written, so `end` is still the last cell of the range.
    pub fn parse(text: &str) -> Option<CellRange> {
        let mut beg = CellPos::new(1, 1);
        let rest = beg.parse_prefix(text.as_bytes())?;
        let rest = rest.strip_prefix(b":")?;
        let mut end = CellPos::new(MAX_CELL_ROWS, MAX_CELL_COLS);
        end.parse_prefix(rest)?;
        Some(CellRange { beg, end })
    }

    pub fn contains_row(&self, row: u64) -> bool {
        row >= self.beg.row && row < self.end.row
    }

    pub fn contains_col(&self, col: u64) -> bool {
        col >= self.beg.col && col < self.end.col
    }

    pub fn contains(&self, pos: CellPos) -> bool {
        self.contains_col(pos.col) && self.contains_row(pos.row)
    }

    pub fn width(&self) -> u64 {
        self.end.col.wrapping_sub(self.beg.col)
    }

    pub fn is_valid(&self) -> bool {
        self.beg.row <= self.end.row && self.beg.col <= self.end.col
    }
}

/// What the `t` attribute of a cell says about its value.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CellType {
    Unknown,
    Number,
    Boolean,
    SharedString,
    InlineString,
    Date,
    Error,
    FormulaString,
}

impl CellType {
    /// Only the first letter decides for numbers, dates, booleans and errors.
    pub fn from_attribute(value: &str) -> CellType {
        match value.as_bytes() {
            [b'n', ..] => CellType::Number,
            [b's'] => CellType::SharedString,
            [b's', b't', b'r'] => CellType::FormulaString,
            [b'd', ..] => CellType::Date,
            [b'b', ..] => CellType::Boolean,
            [b'e', ..] => CellType::Error,
            b"inlineStr" => CellType::InlineString,
            _ => CellType::Unknown,
        }
    }
}

/// Parses a decimal integer the way C's `strtol` does: leading whitespace and a sign are allowed, parsing stops at
/// the first character that is not a digit, and values that do not fit are clamped.
pub fn parse_integer(text: &str) -> i64 {
    let text = text.trim_start_matches([' ', '\t', '\n', '\x0b', '\x0c', '\r']);
    let (negative, digits) = match text.as_bytes() {
        [b'-', digits @ ..] => (true, digits),
        [b'+', digits @ ..] => (false, digits),
        digits => (false, digits),
    };
    let mut value: i64 = 0;
    for digit in digits.iter().take_while(|byte| byte.is_ascii_digit()) {
        let digit = i64::from(digit - b'0');
        let next = value.checked_mul(10).and_then(|value| {
            if negative {
                value.checked_sub(digit)
            } else {
                value.checked_add(digit)
            }
        });
        match next {
            Some(next) => value = next,
            None => return if negative { i64::MIN } else { i64::MAX },
        }
    }
    value
}

#[cfg(test)]
mod tests {
    use super::*;

    fn parsed(text: &str) -> Option<(u64, u64, &str)> {
        let mut pos = CellPos::new(77, 88);
        let rest = pos.parse_prefix(text.as_bytes())?;
        Some((pos.row, pos.col, std::str::from_utf8(rest).unwrap()))
    }

    #[test]
    fn cell_references() {
        assert_eq!(parsed("A1"), Some((1, 1, "")));
        assert_eq!(parsed("XFD1048576"), Some((1_048_576, 16384, "")));
        assert_eq!(parsed("AA10:B2"), Some((10, 27, ":B2")));
        assert_eq!(parsed("B"), Some((77, 2, "")));
        assert_eq!(parsed("12"), Some((12, 88, "")));
        assert_eq!(parsed("B0"), Some((77, 2, "0")));
        assert_eq!(parsed("a1"), None);
        assert_eq!(parsed("0"), None);
        assert_eq!(parsed(""), None);
        assert_eq!(parsed(":A1"), None);
    }

    #[test]
    fn cell_references_wrap_instead_of_overflowing() {
        let (row, col, rest) =
            parsed("ZZZZZZZZZZZZZZZZZZZZZZZZ99999999999999999999999999").unwrap();
        assert_eq!(rest, "");
        assert_ne!((row, col), (77, 88));
    }

    #[test]
    fn column_names() {
        for (col, name) in [
            (1, "A"),
            (26, "Z"),
            (27, "AA"),
            (702, "ZZ"),
            (703, "AAA"),
            (16384, "XFD"),
        ] {
            assert_eq!(CellPos::new(1, col).column_name(), name);
        }
        assert_eq!(CellPos::new(12, 28).to_string(), "AB12");
    }

    #[test]
    fn ranges() {
        let range = CellRange::parse("B2:D7").unwrap();
        assert_eq!(
            (range.beg, range.end),
            (CellPos::new(2, 2), CellPos::new(7, 4))
        );
        let range = CellRange::parse("B:D").unwrap();
        assert_eq!(
            (range.beg, range.end),
            (CellPos::new(1, 2), CellPos::new(MAX_CELL_ROWS, 4))
        );
        let range = CellRange::parse("2:7").unwrap();
        assert_eq!(
            (range.beg, range.end),
            (CellPos::new(2, 1), CellPos::new(7, MAX_CELL_COLS))
        );
        assert!(CellRange::parse("A1:B2 and more").is_some());
        for invalid in ["A1", "A1:", ":B2", "a1:b2", "A1-B2", ""] {
            assert_eq!(CellRange::parse(invalid), None, "{invalid}");
        }
        assert!(!CellRange::parse("Z1:A5").unwrap().is_valid());
        assert!(!CellRange::parse("A5:A1").unwrap().is_valid());
        assert!(CellRange::parse("A1:A1").unwrap().is_valid());
    }

    #[test]
    fn range_membership() {
        let range = CellRange {
            beg: CellPos::new(2, 2),
            end: CellPos::new(4, 5),
        };
        assert_eq!(range.width(), 3);
        assert!(range.contains(CellPos::new(2, 2)) && range.contains(CellPos::new(3, 4)));
        assert!(!range.contains(CellPos::new(4, 2)) && !range.contains(CellPos::new(2, 5)));
        assert!(!range.contains(CellPos::new(1, 2)) && !range.contains(CellPos::new(2, 1)));
    }

    #[test]
    fn cell_types() {
        for (text, cell_type) in [
            ("n", CellType::Number),
            ("number", CellType::Number),
            ("s", CellType::SharedString),
            ("str", CellType::FormulaString),
            ("st", CellType::Unknown),
            ("d", CellType::Date),
            ("b", CellType::Boolean),
            ("e", CellType::Error),
            ("inlineStr", CellType::InlineString),
            ("inline", CellType::Unknown),
            ("", CellType::Unknown),
            ("x", CellType::Unknown),
        ] {
            assert_eq!(CellType::from_attribute(text), cell_type, "{text}");
        }
    }

    #[test]
    fn integers_like_strtol() {
        for (text, value) in [
            ("42", 42),
            ("  -7abc", -7),
            ("+3", 3),
            ("", 0),
            ("abc", 0),
            ("-", 0),
            ("12.9", 12),
            ("99999999999999999999", i64::MAX),
            ("-99999999999999999999", i64::MIN),
            ("\t\n 5", 5),
        ] {
            assert_eq!(parse_integer(text), value, "{text:?}");
        }
    }
}
