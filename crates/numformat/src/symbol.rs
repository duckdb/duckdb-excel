//! The vocabulary of a scanned format code: keywords, symbol kinds and format types.

use crate::text::{Text, starts_with_upper};

/// A keyword of the format code language.
///
/// The scanner tries keywords in a fixed order: `GENERAL` wins anywhere, calendar keywords come
/// before the classic ones, and of two keywords the longer one is tried first. It never finds
/// `Exponent`, `Minute` and `Minute2` that way but derives them from `EraYear`, `Month` and
/// `Month2` by context.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Keyword {
    /// `E` of a scientific format.
    Exponent,
    AmPm,
    AP,
    Minute,
    Minute2,
    Month,
    Month2,
    MonthAbbreviation,
    MonthName,
    MonthInitial,
    Hour,
    Hour2,
    Second,
    Second2,
    Quarter,
    QuarterName,
    Day,
    Day2,
    DayAbbreviation,
    DayName,
    Year2,
    Year4,
    /// `NN`, day of week abbreviation.
    WeekdayAbbreviation,
    /// `NNN`, day of week name.
    WeekdayName,
    /// `NNNN`, day of week name followed by the long date separator.
    WeekdayNameSeparator,
    /// `CCC`, currency bank symbol.
    BankSymbol,
    General,
    WeekOfYear,
    /// `AAA`, day of week abbreviation.
    AltDayAbbreviation,
    /// `AAAA`, day of week name.
    AltDayName,
    /// `E`, calendar year without leading zero.
    EraYear,
    /// `EE`, calendar year.
    EraYear2,
    Era,
    Era2,
    EraName,
    /// `R`, same as `EE`.
    EraYearShort,
    /// `RR`, era name and year.
    EraYearLong,
}

impl Keyword {
    const SEARCH_ORDER: [Keyword; 34] = [
        Keyword::General,
        Keyword::EraYearLong,
        Keyword::EraYearShort,
        Keyword::EraName,
        Keyword::Era2,
        Keyword::Era,
        Keyword::EraYear2,
        Keyword::EraYear,
        Keyword::AltDayName,
        Keyword::AltDayAbbreviation,
        Keyword::MonthInitial,
        Keyword::WeekOfYear,
        Keyword::WeekdayNameSeparator,
        Keyword::WeekdayName,
        Keyword::BankSymbol,
        Keyword::WeekdayAbbreviation,
        Keyword::Year4,
        Keyword::Year2,
        Keyword::DayName,
        Keyword::DayAbbreviation,
        Keyword::Day2,
        Keyword::Day,
        Keyword::QuarterName,
        Keyword::Quarter,
        Keyword::Second2,
        Keyword::Second,
        Keyword::Hour2,
        Keyword::Hour,
        Keyword::MonthName,
        Keyword::MonthAbbreviation,
        Keyword::Month2,
        Keyword::Month,
        Keyword::AP,
        Keyword::AmPm,
    ];

    /// The upper case spelling.
    pub(crate) fn text(self) -> &'static [u8] {
        match self {
            Keyword::Exponent | Keyword::EraYear => b"E",
            Keyword::AmPm => b"AM/PM",
            Keyword::AP => b"A/P",
            Keyword::Minute | Keyword::Month => b"M",
            Keyword::Minute2 | Keyword::Month2 => b"MM",
            Keyword::MonthAbbreviation => b"MMM",
            Keyword::MonthName => b"MMMM",
            Keyword::MonthInitial => b"MMMMM",
            Keyword::Hour => b"H",
            Keyword::Hour2 => b"HH",
            Keyword::Second => b"S",
            Keyword::Second2 => b"SS",
            Keyword::Quarter => b"Q",
            Keyword::QuarterName => b"QQ",
            Keyword::Day => b"D",
            Keyword::Day2 => b"DD",
            Keyword::DayAbbreviation => b"DDD",
            Keyword::DayName => b"DDDD",
            Keyword::Year2 => b"YY",
            Keyword::Year4 => b"YYYY",
            Keyword::WeekdayAbbreviation => b"NN",
            Keyword::WeekdayName => b"NNN",
            Keyword::WeekdayNameSeparator => b"NNNN",
            Keyword::BankSymbol => b"CCC",
            Keyword::General => b"GENERAL",
            Keyword::WeekOfYear => b"WW",
            Keyword::AltDayAbbreviation => b"AAA",
            Keyword::AltDayName => b"AAAA",
            Keyword::EraYear2 => b"EE",
            Keyword::Era => b"G",
            Keyword::Era2 => b"GG",
            Keyword::EraName => b"GGG",
            Keyword::EraYearShort => b"R",
            Keyword::EraYearLong => b"RR",
        }
    }

    /// The keyword `text` starts with, ignoring case.
    pub(crate) fn at_start_of(text: &[u8]) -> Option<Keyword> {
        Keyword::SEARCH_ORDER
            .into_iter()
            .find(|keyword| starts_with_upper(text, keyword.text()))
    }

    /// Keywords that make a format a date format and are displayed from the calendar.
    pub(crate) fn is_date(self) -> bool {
        !matches!(
            self,
            Keyword::Exponent
                | Keyword::AmPm
                | Keyword::AP
                | Keyword::Minute
                | Keyword::Minute2
                | Keyword::Hour
                | Keyword::Hour2
                | Keyword::Second
                | Keyword::Second2
                | Keyword::BankSymbol
                | Keyword::General
        )
    }
}

/// The kind of one symbol of a scanned format code.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum SymbolKind {
    /// A slot the scanner has not filled.
    None,
    Keyword(Keyword),
    /// Literal text.
    String,
    /// A character with a meaning that depends on the format type.
    Delimiter,
    /// `_x`, a blank as wide as `x`.
    Blank,
    /// `*x`, fill with `x`.
    Star,
    /// A run of digit placeholders.
    Digit,
    DecimalSeparator,
    ThousandSeparator,
    /// `E+` or `E-` of a scientific format.
    Exponent,
    /// `/` of a fraction.
    Fraction,
    /// A symbol merged into another one.
    Empty,
    /// The blank between integer and fraction.
    FractionBlank,
    /// The symbol inside `[$...]`.
    Currency,
    /// `[$` or `]`.
    CurrencyDelimiter,
    /// `-xxx` inside `[$...]`.
    CurrencyExtension,
    DateSeparator,
    TimeSeparator,
    Time100SecSeparator,
    Percent,
}

impl SymbolKind {
    pub(crate) fn is_keyword(self) -> bool {
        matches!(self, SymbolKind::Keyword(_))
    }
}

/// One symbol of a scanned format code.
#[derive(Clone, Debug)]
pub(crate) struct Symbol {
    pub(crate) kind: SymbolKind,
    pub(crate) text: Text,
}

/// What a format code formats.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum FormatType {
    /// Nothing recognised; such a format produces no output.
    Undefined,
    /// Literal text only.
    Defined,
    Date,
    Time,
    DateTime,
    Currency,
    Number,
    Scientific,
    Fraction,
    Percent,
    Text,
}

impl FormatType {
    pub(crate) fn has_date(self) -> bool {
        matches!(self, FormatType::Date | FormatType::DateTime)
    }

    pub(crate) fn has_time(self) -> bool {
        matches!(self, FormatType::Time | FormatType::DateTime)
    }
}
