//! The vocabulary of a scanned format code: keywords, symbol kinds and format types.

use crate::text::{Text, starts_with_ignoring_case};

/// A keyword of the format code language.
///
/// The scanner tries keywords in a fixed order: `GENERAL` wins anywhere, calendar keywords come
/// before the classic ones, and of two keywords the longer one is tried first. It never finds
/// `Exponent`, `Minute` and `Minute2` that way but derives them from `EraYear`, `Month` and
/// `Month2` by context.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum Keyword {
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
    WeekdayAbbreviation,
    WeekdayName,
    WeekdayNameSeparator,
    BankSymbol,
    General,
    WeekOfYear,
    AltDayAbbreviation,
    AltDayName,
    EraYear,
    EraYear2,
    Era,
    Era2,
    EraName,
    EraYearShort,
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

    pub(crate) fn spelling(self) -> &'static [u8] {
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

    pub(crate) fn at_start_of(text: &[u8]) -> Option<Keyword> {
        Keyword::SEARCH_ORDER
            .into_iter()
            .find(|keyword| starts_with_ignoring_case(text, keyword.spelling()))
    }

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
///
/// `None` is a slot the scanner has not filled, `Empty` a symbol merged into another one and
/// `Delimiter` a character whose meaning depends on the format type. `Blank` is `_x`, a blank as
/// wide as `x`, and `Star` is `*x`, fill with `x`. `Digit` is a run of digit placeholders.
/// `Currency` is the symbol inside `[$...]`, `CurrencyDelimiter` its `[$` or `]` and
/// `CurrencyExtension` its `-xxx` part. `FractionBlank` separates integer and fraction.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum SymbolKind {
    None,
    Keyword(Keyword),
    String,
    Delimiter,
    Blank,
    Star,
    Digit,
    DecimalSeparator,
    ThousandSeparator,
    Exponent,
    Fraction,
    Empty,
    FractionBlank,
    Currency,
    CurrencyDelimiter,
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
///
/// `Undefined` is a section in which nothing was recognised; it produces no output. `Defined`
/// is one with literal text only.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) enum FormatType {
    Undefined,
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
