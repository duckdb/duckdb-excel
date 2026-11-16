//! A parsed format code and the formatting of values with it.
//!
//! A format code has up to four sections separated by `;`: by default for positive numbers,
//! negative numbers, zero and text. `[<100]`-style conditions in front of the first two sections
//! replace the default ranges. Other bracketed prefixes, such as colours and locales, are skipped.
//! While the sections are read the format code is rewritten in place, and positions in it are
//! 16-bit quantities; both show in how overlong or malformed codes are read.
//!
//! A format whose sections are all text formats shows numbers in the standard notation, which is
//! also what `General` stands for: two decimals without trailing zeros, scientific above 1e15.
//!
//! Number-like output is assembled from the last symbol of a section to the first: `number_fill`
//! merges a string of digits with the symbols down to a stop symbol, `fill_with_thousands` does
//! the same for the integer digits and adds the group separators.

use crate::FormatError;
use crate::calendar::{CalendarDate, EPOCH_OFFSET};
use crate::condition::parse_limit;
use crate::locale;
use crate::math::{Notation, double_to_string, integer_digits, pow, round_corrected, to_u64};
use crate::scanner::{DigitGrouping, FormatScanner, GENERAL_IN_FORMAT, ScanInfo};
use crate::symbol::{FormatType, Keyword, Symbol, SymbolKind};
use crate::text::{self, Text, count_from};

/// Largest value of the integer part of fractions and of elapsed seconds.
const MAX_U32: f64 = 4294967295.0;

/// Comparison of a condition.
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum LimitOp {
    None,
    Equal,
    NotEqual,
    Less,
    LessEqual,
    Greater,
    GreaterEqual,
}

impl LimitOp {
    fn check(self, number: f64, limit: f64) -> Option<bool> {
        match self {
            LimitOp::None => None,
            LimitOp::Equal => Some(number == limit),
            LimitOp::NotEqual => Some(number != limit),
            LimitOp::Less => Some(number < limit),
            LimitOp::LessEqual => Some(number <= limit),
            LimitOp::Greater => Some(number > limit),
            LimitOp::GreaterEqual => Some(number >= limit),
        }
    }
}

/// What the format code holds at the current position.
#[derive(Clone, Copy, PartialEq, Eq)]
enum BracketSymbol {
    /// The text of a section.
    Format,
    /// A bracketed prefix without function, such as `[Red]` or `[$-409]`.
    Prefix,
    Error,
    Condition(LimitOp),
}

/// One section of a format code.
#[derive(Clone, Debug)]
struct Section {
    symbols: Vec<Symbol>,
    info: ScanInfo,
}

impl Section {
    fn kind(&self, index: usize) -> SymbolKind {
        self.symbols
            .get(index)
            .map_or(SymbolKind::None, |symbol| symbol.kind)
    }

    fn text(&self, index: usize) -> &[u8] {
        self.symbols
            .get(index)
            .map_or(&[], |symbol| symbol.text.as_slice())
    }

    fn shows_separator(&self, index: usize) -> bool {
        index == 0
            || !matches!(
                self.kind(index - 1),
                SymbolKind::Digit | SymbolKind::ThousandSeparator
            )
            || self.kind(index + 1) == SymbolKind::Digit
    }
}

/// A parsed Excel number format that can be applied to many values.
#[derive(Clone, Debug)]
pub struct NumberFormat {
    sections: [Section; 3],
    limit1: f64,
    limit2: f64,
    op1: LimitOp,
    op2: LimitOp,
    is_text: bool,
}

fn is_escaped(format: &[u8], before: u16) -> Result<bool, FormatError> {
    let mut escaped = false;
    let mut position = usize::from(before);
    // An odd run of \, _ and * makes the next character part of an escape, a blank or a fill.
    while position > 0 && matches!(text::at(format, position)?, b'*' | b'\\' | b'_') {
        escaped = !escaped;
        position -= 1;
    }
    Ok(escaped)
}

fn erase_blanks(format: &mut Text, position: u16, length: u16) -> Result<u16, FormatError> {
    let start = usize::from(position);
    let blanks = format
        .get(start..usize::from(length))
        .ok_or(FormatError)?
        .iter()
        .take_while(|&&byte| byte == b' ')
        .count();
    text::erase(format, start, blanks)?;
    Ok(blanks as u16)
}

fn next_bracket_symbol(
    format: &mut Text,
    position: &mut u16,
    symbol: &mut Text,
) -> Result<BracketSymbol, FormatError> {
    enum State {
        Start,
        Bracketed,
        Condition,
        Section,
        Prefix,
        Time,
    }
    let is_time_letter = |token: u8| matches!(token.to_ascii_uppercase(), b'H' | b'M' | b'S');
    let mut kind = BracketSymbol::Format;
    let mut letter = b' ';
    // With 16-bit positions a longer format code is read up to its length modulo 65536.
    let mut length = format.len() as u16;
    let mut state = State::Start;
    symbol.clear();
    while *position < length {
        let token = text::at(format, usize::from(*position))?;
        *position += 1;
        if token == b' ' && matches!(state, State::Start | State::Bracketed | State::Condition) {
            *position -= 1;
            length -= erase_blanks(format, *position, length)?;
            continue;
        }
        match state {
            State::Start => match token {
                b'[' => {
                    state = State::Bracketed;
                    symbol.push(token);
                }
                b';' => {
                    state = State::Section;
                    *position -= 1;
                }
                b']' => return Ok(BracketSymbol::Error),
                _ => {
                    symbol.push(token);
                    state = State::Section;
                }
            },
            State::Bracketed => match token {
                b'<' | b'>' | b'=' => {
                    text::erase_all(symbol, b'[');
                    symbol.push(token);
                    letter = token;
                    state = State::Condition;
                    kind = BracketSymbol::Condition(match token {
                        b'<' => LimitOp::Less,
                        b'>' => LimitOp::Greater,
                        _ => LimitOp::Equal,
                    });
                }
                b'$' => {
                    if text::at(format, usize::from(*position))? == b'-' {
                        text::erase_all(symbol, b'[');
                        kind = BracketSymbol::Prefix;
                        state = State::Prefix;
                    } else {
                        state = State::Section;
                    }
                    symbol.push(token);
                }
                b'~' => {
                    symbol.push(token);
                    state = State::Section;
                }
                _ if is_time_letter(token) => {
                    symbol.push(token);
                    state = State::Time;
                    letter = token;
                }
                _ => {
                    text::erase_all(symbol, b'[');
                    symbol.push(token);
                    kind = BracketSymbol::Prefix;
                    state = State::Prefix;
                }
            },
            State::Section => {
                if token == b';' && *position >= 2 && !is_escaped(format, *position - 2)? {
                    break;
                }
                symbol.push(token);
            }
            State::Time => {
                if token == b']' {
                    symbol.push(token);
                    state = State::Section;
                    kind = BracketSymbol::Format;
                } else if is_time_letter(token) && letter == token {
                    symbol.push(token);
                    letter = b' ';
                } else {
                    // Not an elapsed time bracket after all; mixed time letters stay section text.
                    if !is_time_letter(token) {
                        kind = BracketSymbol::Prefix;
                    }
                    text::erase_all(symbol, b'[');
                    symbol.push(token);
                    state = State::Prefix;
                }
            }
            State::Condition => match (token, letter) {
                (b'>', b'<') => {
                    symbol.push(token);
                    return Ok(BracketSymbol::Condition(LimitOp::NotEqual));
                }
                (b'=', b'<') => {
                    symbol.push(token);
                    letter = b' ';
                    kind = BracketSymbol::Condition(LimitOp::LessEqual);
                }
                (b'=', b'>') => {
                    symbol.push(token);
                    letter = b' ';
                    kind = BracketSymbol::Condition(LimitOp::GreaterEqual);
                }
                (b'<' | b'>' | b'=', _) => return Ok(BracketSymbol::Error),
                _ => {
                    *position -= 1;
                    break;
                }
            },
            State::Prefix => {
                if token == b']' {
                    break;
                }
                symbol.push(token);
            }
        }
    }
    Ok(kind)
}

fn read_limit_text(
    format: &mut Text,
    position: &mut u16,
    symbol: &mut Text,
) -> Result<u16, FormatError> {
    let start = *position;
    let mut length = format.len() as u16;
    symbol.clear();
    while *position < length {
        let token = text::at(format, usize::from(*position))?;
        if token == b']' {
            break;
        }
        if token == b' ' {
            length -= erase_blanks(format, *position, length)?;
        } else {
            *position += 1;
            symbol.push(token);
        }
    }
    Ok(*position - start)
}

fn replace_section(
    format: &mut Text,
    start: u16,
    position: &mut u16,
    replacement: &[u8],
) -> Result<(), FormatError> {
    let count = count_from(i32::from(*position) - i32::from(start));
    text::erase(format, usize::from(start), count)?;
    text::insert(format, usize::from(start), replacement)?;
    *position = (usize::from(start) + replacement.len()) as u16;
    Ok(())
}

fn last_byte(format: &[u8]) -> Result<u8, FormatError> {
    format.last().copied().ok_or(FormatError)
}

impl NumberFormat {
    /// Parses an Excel number format.
    pub fn new(format: &str) -> Result<NumberFormat, FormatError> {
        const UNSCANNED: Section = Section {
            symbols: Vec::new(),
            info: ScanInfo::UNSCANNED,
        };
        let mut format: Text = format.as_bytes().to_vec();
        let mut scanner = FormatScanner::new();
        let mut sections = [UNSCANNED; 3];
        let mut limits = [0.0; 2];
        let mut op1 = LimitOp::None;
        let mut op2 = LimitOp::None;
        let mut format_type = FormatType::Defined;
        let mut has_condition = false;
        let mut failed = false;
        let mut cancel = false;
        let mut position = 0u16;
        let mut section_text = Text::new();
        let mut index = 0;

        while index < 4 && !cancel {
            let mut start = position;
            let mut symbol;
            // Bracketed prefixes come first: conditions, colours and the like.
            loop {
                symbol = next_bracket_symbol(&mut format, &mut position, &mut section_text)?;
                match symbol {
                    BracketSymbol::Condition(op) => {
                        if index == 0 && !has_condition {
                            has_condition = true;
                            op1 = op;
                        } else if index == 1 && has_condition {
                            op2 = op;
                        } else {
                            failed = true;
                            break;
                        }
                        let count = read_limit_text(&mut format, &mut position, &mut section_text)?;
                        let limit = if count > 0 {
                            parse_limit(&section_text)?
                        } else {
                            None
                        };
                        // Anything but a number is replaced by 0, also in the format code.
                        if limit.is_none() {
                            position -= count;
                            text::erase(&mut format, usize::from(position), usize::from(count))?;
                            text::insert_byte(&mut format, usize::from(position), b'0')?;
                            position = position.wrapping_add(1);
                        }
                        limits[index] = limit.unwrap_or(0.0);
                        if text::at(&format, usize::from(position))? != b']' {
                            failed = true;
                            break;
                        }
                        position = position.wrapping_add(1);
                        start = position;
                    }
                    BracketSymbol::Prefix => {
                        replace_section(&mut format, start, &mut position, &section_text)?;
                        text::insert_byte(&mut format, usize::from(position), b']')?;
                        text::insert_byte(&mut format, usize::from(start), b'[')?;
                        position = position.wrapping_add(2);
                        start = position;
                    }
                    BracketSymbol::Format | BracketSymbol::Error => break,
                }
            }

            if failed || symbol == BracketSymbol::Error {
                failed = true;
                cancel = true;
            } else {
                // Without conditions the sections are for positive, negative and zero values.
                if index == 1 && op1 == LimitOp::None {
                    op1 = LimitOp::Greater;
                } else if index == 2 && op2 == LimitOp::None {
                    op2 = LimitOp::Less;
                }
                if !section_text.is_empty() {
                    let result = scanner.scan(&mut section_text)?;
                    let count = scanner.result_count();
                    if result == 0 && count != 0 {
                        replace_section(&mut format, start, &mut position, &section_text)?;
                        if usize::from(position) < format.len() {
                            text::insert_byte(&mut format, usize::from(position), b';')?;
                            position = position.wrapping_add(1);
                        }
                        let info = scanner.info();
                        let symbols = scanner.take_symbols(count);
                        if index == 0 {
                            format_type = info.format_type;
                        } else if index < 3 && info.format_type != format_type {
                            format_type = FormatType::Defined;
                        }
                        if let Some(section) = sections.get_mut(index) {
                            *section = Section { symbols, info };
                        }
                    } else {
                        failed = true;
                        cancel = true;
                    }
                }
            }
            if format.len() == usize::from(position) {
                if index == 2 && symbol == BracketSymbol::Format {
                    let before = usize::from(position).checked_sub(1).ok_or(FormatError)?;
                    text::at(&format, before)?;
                }
                cancel = true;
            }
            index += 1;
        }

        // A section left out after a condition formats as General, unless it is explicitly empty.
        if has_condition && !failed {
            let general =
                if index == 1 && sections[0].symbols.is_empty() && last_byte(&format)? != b';' {
                    Some(0)
                } else if index == 1
                    && sections[1].symbols.is_empty()
                    && last_byte(&format)? != b';'
                    && (sections[0].symbols.len() > 1
                        || (sections[0].symbols.len() == 1
                            && sections[0].kind(0) != SymbolKind::Keyword(Keyword::General)))
                {
                    Some(1)
                } else if index == 2
                    && sections[2].symbols.is_empty()
                    && last_byte(&format)? != b';'
                    && op2 != LimitOp::None
                {
                    Some(2)
                } else {
                    None
                };
            if let Some(general) = general {
                let mut text = Keyword::General.text().to_vec();
                if scanner.scan(&mut text)? == 0 && scanner.result_count() != 0 {
                    let info = scanner.info();
                    let symbols = scanner.take_symbols(scanner.result_count());
                    sections[general] = Section { symbols, info };
                }
            }
        }

        let [limit1, limit2] = limits;
        if sections[2].symbols.is_empty()
            && op1 == LimitOp::Greater
            && op2 == LimitOp::None
            && limit1 == 0.0
            && limit2 == 0.0
        {
            op1 = LimitOp::GreaterEqual;
        }
        Ok(NumberFormat {
            sections,
            limit1,
            limit2,
            op1,
            op2,
            is_text: format_type == FormatType::Text,
        })
    }

    /// Formats `value`. Bytes the format cuts out of a multi-byte character come out as U+FFFD.
    pub fn format(&self, value: f64) -> Result<String, FormatError> {
        let bytes = self.format_bytes(value)?;
        Ok(String::from_utf8(bytes)
            .unwrap_or_else(|error| String::from_utf8_lossy(error.as_bytes()).into_owned()))
    }

    /// Formats `value` like [`NumberFormat::format`] but returns the raw bytes, which need not be
    /// valid UTF-8.
    pub fn format_bytes(&self, value: f64) -> Result<Vec<u8>, FormatError> {
        let mut number = value;
        if self.is_text {
            return standard(number);
        }
        let index = match self.op1.check(number, self.limit1) {
            None | Some(true) => 0,
            Some(false) => match self.op2.check(number, self.limit2) {
                None | Some(true) => 1,
                Some(false) => 2,
            },
        };
        // The sign is dropped where the section itself stands for "negative".
        let section_is_negative = match index {
            0 => self.first_section_is_negative(),
            1 => self.second_section_is_negative(),
            _ => false,
        };
        if section_is_negative && number < 0.0 {
            number = -number;
        }
        let section = &self.sections[index];
        let signed = index == 0;
        match section.info.format_type {
            _ if section.symbols.is_empty() => Ok(Text::new()),
            FormatType::Text | FormatType::Defined => Ok(literal_output(section)),
            FormatType::Date => date_output(section, number),
            FormatType::Time => time_output(section, number, signed),
            FormatType::DateTime => date_time_output(section, number),
            FormatType::Number | FormatType::Percent | FormatType::Currency => {
                number_output(section, number, signed)
            }
            FormatType::Fraction => fraction_output(section, number, signed),
            FormatType::Scientific => scientific_output(section, number, signed),
            FormatType::Undefined => Ok(Text::new()),
        }
    }

    fn limits_are_zero(&self) -> bool {
        self.limit1 == 0.0 && self.limit2 == 0.0
    }

    fn second_section_is_negative(&self) -> bool {
        use LimitOp::{Greater, GreaterEqual, Less, None};
        self.limits_are_zero()
            && matches!(
                (self.op1, self.op2),
                (GreaterEqual, None) | (Greater, Less) | (None, None)
            )
    }

    fn first_section_is_negative(&self) -> bool {
        use LimitOp::{Equal, Greater, GreaterEqual, Less, LessEqual, None};
        self.limits_are_zero()
            && matches!(
                (self.op1, self.op2),
                (Less, Greater | Equal | GreaterEqual | None) | (LessEqual, None | Greater)
            )
    }
}

fn standard(number: f64) -> Result<Text, FormatError> {
    if number.abs() > 1.0e15 {
        return Ok(double_to_string(number, Notation::Exponential, 2, false));
    }
    let text = double_to_string(number, Notation::Fixed, 2, true);
    // Only a negative number can come out empty.
    if text.is_empty() {
        return Err(FormatError);
    }
    Ok(text)
}

fn general(number: f64) -> Result<Text, FormatError> {
    let mut text = standard(number)?;
    text::erase_leading(&mut text, b'-');
    Ok(text)
}

fn push_int(text: &mut Text, value: i32) {
    text.extend_from_slice(value.to_string().as_bytes());
}

fn push_int2(text: &mut Text, value: i32) {
    if (0..10).contains(&value) {
        text.push(b'0');
    }
    push_int(text, value);
}

fn literal_output(section: &Section) -> Text {
    let mut output = Text::new();
    for symbol in &section.symbols {
        match symbol.kind {
            SymbolKind::Blank => output.push(b' '),
            SymbolKind::String | SymbolKind::Currency => output.extend_from_slice(&symbol.text),
            SymbolKind::ThousandSeparator if section.info.thousand_count == 0 => {
                output.extend_from_slice(&symbol.text);
            }
            _ => {}
        }
    }
    output
}

fn append_date_keyword(
    output: &mut Text,
    keyword: Keyword,
    date: &CalendarDate,
) -> Result<(), FormatError> {
    let month = locale::MONTHS.get(usize::from(date.month));
    let day = locale::DAYS_OF_WEEK.get(usize::from(date.day_of_week));
    let quarter = &locale::QUARTERS[usize::from(date.month / 3).min(3)];
    let year = i32::from(date.year as i16);
    let name = |name: Option<&locale::Name>, full: bool| {
        name.map_or("", |name| if full { name.full } else { name.abbreviation })
            .as_bytes()
    };
    match keyword {
        Keyword::Month => push_int(output, i32::from(date.month) + 1),
        Keyword::Month2 => push_int2(output, i32::from(date.month) + 1),
        Keyword::MonthAbbreviation => output.extend_from_slice(name(month, false)),
        Keyword::MonthName => output.extend_from_slice(name(month, true)),
        Keyword::MonthInitial => output.push(text::at(name(month, true), 0)?),
        Keyword::Quarter => output.extend_from_slice(quarter.abbreviation.as_bytes()),
        Keyword::QuarterName => output.extend_from_slice(quarter.full.as_bytes()),
        Keyword::Day => push_int(output, i32::from(date.day)),
        Keyword::Day2 => push_int2(output, i32::from(date.day)),
        Keyword::DayAbbreviation | Keyword::WeekdayAbbreviation | Keyword::AltDayAbbreviation => {
            output.extend_from_slice(name(day, false));
        }
        Keyword::DayName | Keyword::WeekdayName | Keyword::AltDayName => {
            output.extend_from_slice(name(day, true));
        }
        Keyword::WeekdayNameSeparator => {
            output.extend_from_slice(name(day, true));
            output.extend_from_slice(locale::LONG_DATE_DAY_OF_WEEK_SEPARATOR);
        }
        Keyword::Year2 | Keyword::EraYear => {
            if year < 100 {
                push_int(output, year);
            } else {
                push_int2(output, year % 100);
            }
        }
        Keyword::Year4 | Keyword::EraYear2 | Keyword::EraYearShort => push_int(output, year),
        Keyword::WeekOfYear => push_int(output, i32::from(date.week_of_year as i16)),
        Keyword::Era | Keyword::Era2 => {
            output.extend_from_slice(locale::ERA.abbreviation.as_bytes());
        }
        Keyword::EraName => output.extend_from_slice(locale::ERA.full.as_bytes()),
        Keyword::EraYearLong => {
            output.extend_from_slice(locale::ERA.full.as_bytes());
            push_int(output, year);
        }
        _ => {}
    }
    Ok(())
}

/// Hours, minutes and seconds as a time format shows them.
struct Clock {
    hour: u64,
    minute: u64,
    second: u64,
    is_am: bool,
    fraction: Text,
    fraction_digits: u16,
    fraction_position: usize,
}

impl Clock {
    fn new(time: f64, info: &ScanInfo) -> Clock {
        let seconds = to_u64(time.floor());
        let mut fraction = double_to_string(
            time - seconds as f64,
            Notation::Fixed,
            i32::from(info.count_post),
            false,
        );
        text::erase_leading(&mut fraction, b'0');
        text::erase_leading(&mut fraction, b'.');

        let (mut hour, minute, second) = match (info.thousand, info.thousand_count) {
            (false, _) => ((seconds / 3600) % 24, (seconds % 3600) / 60, seconds % 60),
            (true, 3) => (0, 0, seconds),
            (true, 2) => (0, seconds / 60, seconds % 60),
            (true, 1) => (seconds / 3600, (seconds % 3600) / 60, seconds % 60),
            (true, _) => (0, 0, 0),
        };
        let mut is_am = false;
        if info.count_exp != 0 {
            is_am = hour < 12;
            if hour == 0 {
                hour = 12;
            } else if hour > 12 {
                hour -= 12;
            }
        }
        Clock {
            hour,
            minute,
            second,
            is_am,
            fraction,
            fraction_digits: info.count_post,
            fraction_position: 0,
        }
    }

    fn append_fraction(
        &mut self,
        output: &mut Text,
        placeholder: &[u8],
    ) -> Result<(), FormatError> {
        for _ in 0..placeholder.len() {
            if self.fraction_position >= usize::from(self.fraction_digits) {
                break;
            }
            output.push(text::at(&self.fraction, self.fraction_position)?);
            self.fraction_position += 1;
        }
        Ok(())
    }

    fn append_keyword(&self, output: &mut Text, keyword: Keyword) {
        // Units are shown as 32-bit integers; elapsed times beyond that wrap around.
        match keyword {
            Keyword::AmPm => {
                output.extend_from_slice(if self.is_am {
                    locale::TIME_AM
                } else {
                    locale::TIME_PM
                });
            }
            Keyword::AP => output.push(if self.is_am { b'a' } else { b'p' }),
            Keyword::Minute => push_int(output, self.minute as i32),
            Keyword::Minute2 => push_int2(output, self.minute as i32),
            Keyword::Hour => push_int(output, self.hour as i32),
            Keyword::Hour2 => push_int2(output, self.hour as i32),
            Keyword::Second => push_int(output, self.second as i32),
            Keyword::Second2 => push_int2(output, self.second as i32),
            _ => {}
        }
    }
}

fn is_date_time_literal(kind: SymbolKind) -> bool {
    matches!(
        kind,
        SymbolKind::String
            | SymbolKind::Currency
            | SymbolKind::DateSeparator
            | SymbolKind::TimeSeparator
            | SymbolKind::Time100SecSeparator
    )
}

fn date_output(section: &Section, number: f64) -> Result<Text, FormatError> {
    let date = CalendarDate::from_epoch_days(number + EPOCH_OFFSET);
    let mut output = Text::new();
    for symbol in &section.symbols {
        match symbol.kind {
            SymbolKind::Blank => output.push(b' '),
            SymbolKind::Keyword(keyword) => append_date_keyword(&mut output, keyword, &date)?,
            kind if is_date_time_literal(kind) => output.extend_from_slice(&symbol.text),
            _ => {}
        }
    }
    Ok(output)
}

fn time_output(section: &Section, number: f64, signed: bool) -> Result<Text, FormatError> {
    let info = &section.info;
    let mut number = number;
    let mut negative = false;
    if number < 0.0 {
        number = -number;
        negative = signed;
    }
    if info.thousand {
        if number > 1.0e10 {
            return Ok(locale::ERROR_STRING.to_vec());
        }
    } else {
        number -= number.floor();
        // A negative time of day counts back from midnight.
        if negative {
            number = 1.0 - number;
        }
    }
    let time = round_corrected(number * 86400.0, i32::from(info.count_post));
    if time == 0.0 {
        negative = false;
    }
    if time.floor() > MAX_U32 {
        return Ok(locale::ERROR_STRING.to_vec());
    }
    let mut clock = Clock::new(time, info);
    let mut output = Text::new();
    for symbol in &section.symbols {
        match symbol.kind {
            SymbolKind::Blank => output.push(b' '),
            SymbolKind::Digit => clock.append_fraction(&mut output, &symbol.text)?,
            SymbolKind::Keyword(keyword) => clock.append_keyword(&mut output, keyword),
            kind if is_date_time_literal(kind) => output.extend_from_slice(&symbol.text),
            _ => {}
        }
    }
    if negative && info.thousand {
        output.insert(0, b'-');
    }
    Ok(output)
}

fn date_time_output(section: &Section, number: f64) -> Result<Text, FormatError> {
    let info = &section.info;
    let mut number = number + EPOCH_OFFSET;
    let mut time = round_corrected(
        (number - number.floor()) * 86400.0,
        i32::from(info.count_post),
    );
    // A time rounded up to midnight belongs to the next day.
    if time >= 86400.0 {
        time -= 86400.0;
        number = (number + 0.5).floor() + time;
    }
    let date = CalendarDate::from_epoch_days(number);
    let mut clock = Clock::new(time, info);
    let mut output = Text::new();
    for symbol in &section.symbols {
        match symbol.kind {
            SymbolKind::Blank => output.push(b' '),
            SymbolKind::Digit => clock.append_fraction(&mut output, &symbol.text)?,
            SymbolKind::Keyword(keyword) => {
                clock.append_keyword(&mut output, keyword);
                append_date_keyword(&mut output, keyword, &date)?;
            }
            kind if is_date_time_literal(kind) => output.extend_from_slice(&symbol.text),
            _ => {}
        }
    }
    Ok(output)
}

fn digit_fill(
    digits: &mut Text,
    start: u16,
    position: &mut u16,
    info: &ScanInfo,
    digit_count: &mut u16,
    grouping: &mut DigitGrouping,
) -> Result<(), FormatError> {
    if !info.thousand {
        *position = start;
        return Ok(());
    }
    while *position > start {
        if i32::from(*digit_count) == grouping.position() {
            text::insert_byte(digits, usize::from(*position), locale::THOUSAND_SEPARATOR)?;
            grouping.advance();
        }
        *digit_count = digit_count.wrapping_add(1);
        *position -= 1;
    }
    Ok(())
}

fn fill_with_thousands(
    section: &Section,
    digits: &mut Text,
    number: f64,
    mut position: u16,
    mut index: usize,
    wanted_digits: u16,
) -> Result<(), FormatError> {
    let info = &section.info;
    let mut leading_chars = 0u16;
    let mut digit_count = 0u16;
    let mut do_thousands = info.thousand_count == 0;
    let mut grouping = DigitGrouping::new();
    loop {
        let kind = section.kind(index);
        let symbol = section.text(index);
        match kind {
            SymbolKind::DecimalSeparator
            | SymbolKind::String
            | SymbolKind::Currency
            | SymbolKind::Percent => {
                if kind == SymbolKind::DecimalSeparator {
                    grouping.reset();
                }
                text::insert(digits, usize::from(position), symbol)?;
                if position == 0 {
                    leading_chars = leading_chars.wrapping_add(symbol.len() as u16);
                }
            }
            SymbolKind::Blank => text::insert_byte(digits, usize::from(position), b' ')?,
            SymbolKind::ThousandSeparator => {
                if !do_thousands && index + 1 < section.symbols.len() {
                    do_thousands = section.shows_separator(index);
                }
                if do_thousands {
                    if position > 0 {
                        text::insert(digits, usize::from(position), symbol)?;
                    } else if digit_count < wanted_digits {
                        // In front of the number a leading # shows nothing and a ? a blank.
                        let leader = if index > 0 && section.kind(index - 1) == SymbolKind::Digit {
                            section.text(index - 1).last().copied()
                        } else {
                            None
                        };
                        match leader {
                            Some(b'#') => {}
                            Some(b'?') => text::insert_byte(digits, 0, b' ')?,
                            _ => text::insert(digits, 0, symbol)?,
                        }
                    }
                    grouping.advance();
                }
            }
            SymbolKind::Digit => {
                for &placeholder in symbol.iter().rev() {
                    digit_count = digit_count.wrapping_add(1);
                    if position > 0 {
                        position -= 1;
                    } else {
                        match placeholder {
                            b'0' => digits.insert(0, b'0'),
                            b'?' => digits.insert(0, b' '),
                            _ => {}
                        }
                    }
                    if digit_count == wanted_digits && position > 0 {
                        digit_fill(
                            digits,
                            0,
                            &mut position,
                            info,
                            &mut digit_count,
                            &mut grouping,
                        )?;
                    }
                }
            }
            SymbolKind::Keyword(Keyword::BankSymbol) => {
                text::insert(digits, usize::from(position), locale::CURRENCY_BANK_SYMBOL)?;
            }
            SymbolKind::Keyword(Keyword::General) => {
                text::insert(digits, usize::from(position), &general(number)?)?;
            }
            _ => {}
        }
        if index == 0 {
            break;
        }
        index -= 1;
    }
    position = position.wrapping_add(leading_chars);
    if position > leading_chars {
        digit_fill(
            digits,
            leading_chars,
            &mut position,
            info,
            &mut digit_count,
            &mut grouping,
        )?;
    }
    Ok(())
}

fn number_fill(
    section: &Section,
    digits: &mut Text,
    number: f64,
    position: &mut u16,
    index: &mut usize,
    stop: SymbolKind,
) -> Result<(), FormatError> {
    *position = digits.len() as u16;
    let mut do_thousands = section.info.thousand_count == 0;
    while *index > 0 {
        let kind = section.kind(*index);
        if kind == stop {
            break;
        }
        let symbol = section.text(*index);
        match kind {
            SymbolKind::Star => {}
            SymbolKind::Blank => {
                text::insert_byte(digits, usize::from(*position), b' ')?;
                *position = position.wrapping_add(1);
            }
            SymbolKind::ThousandSeparator => {
                if !do_thousands && *index + 1 < section.symbols.len() {
                    do_thousands = section.shows_separator(*index);
                }
                if do_thousands && *position > 0 {
                    text::insert(digits, usize::from(*position), symbol)?;
                }
            }
            SymbolKind::Digit => {
                for &placeholder in symbol.iter().rev() {
                    if *position > 0 {
                        *position -= 1;
                    } else {
                        match placeholder {
                            b'0' => digits.insert(0, b'0'),
                            b'?' => digits.insert(0, b' '),
                            _ => {}
                        }
                    }
                }
            }
            SymbolKind::Keyword(Keyword::BankSymbol) => {
                text::insert(digits, usize::from(*position), locale::CURRENCY_BANK_SYMBOL)?;
            }
            SymbolKind::Keyword(Keyword::General) => {
                text::insert(digits, usize::from(*position), &general(number)?)?;
            }
            _ => text::insert(digits, usize::from(*position), symbol)?,
        }
        *index -= 1;
    }
    Ok(())
}

fn number_output(section: &Section, number: f64, signed: bool) -> Result<Text, FormatError> {
    let info = &section.info;
    let mut number = number;
    let mut negative = false;
    if number < 0.0 {
        negative = signed;
        number = -number;
    }
    if info.format_type == FormatType::Percent {
        if number < 1.7e306 {
            number *= 100.0;
        } else {
            return Ok(locale::ERROR_STRING.to_vec());
        }
    }

    let mut digits = Text::new();
    let mut is_integer = false;
    if info.thousand_count != GENERAL_IN_FORMAT {
        for _ in 0..info.thousand_count {
            number = if number > 2.3e-305 {
                number / 1000.0
            } else {
                0.0
            };
        }
        let integer_digits = if number > 0.0 {
            integer_digits(number)
        } else {
            0
        };
        let decimals = i64::from(info.count_post);
        if decimals != 0 {
            // Not more than 15 significant digits; the rest of the decimals are zeros.
            let significant = 15i64.saturating_sub(integer_digits);
            if integer_digits < 15 && significant < decimals {
                digits = double_to_string(number, Notation::Fixed, significant as i32, false);
                digits.resize(digits.len() + (decimals - significant) as usize, b'0');
            } else {
                digits = double_to_string(number, Notation::Fixed, decimals as i32, false);
            }
            text::erase_leading(&mut digits, b'0');
        } else if number != 0.0 {
            digits = double_to_string(number, Notation::Fixed, 0, false);
            text::erase_leading(&mut digits, b'0');
        }
        if let Some(point) = digits.iter().position(|&byte| byte == b'.') {
            is_integer = digits[point + 1..].iter().all(|&byte| byte == b'0');
            digits.remove(point);
        }
        if digits.is_empty() {
            negative = false;
        }
    }

    // Decimals first, from the last symbol back to the decimal separator.
    let mut position = digits.len() as u16;
    let mut index = section.symbols.len().saturating_sub(1);
    if info.count_post > 0 {
        let mut trailing = true;
        let mut filled = false;
        while index > 0 {
            let kind = section.kind(index);
            if kind == SymbolKind::DecimalSeparator {
                break;
            }
            let symbol = section.text(index);
            let at = usize::from(position);
            match kind {
                SymbolKind::Blank => text::insert_byte(&mut digits, at, b' ')?,
                SymbolKind::String | SymbolKind::Currency | SymbolKind::Percent => {
                    text::insert(&mut digits, at, symbol)?;
                }
                SymbolKind::ThousandSeparator if info.thousand_count == 0 => {
                    text::insert(&mut digits, at, symbol)?;
                }
                SymbolKind::Digit => {
                    for &placeholder in symbol.iter().rev() {
                        position = position.wrapping_sub(1);
                        let at = usize::from(position);
                        if text::at(&digits, at)? != b'0' {
                            trailing = false;
                        }
                        if !trailing {
                            continue;
                        }
                        // Trailing zeros stay for 0, become a blank for ? and go away for #.
                        match placeholder {
                            b'0' => filled = true,
                            b'-' => {
                                if is_integer {
                                    digits[at] = b'-';
                                }
                                filled = true;
                            }
                            b'?' => {
                                digits[at] = b' ';
                                filled = true;
                            }
                            _ if !filled => text::erase(&mut digits, at, 1)?,
                            _ => {}
                        }
                    }
                }
                SymbolKind::Keyword(Keyword::BankSymbol) => {
                    text::insert(&mut digits, at, locale::CURRENCY_BANK_SYMBOL)?;
                }
                SymbolKind::Keyword(Keyword::General) => {
                    text::insert(&mut digits, at, &general(number)?)?;
                }
                _ => {}
            }
            index -= 1;
        }
    }

    fill_with_thousands(
        section,
        &mut digits,
        number,
        position,
        index,
        info.count_pre,
    )?;
    if info.count_post > 0 && digits.len() > 1 && digits.last() == Some(&locale::DECIMAL_SEPARATOR)
    {
        digits.pop();
    }
    if negative {
        digits.insert(0, b'-');
    }
    Ok(digits)
}

fn gcd(mut x: u64, mut y: u64) -> u64 {
    if y == 0 {
        return x;
    }
    let mut remainder = x % y;
    while remainder != 0 {
        x = y;
        y = remainder;
        remainder = x % y;
    }
    y
}

fn approximate_gcd(mut x: u64, mut y: u64) -> u64 {
    if y == 0 {
        return x;
    }
    let mut remainder = x % y;
    while remainder as f64 / y as f64 > 1.0e-2 {
        x = y;
        y = remainder;
        remainder = x % y;
    }
    y
}

fn divide(x: u64, y: u64) -> Result<u64, FormatError> {
    x.checked_div(y).ok_or(FormatError)
}

fn ratio(x: u64, y: u64) -> f64 {
    x as f64 / y as f64
}

fn farey_fraction(number: &mut f64, basis: u64) -> Result<(u64, u64, bool), FormatError> {
    // The search does not end in practice for NaN.
    if number.is_nan() {
        return Err(FormatError);
    }
    let upper_half = *number > 0.5;
    if upper_half {
        *number -= (*number - 0.5) * 2.0;
    }
    let number = *number;

    // Two neighbours in the series enclosing the number, starting from denominator `basis`.
    let mut x0 = to_u64((number * basis as f64).floor());
    let mut y0;
    let mut x1;
    let mut y1;
    if x0 == 0 {
        y0 = 1;
        x1 = 1;
        y1 = basis;
    } else if x0 == basis.wrapping_sub(1) / 2 {
        y0 = basis;
        x1 = 1;
        y1 = 2;
    } else if x0 == 1 {
        y0 = basis;
        x1 = 1;
        y1 = basis.wrapping_sub(1);
    } else {
        y0 = basis;
        x1 = x0;
        y1 = basis.wrapping_sub(1);
        let lower = ratio(x0, y0);
        let mut upper = ratio(x1, y1);
        let divisor = gcd(y0, x0);
        x0 = divide(x0, divisor)?;
        y0 = divide(y0, divisor)?;
        let mut x2;
        let mut y2;
        loop {
            let mut test = ratio(x1, y1);
            loop {
                while test > upper {
                    x1 = x1.wrapping_sub(1);
                    test = ratio(x1, y1);
                }
                while test < lower && y1 > 1 {
                    y1 -= 1;
                    test = ratio(x1, y1);
                }
                if test <= upper {
                    upper = test;
                    break;
                } else if y1 == 1 {
                    break;
                }
            }
            let divisor = gcd(y1, x1);
            x2 = divide(x1, divisor)?;
            y2 = divide(y1, divisor)?;
            if x2.wrapping_mul(y0).wrapping_sub(x0.wrapping_mul(y2)) == 1 || y1 <= 1 {
                break;
            }
            y1 -= 1;
        }
        x1 = x2;
        y1 = y2;
    }

    let mut low = ratio(x0, y0);
    let mut up = ratio(x1, y1);
    // The series has fewer than basis² terms; the bound only guards against an endless walk.
    let mut steps = basis.saturating_mul(basis);
    while number > up {
        steps = steps.checked_sub(1).ok_or(FormatError)?;
        let factor = divide(y0.wrapping_add(basis), y1)?;
        let x2 = factor.wrapping_mul(x1).wrapping_sub(x0);
        let y2 = factor.wrapping_mul(y1).wrapping_sub(y0);
        x0 = x1;
        y0 = y1;
        x1 = x2;
        y1 = y2;
        low = up;
        up = ratio(x1, y1);
    }
    let (mut numerator, denominator) = if number - low < up - number {
        (x0, y0)
    } else {
        (x1, y1)
    };
    let mut rounds_to_one = false;
    if upper_half {
        if numerator == 0 && denominator == 1 {
            rounds_to_one = true;
        } else {
            numerator = denominator.wrapping_sub(numerator);
        }
    }
    Ok((numerator, denominator, rounds_to_one))
}

fn decimal_fraction(
    number: f64,
    basis: u64,
    denominator_digits: u16,
) -> Result<(u64, u64), FormatError> {
    let mut denominator = 10_000_000;
    let mut numerator = to_u64((0.5 + number * 10_000_000.0).floor());
    let divisor = gcd(denominator, numerator);
    if divisor > 1 {
        denominator = divide(denominator, divisor)?;
        numerator = divide(numerator, divisor)?;
    }
    if denominator > basis {
        let divisor = approximate_gcd(denominator, numerator);
        if divisor > 1 {
            denominator = divide(denominator, divisor)?;
            numerator = divide(numerator, divisor)?;
        }
    }
    if denominator > basis {
        denominator = basis;
        let scale = pow(10.0, f64::from(denominator_digits));
        numerator = to_u64((0.5 + number * scale).floor());
        let divisor = approximate_gcd(denominator, numerator);
        if divisor > 1 {
            denominator = divide(denominator, divisor)?;
            numerator = divide(numerator, divisor)?;
        }
    }
    Ok((numerator, denominator))
}

fn fraction_output(section: &Section, number: f64, signed: bool) -> Result<Text, FormatError> {
    let info = &section.info;
    let mut number = number;
    let mut negative = false;
    if number < 0.0 {
        negative = signed;
        number = -number;
    }
    let mut whole = number.floor();
    number -= whole;
    if whole > MAX_U32 || info.count_exp > 9 {
        return Ok(locale::ERROR_STRING.to_vec());
    }
    if info.count_exp == 0 {
        return Ok(Text::new());
    }
    let basis = to_u64(pow(10.0, f64::from(info.count_exp)).floor()).wrapping_sub(1);
    let mut numerator;
    let mut denominator;
    if info.exp_val != 0 {
        denominator = u64::from(info.exp_val);
        let scaled = number * denominator as f64;
        numerator = to_u64(scaled.floor());
        if scaled - numerator as f64 >= 0.5 {
            numerator = numerator.wrapping_add(1);
        }
        if numerator == denominator {
            whole += 1.0;
            denominator = 1;
            numerator = 0;
        }
    } else if info.count_exp <= 3 {
        let rounds_to_one;
        (numerator, denominator, rounds_to_one) = farey_fraction(&mut number, basis)?;
        if rounds_to_one {
            whole += 1.0;
        }
    } else {
        (numerator, denominator) = decimal_fraction(number, basis, info.count_exp)?;
    }

    let mut integer = Text::new();
    if info.count_pre == 0 {
        // No integer part in the format: an improper fraction.
        let total = whole * denominator as f64 + numerator as f64;
        if total > MAX_U32 {
            return Ok(locale::ERROR_STRING.to_vec());
        }
        numerator = to_u64(total.floor());
    } else if !(whole == 0.0 && numerator != 0) {
        // The sign bit shows for -0 and NaN, the only values that can still carry it.
        if whole.is_sign_negative() {
            integer.push(b'-');
        }
        if whole.is_nan() {
            integer.extend_from_slice(b"nan");
        } else {
            integer.extend_from_slice(to_u64(whole).to_string().as_bytes());
        }
    }
    let mut numerator_text = Text::new();
    let mut denominator_text = Text::new();
    let no_fraction = info.count_pre > 0 && numerator == 0;
    if !no_fraction {
        push_int(&mut numerator_text, numerator as i32);
        push_int(&mut denominator_text, denominator as i32);
    }

    // From the last symbol backwards: denominator, numerator, integer part.
    let mut index = section.symbols.len().saturating_sub(1);
    let mut position = 0;
    let mut more = true;
    number_fill(
        section,
        &mut denominator_text,
        number,
        &mut position,
        &mut index,
        SymbolKind::Fraction,
    )?;
    if section.kind(index) == SymbolKind::Fraction {
        let slash = if no_fraction {
            b' '
        } else {
            text::at(section.text(index), 0)?
        };
        denominator_text.insert(0, slash);
        more = index > 0;
        index = index.saturating_sub(1);
    }
    if !more {
        numerator_text.clear();
    } else {
        number_fill(
            section,
            &mut numerator_text,
            number,
            &mut position,
            &mut index,
            SymbolKind::FractionBlank,
        )?;
        if section.kind(index) == SymbolKind::FractionBlank {
            text::insert(&mut numerator_text, 0, section.text(index))?;
            more = index > 0;
            index = index.saturating_sub(1);
        }
    }
    if !more {
        integer.clear();
    } else if !integer.is_empty() {
        let position = integer.len() as u16;
        fill_with_thousands(
            section,
            &mut integer,
            number,
            position,
            index,
            info.count_pre,
        )?;
    }

    let mut output = Text::new();
    if negative && !(numerator == 0 && whole == 0.0) {
        output.push(b'-');
    }
    output.extend_from_slice(&integer);
    output.extend_from_slice(&numerator_text);
    output.extend_from_slice(&denominator_text);
    Ok(output)
}

fn scientific_output(section: &Section, number: f64, signed: bool) -> Result<Text, FormatError> {
    let info = &section.info;
    let mut number = number;
    let mut negative = false;
    if number < 0.0 {
        negative = signed;
        number = -number;
    }
    let significant = i32::from(info.count_pre) + i32::from(info.count_post);
    let mut mantissa = double_to_string(number, Notation::Exponential, significant - 1, false);
    let mut exponent = Text::new();
    let mut exponent_is_negative = false;
    if let Some(e_position) = mantissa.iter().position(|&byte| byte == b'E') {
        let mut start = e_position + 1;
        match text::at(&mantissa, start)? {
            b'-' => {
                exponent_is_negative = true;
                start += 1;
            }
            b'+' => start += 1,
            _ => {}
        }
        exponent.extend_from_slice(text::tail(&mantissa, start)?);
        mantissa.truncate(e_position);
        text::erase_all(&mut mantissa, b'.');
        // More than one integer digit shifts the exponent.
        if info.count_pre != 1 {
            let mut value = exponent.iter().fold(0i32, |value, &digit| {
                value * 10 + i32::from(digit.wrapping_sub(b'0'))
            });
            if exponent_is_negative {
                value = -value;
            }
            value -= i32::from(info.count_pre) - 1;
            exponent_is_negative = value < 0;
            exponent.clear();
            push_int(&mut exponent, value.abs());
        }
    }

    let mut index = section.symbols.len().saturating_sub(1);
    let mut position = 0;
    number_fill(
        section,
        &mut exponent,
        number,
        &mut position,
        &mut index,
        SymbolKind::Exponent,
    )?;
    let mut zeros = 0;
    while zeros < usize::from(position) && text::at(&exponent, zeros)? == b'0' {
        zeros += 1;
    }
    exponent.drain(..zeros);

    let mut more = true;
    if section.kind(index) == SymbolKind::Exponent {
        let symbol = section.text(index);
        if exponent_is_negative {
            exponent.insert(0, b'-');
        } else if symbol.get(1) == Some(&b'+') {
            exponent.insert(0, b'+');
        }
        exponent.insert(0, text::at(symbol, 0)?);
        more = index > 0;
        index = index.saturating_sub(1);
    }
    if !more {
        mantissa.clear();
    } else {
        let position = mantissa.len() as u16;
        fill_with_thousands(
            section,
            &mut mantissa,
            number,
            position,
            index,
            significant as u16,
        )?;
    }
    if negative {
        mantissa.insert(0, b'-');
    }
    mantissa.extend_from_slice(&exponent);
    Ok(mantissa)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn format(format: &str, value: f64) -> Result<String, FormatError> {
        NumberFormat::new(format)?.format(value)
    }

    #[test]
    fn one_parsed_format_formats_many_values() {
        let format = NumberFormat::new("#,##0.00;(#,##0.00);\"zero\"").unwrap();
        assert_eq!(format.format(1234.5).unwrap(), "1,234.50");
        assert_eq!(format.format(-1234.5).unwrap(), "(1,234.50)");
        assert_eq!(format.format(0.0).unwrap(), "zero");
    }

    #[test]
    fn conditions_select_the_section() {
        let format = NumberFormat::new("[<100]\"small\";[>=1000]\"big\";0").unwrap();
        assert_eq!(format.format(99.0).unwrap(), "small");
        assert_eq!(format.format(1000.0).unwrap(), "big");
        assert_eq!(format.format(500.0).unwrap(), "500");
    }

    #[test]
    fn new_fails_on_format_codes_the_scanner_cannot_read() {
        for code in [
            "[$",
            "E",
            "0E",
            "x/",
            "[>5",
            "# ?/99999999999",
            "[>99999999999:00]0",
        ] {
            assert_eq!(NumberFormat::new(code).err(), Some(FormatError), "{code}");
        }
    }

    #[test]
    fn format_fails_on_values_without_result() {
        assert_eq!(format("# ?/?", f64::NAN), Err(FormatError));
        assert_eq!(format("@", -48.5), Err(FormatError));
        assert_eq!(format("0.0555", 1.5), Err(FormatError));
        assert_eq!(format("h:mm:ss.0000", f64::NAN), Err(FormatError));
    }

    #[test]
    fn unusable_sections_show_nothing() {
        assert_eq!(format("#,##0".to_string().as_str(), 1.0).unwrap(), "1");
        assert_eq!(
            format(&format!("#,##0{}", "0".repeat(100)), 1.0).unwrap(),
            ""
        );
        assert_eq!(format("0.00;[", -1.0).unwrap(), "");
        assert_eq!(format("0.00;[", 1.0).unwrap(), "1.00");
    }

    #[test]
    fn digits_of_huge_values_are_written_out() {
        let expected = format!("1{}", "0".repeat(300));
        assert_eq!(format("0", 1e300).unwrap(), expected);
        let expected = format!("17976931348623200{}.00", "0".repeat(292));
        assert_eq!(format("0.00", f64::MAX).unwrap(), expected);
    }

    #[test]
    fn non_finite_values_format() {
        assert_eq!(format("0.00", f64::NAN).unwrap(), "N.aN");
        assert_eq!(format("0.00", f64::NEG_INFINITY).unwrap(), "-I.NF");
        assert_eq!(format("General", f64::INFINITY).unwrap(), "INF");
        assert_eq!(format("yyyy-mm-dd", f64::NAN).unwrap(), "0-01-01");
        assert_eq!(format("hh:mm:ss", f64::NAN).unwrap(), "15:30:08");
        assert_eq!(format("0%", f64::INFINITY).unwrap(), "###");
    }

    #[test]
    fn text_passes_through_as_bytes() {
        assert_eq!(format("€", 1234.567).unwrap(), "€");
        assert_eq!(format("0.0 \"日本語\"", 1.25).unwrap(), "1.3 日本語");
        assert_eq!(format("yyyy\"年\"m\"月\"", 45000.0).unwrap(), "2023年3月");
    }

    #[test]
    fn split_multi_byte_characters_are_replaced_in_strings_only() {
        let format = NumberFormat::new("_€0").unwrap();
        assert_eq!(format.format_bytes(5.0).unwrap(), b" \x82\xAC5");
        assert_eq!(format.format(5.0).unwrap(), " \u{FFFD}\u{FFFD}5");
    }
}
