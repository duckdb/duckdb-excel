//! Scanner for one section of a format code.
//!
//! A section is scanned in three passes: `divide` splits it into symbols, `scan_type` derives the
//! format type from them and `final_scan` gives every symbol its final kind for that type and
//! rebuilds the section text from the symbols.
//!
//! The passes report a problem as the position where it was found. Position 0 means "no problem",
//! so a problem at the very start of a section goes unnoticed and leaves the section half
//! scanned; such sections still format in a defined way, which is why positions are tracked at
//! all. The parts of the final scan return `true` for a problem at the current position.
//!
//! At most [`MAX_SYMBOLS`] symbols are scanned, the rest of a longer section is dropped. Group
//! separators are inserted as extra symbols, and a section that has no room for them fails.
//!
//! The rebuilt text only matters for its length, with one exception: in currency formats the
//! currency symbol of a quoted literal is written back at its offset in the unquoted literal,
//! where the rebuilt text can hold another character.

use crate::FormatError;
use crate::locale;
use crate::symbol::{FormatType, Keyword, Symbol, SymbolKind};
use crate::text::{self, Text};

const MAX_SYMBOLS: usize = 100;

/// Value of the thousands count that marks a `General` keyword in a number format.
pub(crate) const GENERAL_IN_FORMAT: u16 = 1000;

/// Positions from the right at which group separators go: every third digit.
pub(crate) struct DigitGrouping {
    position: i32,
}

impl DigitGrouping {
    pub(crate) fn new() -> DigitGrouping {
        DigitGrouping { position: 3 }
    }

    pub(crate) fn position(&self) -> i32 {
        self.position
    }

    pub(crate) fn advance(&mut self) {
        self.position = self.position.saturating_add(3);
    }

    pub(crate) fn reset(&mut self) {
        self.position = 3;
    }
}

/// The numbers a scan derives besides the symbols.
#[derive(Clone, Copy, Debug)]
pub(crate) struct ScanInfo {
    pub(crate) format_type: FormatType,
    /// Group separators are wanted; for time formats: has a `[h]`-like elapsed time part.
    pub(crate) thousand: bool,
    /// Number of trailing group separators that scale the value; for time formats the unit of the
    /// elapsed time part (1 hours, 2 minutes, 3 seconds); [`GENERAL_IN_FORMAT`] for `General`.
    pub(crate) thousand_count: u16,
    /// Digits before the decimal separator.
    pub(crate) count_pre: u16,
    /// Digits after the decimal separator.
    pub(crate) count_post: u16,
    /// Digits of the exponent or of the denominator; 1 if a time format has AM/PM.
    pub(crate) count_exp: u16,
    /// A fixed denominator, 0 if none.
    pub(crate) exp_val: u32,
}

impl ScanInfo {
    pub(crate) const UNSCANNED: ScanInfo = ScanInfo {
        format_type: FormatType::Undefined,
        thousand: false,
        thousand_count: 0,
        count_pre: 0,
        count_post: 0,
        count_exp: 0,
        exp_val: 0,
    };
}

/// Scanner state; reusable for any number of sections.
///
/// `result_count` is the number of symbols not merged into others, 0 until the final scan.
/// `has_currency` tells whether the section has a currency symbol outside of quotes.
pub(crate) struct FormatScanner {
    symbols: Vec<Symbol>,
    result_count: u16,
    info: ScanInfo,
    exponent_position: usize,
    blank_position: usize,
    decimal_position: usize,
    has_exponent: bool,
    has_fraction: bool,
    has_blank: bool,
    has_decimal_separator: bool,
    has_currency: bool,
}

fn is_digit_placeholder(byte: u8) -> bool {
    matches!(byte, b'#' | b'0' | b'?')
}

fn has_unquoted_currency(format: &[u8]) -> bool {
    let mut quoted = false;
    let mut previous = None;
    for &byte in format {
        if byte == b'"' {
            quoted = match previous {
                None => true,
                Some(previous) if quoted => previous == 0,
                Some(previous) => previous != b'\\',
            };
        } else if byte == locale::CURRENCY_SYMBOL
            && !quoted
            && !matches!(previous, Some(b'"' | b'\\'))
        {
            return true;
        }
        previous = Some(byte);
    }
    false
}

fn remove_quotes(text: &mut Text) -> bool {
    if text.len() < 2 {
        return false;
    }
    if text[0] == b'"' && text[text.len() - 1] == b'"' {
        text.pop();
        text.remove(0);
        true
    } else if text[0] == b'\\' {
        text.remove(0);
        true
    } else {
        false
    }
}

fn next_symbol(format: &[u8], position: &mut usize) -> Result<Symbol, FormatError> {
    enum State {
        Start,
        Escaped,
        Quoted,
        Word,
        Second,
    }
    let mut kind = SymbolKind::None;
    let mut state = State::Start;
    let mut symbol = Text::new();
    while *position < format.len() {
        let start = *position;
        let token = format[start];
        *position += 1;
        match state {
            State::Start => match token {
                b'#' | b'0' | b'?' | b'%' | b'@' | b'[' | b']' | b',' | b'.' | b'/' | b'\''
                | b' ' | b':' | b'-' => {
                    kind = SymbolKind::Delimiter;
                    symbol.push(token);
                    break;
                }
                b'*' => {
                    kind = SymbolKind::Star;
                    symbol.push(token);
                    state = State::Second;
                }
                b'_' => {
                    kind = SymbolKind::Blank;
                    symbol.push(token);
                    state = State::Second;
                }
                b'"' => {
                    kind = SymbolKind::String;
                    symbol.push(token);
                    state = State::Quoted;
                }
                b'\\' => {
                    kind = SymbolKind::String;
                    symbol.push(token);
                    state = State::Escaped;
                }
                b'$' | b'+' | b'(' | b')' => {
                    kind = SymbolKind::String;
                    symbol.push(token);
                    break;
                }
                _ if token.is_ascii_alphabetic() => {
                    let Some(mut keyword) = Keyword::at_start_of(&format[start..]) else {
                        symbol.push(token);
                        state = State::Word;
                        continue;
                    };
                    let length = keyword.text().len();
                    symbol.extend_from_slice(&format[start..start + length]);
                    *position = start + length;
                    // An E is the exponent of a scientific format if a sign or a digit follows.
                    if keyword == Keyword::EraYear {
                        match text::at(format, *position)? {
                            sign @ (b'+' | b'-') => {
                                symbol.push(sign);
                                keyword = Keyword::Exponent;
                                *position += 1;
                            }
                            b'0' | b'#' => keyword = Keyword::Exponent,
                            _ => {}
                        }
                    }
                    kind = SymbolKind::Keyword(keyword);
                    break;
                }
                _ => {
                    kind = if token.is_ascii_digit() {
                        SymbolKind::Digit
                    } else {
                        SymbolKind::String
                    };
                    symbol.push(token);
                    break;
                }
            },
            State::Escaped | State::Second => {
                symbol.push(token);
                break;
            }
            State::Quoted => {
                symbol.push(token);
                if token == b'"' {
                    break;
                }
            }
            State::Word => {
                kind = SymbolKind::String;
                if token.is_ascii_alphabetic() {
                    if Keyword::at_start_of(&format[start..]).is_some() {
                        *position = start;
                        break;
                    }
                    symbol.push(token);
                    continue;
                }
                // The slash of a split "A/P" or "AM/PM" stays in the word.
                if token == b'/' && matches!(text::at(format, *position)?, b'P' | b'p') {
                    let is_am_pm = match symbol.as_slice() {
                        [b'A' | b'a'] => true,
                        [b'A' | b'a', b'M' | b'm'] => {
                            matches!(text::at(format, *position + 1)?, b'M' | b'm')
                        }
                        _ => false,
                    };
                    if is_am_pm {
                        symbol.push(token);
                        continue;
                    }
                }
                *position = start;
                break;
            }
        }
    }
    if matches!(state, State::Word) {
        kind = SymbolKind::String;
    }
    Ok(Symbol { kind, text: symbol })
}

impl FormatScanner {
    pub(crate) fn new() -> FormatScanner {
        FormatScanner {
            symbols: Vec::new(),
            result_count: 0,
            info: ScanInfo::UNSCANNED,
            exponent_position: 0,
            blank_position: 0,
            decimal_position: 0,
            has_exponent: false,
            has_fraction: false,
            has_blank: false,
            has_decimal_separator: false,
            has_currency: false,
        }
    }

    /// Scans one section and rebuilds `format` from its symbols; anything but 0 is a failure.
    pub(crate) fn scan(&mut self, format: &mut Text) -> Result<u16, FormatError> {
        let mut result = self.divide(format)?;
        if result == 0 {
            result = self.scan_type()?;
        }
        if result == 0 {
            result = self.final_scan(format)?;
        }
        Ok(result)
    }

    /// Number of symbols the last scan produced, 0 if it did not get to the final scan.
    pub(crate) fn result_count(&self) -> u16 {
        self.result_count
    }

    pub(crate) fn info(&self) -> ScanInfo {
        self.info
    }

    /// The first `count` symbols of the last scan that were not merged into others.
    pub(crate) fn take_symbols(&mut self, count: u16) -> Vec<Symbol> {
        let mut symbols = std::mem::take(&mut self.symbols);
        symbols.retain(|symbol| symbol.kind != SymbolKind::Empty);
        // Unused slots of the symbol table make up for a count above the symbols found.
        let unused = MAX_SYMBOLS.saturating_sub(symbols.len());
        let count = usize::from(count).min(symbols.len() + unused);
        symbols.resize(
            count,
            Symbol {
                kind: SymbolKind::None,
                text: Text::new(),
            },
        );
        symbols
    }

    fn count(&self) -> usize {
        self.symbols.len()
    }

    fn kind(&self, index: usize) -> SymbolKind {
        self.symbols
            .get(index)
            .map_or(SymbolKind::None, |symbol| symbol.kind)
    }

    fn keyword(&self, index: usize) -> Option<Keyword> {
        match self.kind(index) {
            SymbolKind::Keyword(keyword) => Some(keyword),
            _ => None,
        }
    }

    fn text(&self, index: usize) -> &[u8] {
        self.symbols
            .get(index)
            .map_or(&[], |symbol| symbol.text.as_slice())
    }

    fn first(&self, index: usize) -> Result<u8, FormatError> {
        text::at(self.text(index), 0)
    }

    fn is(&self, index: usize, byte: u8) -> bool {
        self.text(index) == [byte]
    }

    fn set_kind(&mut self, index: usize, kind: SymbolKind) {
        if let Some(symbol) = self.symbols.get_mut(index) {
            symbol.kind = kind;
        }
    }

    fn set_text(&mut self, index: usize, text: &[u8]) {
        if let Some(symbol) = self.symbols.get_mut(index) {
            symbol.text.clear();
            symbol.text.extend_from_slice(text);
        }
    }

    fn merge_into(&mut self, target: usize, index: usize) {
        let text = self
            .symbols
            .get_mut(index)
            .map(|symbol| std::mem::take(&mut symbol.text))
            .unwrap_or_default();
        if let Some(target) = self.symbols.get_mut(target) {
            target.text.extend_from_slice(&text);
        }
        self.drop_symbol(index);
    }

    fn drop_symbol(&mut self, index: usize) {
        self.set_kind(index, SymbolKind::Empty);
        self.result_count = self.result_count.wrapping_sub(1);
    }

    fn advance(&self, position: &mut u16, index: usize) {
        *position = position.wrapping_add(self.text(index).len() as u16);
    }

    fn reset(&mut self) {
        self.symbols.clear();
        self.result_count = 0;
        self.info = ScanInfo::UNSCANNED;
        self.exponent_position = 0;
        self.blank_position = 0;
        self.decimal_position = 0;
        self.has_exponent = false;
        self.has_fraction = false;
        self.has_blank = false;
        self.has_decimal_separator = false;
    }

    fn divide(&mut self, format: &[u8]) -> Result<u16, FormatError> {
        self.has_currency = has_unquoted_currency(format);
        self.reset();
        let mut position = 0;
        let mut has_star = false;
        while position < format.len() && self.symbols.len() < MAX_SYMBOLS {
            let symbol = next_symbol(format, &mut position)?;
            if symbol.kind == SymbolKind::Star {
                if has_star {
                    return Ok(position as u16);
                }
                has_star = true;
            }
            self.symbols.push(symbol);
        }
        Ok(0)
    }

    fn skip_strings(&self, index: &mut usize, position: &mut u16) {
        while matches!(
            self.kind(*index),
            SymbolKind::String | SymbolKind::Blank | SymbolKind::Star
        ) {
            self.advance(position, *index);
            *index += 1;
        }
    }

    fn previous_keyword(&self, index: usize) -> Option<Keyword> {
        if index == 0 || index >= self.count() {
            return None;
        }
        let mut index = index - 1;
        while index > 0 && !self.kind(index).is_keyword() {
            index -= 1;
        }
        self.keyword(index)
    }

    fn next_keyword(&self, index: usize) -> Option<Keyword> {
        if index + 1 >= self.count() {
            return None;
        }
        let mut index = index + 1;
        while index + 1 < self.count() && !self.kind(index).is_keyword() {
            index += 1;
        }
        self.keyword(index)
    }

    fn previous_kind(&self, index: usize) -> SymbolKind {
        if index == 0 || index >= self.count() {
            return SymbolKind::None;
        }
        let mut index = index - 1;
        while index > 0 && self.kind(index) == SymbolKind::Empty {
            index -= 1;
        }
        self.kind(index)
    }

    fn is_skipped_for_neighbours(&self, index: usize) -> bool {
        matches!(
            self.kind(index),
            SymbolKind::Empty | SymbolKind::String | SymbolKind::Star | SymbolKind::Blank
        )
    }

    fn previous_char(&self, index: usize) -> u8 {
        if index == 0 || index >= self.count() {
            return b' ';
        }
        let mut index = index - 1;
        while index > 0 && self.is_skipped_for_neighbours(index) {
            index -= 1;
        }
        self.text(index).last().copied().unwrap_or(b' ')
    }

    fn next_char(&self, index: usize) -> u8 {
        if index + 1 >= self.count() {
            return b' ';
        }
        let mut index = index + 1;
        while index + 1 < self.count() && self.is_skipped_for_neighbours(index) {
            index += 1;
        }
        self.text(index).first().copied().unwrap_or(b' ')
    }

    fn is_100sec_zero(&self, index: usize, had_decimal_separator: bool) -> bool {
        matches!(
            self.previous_keyword(index),
            Some(Keyword::Second | Keyword::Second2)
        ) && (had_decimal_separator || (index > 0 && self.kind(index - 1) == SymbolKind::String))
    }

    fn scan_type(&mut self) -> Result<u16, FormatError> {
        use FormatType::{
            Currency, Date, DateTime, Defined, Fraction, Number, Percent, Scientific, Text, Time,
            Undefined,
        };
        let mut position = 0u16;
        let mut index = 0;
        let mut have_general = false;
        let mut have_digits = false;

        self.skip_strings(&mut index, &mut position);
        while index < self.count() {
            let mut match_bracket = false;
            let new_type = if let Some(keyword) = self.keyword(index) {
                match keyword {
                    Keyword::Exponent => Scientific,
                    Keyword::AmPm
                    | Keyword::AP
                    | Keyword::Hour
                    | Keyword::Hour2
                    | Keyword::Second
                    | Keyword::Second2 => Time,
                    Keyword::Month | Keyword::Month2 => {
                        // Minutes follow hours, precede seconds or open an elapsed time bracket.
                        let is_minute = matches!(
                            self.previous_keyword(index),
                            Some(Keyword::Hour | Keyword::Hour2)
                        ) || matches!(
                            self.next_keyword(index),
                            Some(Keyword::Second | Keyword::Second2)
                        ) || self.previous_char(index) == b'[';
                        if is_minute {
                            let minute = if keyword == Keyword::Month {
                                Keyword::Minute
                            } else {
                                Keyword::Minute2
                            };
                            self.set_kind(index, SymbolKind::Keyword(minute));
                            Time
                        } else {
                            Date
                        }
                    }
                    Keyword::BankSymbol => Currency,
                    Keyword::General => {
                        have_general = true;
                        Number
                    }
                    Keyword::Minute | Keyword::Minute2 => Undefined,
                    _ => Date,
                }
            } else {
                match self.first(index)? {
                    b'#' | b'?' => {
                        have_digits = true;
                        Number
                    }
                    b'0' => {
                        if self.info.format_type.has_time() {
                            if !self.is_100sec_zero(index, self.has_decimal_separator) {
                                return Ok(position);
                            }
                            self.has_decimal_separator = true;
                            Time
                        } else {
                            have_digits = true;
                            Number
                        }
                    }
                    b'%' => Percent,
                    b'/' => Fraction,
                    b'[' => {
                        let bracketed = if index + 1 < self.count()
                            && self.kind(index + 1) == SymbolKind::String
                        {
                            self.text(index + 1).first().copied()
                        } else {
                            None
                        };
                        match bracketed {
                            Some(b'$') => {
                                match_bracket = true;
                                Currency
                            }
                            Some(b'~') => {
                                match_bracket = true;
                                Date
                            }
                            _ => match self.next_keyword(index) {
                                Some(
                                    Keyword::Hour
                                    | Keyword::Hour2
                                    | Keyword::Month
                                    | Keyword::Month2
                                    | Keyword::Second
                                    | Keyword::Second2,
                                ) => Time,
                                _ => return Ok(position),
                            },
                        }
                    }
                    b'@' => Text,
                    _ => {
                        if self.is(index, locale::TIME_100SEC_SEPARATOR) {
                            self.has_decimal_separator = true;
                        }
                        Undefined
                    }
                }
            };

            let scanned = self.info.format_type;
            let is_date_separator = self.is(index, locale::DATE_SEPARATOR);
            let is_time_separator = self.is(index, locale::TIME_SEPARATOR);
            self.info.format_type = match (scanned, new_type) {
                (Undefined, _) => new_type,
                (Text, _) | (_, Text) => Text,
                (_, Undefined) => scanned,
                _ if scanned == new_type => scanned,
                (Date, Time) | (Time, Date) => DateTime,
                (Date | Time | DateTime, Fraction) => scanned,
                (DateTime, Date | Time) => scanned,
                (Date | Time | DateTime, _) if self.has_currency => Undefined,
                (Date, _) if is_date_separator => scanned,
                (Time, _) if is_time_separator => scanned,
                (DateTime, _) if is_date_separator || is_time_separator => scanned,
                (Date | Time | DateTime, _) => return Ok(position),
                (Percent | Scientific | Fraction, Number) => scanned,
                (Percent | Scientific | Fraction, _) => return Ok(position),
                (Number, Scientific | Percent | Fraction | Currency) => new_type,
                (Number, _) if self.has_currency => Undefined,
                (Number, _) => return Ok(position),
                _ => scanned,
            };

            self.advance(&mut position, index);
            index += 1;
            // Nothing inside [$...] or [~...] takes part in the type detection.
            if match_bracket {
                while match_bracket && index < self.count() {
                    if self.kind(index) == SymbolKind::Delimiter && self.first(index)? == b']' {
                        match_bracket = false;
                    } else {
                        self.set_kind(index, SymbolKind::String);
                    }
                    self.advance(&mut position, index);
                    index += 1;
                }
                if match_bracket {
                    return Ok(position);
                }
            }
            self.skip_strings(&mut index, &mut position);
        }

        // A currency symbol makes a currency format only together with digit placeholders.
        if matches!(self.info.format_type, Number | Undefined)
            && self.has_currency
            && !have_general
            && have_digits
        {
            self.info.format_type = Currency;
        }
        if self.info.format_type == Undefined {
            self.info.format_type = Defined;
        }
        Ok(0)
    }

    fn insert_symbol(&mut self, position: &mut usize, kind: SymbolKind, text: &[u8]) -> bool {
        let count = self.count();
        if count >= MAX_SYMBOLS || *position > count {
            return false;
        }
        self.result_count = self.result_count.wrapping_add(1);
        let symbol = Symbol {
            kind,
            text: text.to_vec(),
        };
        if *position > 0 && self.kind(*position - 1) == SymbolKind::Empty {
            *position -= 1;
            self.symbols[*position] = symbol;
        } else {
            if count + 1 >= MAX_SYMBOLS {
                return false;
            }
            self.symbols.insert(*position, symbol);
        }
        true
    }

    fn final_scan(&mut self, format: &mut Text) -> Result<u16, FormatError> {
        let mut position = 0u16;
        self.result_count = self.count() as u16;
        self.has_decimal_separator = false;

        let failure = match self.info.format_type {
            FormatType::Text | FormatType::Defined => self.final_scan_text(&mut position)?,
            FormatType::Number
            | FormatType::Percent
            | FormatType::Currency
            | FormatType::Scientific
            | FormatType::Fraction => self.final_scan_number(&mut position)?,
            FormatType::Date => self.final_scan_date(&mut position),
            FormatType::Time | FormatType::DateTime => self.final_scan_time(&mut position)?,
            FormatType::Undefined => false,
        };
        if failure {
            return Ok(position);
        }
        let info = &self.info;
        let incomplete = match info.format_type {
            FormatType::Scientific => {
                u32::from(info.count_pre) + u32::from(info.count_post) == 0 || info.count_exp == 0
            }
            FormatType::Fraction => info.count_exp > 8 || info.count_exp == 0,
            _ => false,
        };
        if incomplete {
            return Ok(position);
        }
        self.rebuild(format)?;
        Ok(0)
    }

    fn final_scan_text(&mut self, position: &mut u16) -> Result<bool, FormatError> {
        for index in 0..self.count() {
            match self.kind(index) {
                SymbolKind::Blank | SymbolKind::Star | SymbolKind::Keyword(Keyword::General) => {}
                kind => {
                    if kind != SymbolKind::Delimiter || self.first(index)? != b'@' {
                        self.set_kind(index, SymbolKind::String);
                    }
                }
            }
            self.advance(position, index);
        }
        Ok(false)
    }

    fn final_scan_number(&mut self, position: &mut u16) -> Result<bool, FormatError> {
        let format_type = self.info.format_type;
        let is_fraction = format_type == FormatType::Fraction;
        let mut index = 0;
        let mut counter = 0u16;
        while index < self.count() {
            match self.kind(index) {
                SymbolKind::Blank | SymbolKind::Star | SymbolKind::Keyword(Keyword::BankSymbol) => {
                    self.advance(position, index);
                    index += 1;
                }
                SymbolKind::Keyword(Keyword::General) => {
                    self.info.thousand_count = GENERAL_IN_FORMAT;
                    self.advance(position, index);
                    index += 1;
                }
                kind @ (SymbolKind::String | SymbolKind::Keyword(_)) => {
                    if format_type == FormatType::Scientific
                        && kind == SymbolKind::Keyword(Keyword::Exponent)
                    {
                        if self.has_exponent {
                            return Ok(true);
                        }
                        self.has_exponent = true;
                        self.exponent_position = index;
                        if self.has_decimal_separator {
                            self.info.count_post = counter;
                        } else {
                            self.info.count_pre = counter;
                        }
                        counter = 0;
                        self.set_kind(index, SymbolKind::Exponent);
                    } else if is_fraction && self.first(index)? == b' ' {
                        if !self.has_blank && !self.has_fraction {
                            if self.has_decimal_separator && counter > 0 {
                                return Ok(true);
                            }
                            self.has_blank = true;
                            self.blank_position = index;
                            self.info.count_pre = counter;
                            counter = 0;
                        }
                        self.set_kind(index, SymbolKind::FractionBlank);
                    } else {
                        self.set_kind(index, SymbolKind::String);
                    }
                    self.advance(position, index);
                    index += 1;
                }
                SymbolKind::Delimiter => {
                    if self.final_scan_delimiter(position, &mut index, &mut counter)? {
                        return Ok(true);
                    }
                }
                SymbolKind::Digit if is_fraction => {
                    // Literal digits are a fixed denominator.
                    let mut digits = self.text(index).to_vec();
                    self.advance(position, index);
                    index += 1;
                    counter = counter.wrapping_add(1);
                    while index < self.count()
                        && (self.kind(index) == SymbolKind::Digit || self.first(index)? == b'0')
                    {
                        self.set_kind(index, SymbolKind::Digit);
                        digits.extend_from_slice(self.text(index));
                        self.advance(position, index);
                        counter = counter.wrapping_add(1);
                        index += 1;
                    }
                    self.info.exp_val = digits
                        .iter()
                        .try_fold(0i32, |value, &digit| {
                            value
                                .checked_mul(10)?
                                .checked_add(i32::from(digit.wrapping_sub(b'0')))
                        })
                        .ok_or(FormatError)? as u32;
                }
                _ => {
                    self.advance(position, index);
                    index += 1;
                }
            }
        }

        if is_fraction {
            if self.has_fraction {
                self.info.count_exp = counter;
            } else if self.has_blank {
                self.info.count_post = counter;
            } else {
                self.info.count_pre = counter;
            }
        } else if self.has_exponent {
            self.info.count_exp = counter;
        } else if self.has_decimal_separator {
            self.info.count_post = counter;
        } else {
            self.info.count_pre = counter;
        }

        if self.info.thousand && !self.insert_group_separators(index) {
            return Ok(true);
        }

        // Adjacent digit placeholders become one symbol.
        let mut index = 0;
        while index < self.count() {
            if self.kind(index) == SymbolKind::Digit {
                let target = index;
                index += 1;
                while index < self.count() && self.kind(index) == SymbolKind::Digit {
                    self.merge_into(target, index);
                    index += 1;
                }
            }
            index += 1;
        }
        Ok(false)
    }

    fn insert_group_separators(&mut self, end: usize) -> bool {
        let end = if self.has_fraction {
            if self.has_blank {
                self.blank_position
            } else {
                0
            }
        } else if self.has_decimal_separator {
            self.decimal_position
        } else if self.has_exponent {
            self.exponent_position
        } else {
            end
        };
        let mut digit_count = 0u16;
        let mut grouping = DigitGrouping::new();
        let mut first_digit = end;
        let mut first_separator = end;
        let mut index = end;
        while index > 0 {
            index -= 1;
            if self.kind(index) != SymbolKind::Digit {
                continue;
            }
            first_digit = index;
            digit_count = digit_count.wrapping_add(self.text(index).len() as u16);
            if index > 0 && i32::from(digit_count) >= grouping.position() {
                if !self.insert_symbol(
                    &mut index,
                    SymbolKind::ThousandSeparator,
                    &[locale::THOUSAND_SEPARATOR],
                ) {
                    return false;
                }
                first_digit = index + 1;
                first_separator = index;
                grouping.advance();
            }
        }
        // A separator in front of the first digit, as in "text",000, goes away again.
        if first_separator < first_digit {
            self.drop_symbol(first_separator);
        }
        true
    }

    fn final_scan_delimiter(
        &mut self,
        position: &mut u16,
        index: &mut usize,
        counter: &mut u16,
    ) -> Result<bool, FormatError> {
        let format_type = self.info.format_type;
        let is_fraction = format_type == FormatType::Fraction;
        let here = self.first(*index)?;
        match here {
            b'#' | b'0' | b'?' => {
                if self.info.thousand_count > 0 {
                    return Ok(true);
                }
                loop {
                    self.set_kind(*index, SymbolKind::Digit);
                    self.advance(position, *index);
                    *index += 1;
                    *counter = counter.wrapping_add(1);
                    if *index >= self.count() || !is_digit_placeholder(self.first(*index)?) {
                        break;
                    }
                }
            }
            b'-' => {
                // Dashes right after the decimal separator are digit placeholders, as in "0.--".
                let after_decimal_separator = self.has_decimal_separator
                    && self.decimal_position + 1 == *index
                    && self.kind(self.decimal_position) == SymbolKind::DecimalSeparator;
                if after_decimal_separator {
                    let target = *index;
                    self.set_kind(target, SymbolKind::Digit);
                    self.advance(position, target);
                    *index += 1;
                    *counter = counter.wrapping_add(1);
                    while *index < self.count() && self.first(*index)? == b'-' {
                        // In currency formats the last of more than two dashes is a minus sign.
                        if format_type == FormatType::Currency
                            && self.text(target).len() >= 2
                            && (*index == self.count() - 1 || self.first(*index + 1)? != b'-')
                        {
                            break;
                        }
                        self.advance(position, *index);
                        self.merge_into(target, *index);
                        *counter = counter.wrapping_add(1);
                        *index += 1;
                    }
                } else {
                    self.set_kind(*index, SymbolKind::String);
                    self.advance(position, *index);
                    *index += 1;
                }
            }
            locale::THOUSAND_SEPARATOR => {
                let previous = self.previous_char(*index);
                let after_digit = *index > 0 && matches!(previous, b'#' | b'0');
                if self.has_exponent || self.has_blank || self.has_fraction {
                    self.advance(position, *index);
                    self.drop_symbol(*index);
                    *index += 1;
                } else if after_digit
                    && *index + 1 < self.count()
                    && matches!(self.next_char(*index), b'#' | b'0')
                {
                    // Between digits: grouping is wanted, the separators are placed afterwards.
                    self.advance(position, *index);
                    self.info.thousand = true;
                    self.drop_symbol(*index);
                    *index += 1;
                } else if after_digit
                    && self.previous_kind(*index) == SymbolKind::Digit
                    && self.info.thousand_count < GENERAL_IN_FORMAT
                {
                    // Trailing separators divide the value by thousands.
                    loop {
                        self.info.thousand_count += 1;
                        self.set_kind(*index, SymbolKind::ThousandSeparator);
                        self.advance(position, *index);
                        *index += 1;
                        if !self.is(*index, locale::THOUSAND_SEPARATOR) {
                            break;
                        }
                    }
                } else {
                    self.final_scan_literal_run(position, index, here);
                }
            }
            locale::DECIMAL_SEPARATOR => {
                if self.has_blank || self.has_fraction {
                    return Ok(true);
                } else if self.has_exponent {
                    self.advance(position, *index);
                    self.drop_symbol(*index);
                    *index += 1;
                } else if self.has_decimal_separator {
                    self.final_scan_literal_run(position, index, here);
                } else {
                    self.advance(position, *index);
                    self.set_kind(*index, SymbolKind::DecimalSeparator);
                    self.has_decimal_separator = true;
                    self.decimal_position = *index;
                    self.info.count_pre = *counter;
                    *counter = 0;
                    *index += 1;
                }
            }
            b'\'' | b' ' => {
                if here == b' ' && is_fraction {
                    // The blank between integer and fraction; the caller sees it again as a string.
                    if !self.has_blank && !self.has_fraction {
                        if self.has_decimal_separator && *counter > 0 {
                            return Ok(true);
                        }
                        self.has_blank = true;
                        self.blank_position = *index;
                        self.info.count_pre = *counter;
                        *counter = 0;
                    }
                    self.set_kind(*index, SymbolKind::String);
                    self.advance(position, *index);
                } else {
                    self.final_scan_literal_run(position, index, here);
                }
            }
            b'/' => {
                if is_fraction {
                    if *index == 0
                        || !matches!(self.kind(*index - 1), SymbolKind::Digit | SymbolKind::Empty)
                    {
                        *position = (*position).max(1);
                        return Ok(true);
                    } else if !self.has_fraction || (self.has_decimal_separator && *counter > 0) {
                        self.has_fraction = true;
                        self.info.count_post = *counter;
                        *counter = 0;
                        self.set_kind(*index, SymbolKind::Fraction);
                        self.advance(position, *index);
                        *index += 1;
                    } else {
                        return Ok(true);
                    }
                } else {
                    self.set_kind(*index, SymbolKind::String);
                    self.advance(position, *index);
                    *index += 1;
                }
            }
            b'[' => {
                let opens_currency = format_type == FormatType::Currency
                    && *index + 1 < self.count()
                    && self.kind(*index + 1) == SymbolKind::String
                    && self.first(*index + 1)? == b'$';
                if opens_currency {
                    if self.final_scan_currency(position, index)? {
                        return Ok(true);
                    }
                } else {
                    self.set_kind(*index, SymbolKind::String);
                    self.advance(position, *index);
                    *index += 1;
                }
            }
            _ => {
                let kind = if format_type == FormatType::Percent && here == b'%' {
                    SymbolKind::Percent
                } else {
                    SymbolKind::String
                };
                self.set_kind(*index, kind);
                self.advance(position, *index);
                *index += 1;
            }
        }
        Ok(false)
    }

    fn final_scan_literal_run(&mut self, position: &mut u16, index: &mut usize, byte: u8) {
        let target = *index;
        self.set_kind(target, SymbolKind::String);
        self.advance(position, target);
        *index += 1;
        while self.is(*index, byte) {
            self.advance(position, *index);
            self.merge_into(target, *index);
            *index += 1;
        }
    }

    fn final_scan_currency(
        &mut self,
        position: &mut u16,
        index: &mut usize,
    ) -> Result<bool, FormatError> {
        self.advance(position, *index);
        self.set_kind(*index, SymbolKind::CurrencyDelimiter);
        *index += 1;
        self.advance(position, *index);
        self.merge_into(*index - 1, *index);
        *index += 1;
        if *index >= self.count() {
            return Ok(true);
        }
        self.advance(position, *index);
        let currency = *index;
        let mut target = currency;
        self.set_kind(currency, SymbolKind::Currency);
        let mut had_dash = false;
        *index += 1;
        while *index < self.count() && self.first(*index)? != b']' {
            self.advance(position, *index);
            if !had_dash && self.first(*index)? == b'-' {
                had_dash = true;
                target = *index;
                self.set_kind(*index, SymbolKind::CurrencyExtension);
            } else {
                self.merge_into(target, *index);
            }
            *index += 1;
        }
        if !self.text(currency).is_empty() && *index < self.count() && self.first(*index)? == b']' {
            self.set_kind(*index, SymbolKind::CurrencyDelimiter);
            self.advance(position, *index);
            *index += 1;
            Ok(false)
        } else {
            Ok(true)
        }
    }

    fn final_scan_date(&mut self, position: &mut u16) -> bool {
        for index in 0..self.count() {
            match self.kind(index) {
                SymbolKind::Blank | SymbolKind::Star | SymbolKind::String => {}
                SymbolKind::Delimiter => {
                    let kind = if self.is(index, locale::DATE_SEPARATOR) {
                        SymbolKind::DateSeparator
                    } else {
                        SymbolKind::String
                    };
                    self.set_kind(index, kind);
                }
                SymbolKind::Keyword(keyword) if keyword.is_date() => {
                    self.set_text(index, keyword.text());
                }
                _ => self.set_kind(index, SymbolKind::String),
            }
            self.advance(position, index);
        }
        false
    }

    fn final_scan_time(&mut self, position: &mut u16) -> Result<bool, FormatError> {
        let is_time = self.info.format_type == FormatType::Time;
        // In a date-time format the last keyword decides how separators and zeros are read.
        let mut time_part = is_time;
        let mut has_am_pm = false;
        let mut index = 0;
        let mut counter = 0u16;
        while index < self.count() {
            match self.kind(index) {
                SymbolKind::Blank | SymbolKind::Star | SymbolKind::String => {
                    self.advance(position, index);
                    index += 1;
                }
                SymbolKind::Delimiter => match self.first(index)? {
                    b'0' => {
                        if !time_part || !self.is_100sec_zero(index, self.has_decimal_separator) {
                            return Ok(true);
                        }
                        self.has_decimal_separator = true;
                        let target = index;
                        self.set_kind(target, SymbolKind::Digit);
                        index += 1;
                        self.advance(position, index);
                        counter = counter.wrapping_add(1);
                        while index < self.count() && self.first(index)? == b'0' {
                            self.advance(position, index);
                            self.merge_into(target, index);
                            counter = counter.wrapping_add(1);
                            index += 1;
                        }
                    }
                    b'#' | b'?' => return Ok(true),
                    b'[' if is_time => {
                        if self.info.thousand {
                            return Ok(true);
                        }
                        self.info.thousand = true;
                        self.info.thousand_count = match self.next_char(index).to_ascii_uppercase()
                        {
                            b'H' => 1,
                            b'M' => 2,
                            b'S' => 3,
                            _ => return Ok(true),
                        };
                        self.advance(position, index);
                        index += 1;
                    }
                    b']' if is_time => {
                        if !self.info.thousand {
                            return Ok(true);
                        }
                        self.advance(position, index);
                        index += 1;
                    }
                    _ => {
                        self.advance(position, index);
                        let kind = if time_part && self.is(index, locale::TIME_SEPARATOR) {
                            SymbolKind::TimeSeparator
                        } else if time_part && self.is(index, locale::TIME_100SEC_SEPARATOR) {
                            self.has_decimal_separator = true;
                            SymbolKind::Time100SecSeparator
                        } else if !time_part && self.is(index, locale::DATE_SEPARATOR) {
                            SymbolKind::DateSeparator
                        } else {
                            SymbolKind::String
                        };
                        self.set_kind(index, kind);
                        index += 1;
                    }
                },
                SymbolKind::Keyword(keyword) => {
                    let keeps_kind = match keyword {
                        Keyword::AmPm | Keyword::AP => {
                            has_am_pm = true;
                            time_part = true;
                            true
                        }
                        Keyword::Minute
                        | Keyword::Minute2
                        | Keyword::Hour
                        | Keyword::Hour2
                        | Keyword::Second
                        | Keyword::Second2 => {
                            time_part = true;
                            true
                        }
                        _ if !is_time && keyword.is_date() => {
                            time_part = false;
                            true
                        }
                        _ => false,
                    };
                    if keeps_kind {
                        self.set_text(index, keyword.text());
                    } else {
                        self.set_kind(index, SymbolKind::String);
                    }
                    self.advance(position, index);
                    index += 1;
                }
                _ => {
                    self.set_kind(index, SymbolKind::String);
                    self.advance(position, index);
                    index += 1;
                }
            }
        }
        self.info.count_post = counter;
        if has_am_pm {
            self.info.count_exp = 1;
        }
        Ok(false)
    }

    fn rebuild(&mut self, format: &mut Text) -> Result<(), FormatError> {
        let format_type = self.info.format_type;
        let is_currency = format_type == FormatType::Currency;
        format.clear();
        let mut index = 0;
        while index < self.count() {
            match self.kind(index) {
                SymbolKind::String => {
                    // Start of the part of the literal, in the format and in the merged symbol,
                    // that has not been searched for the currency symbol yet.
                    let mut format_start = format.len();
                    let mut symbol_start = 0;
                    let target = index;
                    loop {
                        match *self.text(index) {
                            // Simple escapes lose their backslash where it has no function.
                            [b'\\', escaped @ (b'+' | b'-')] => format.push(escaped),
                            [b'\\', escaped @ (b' ' | b'.' | b'/')] => {
                                let keeps_backslash = match escaped {
                                    b'.' => !format_type.has_date() || format_type.has_time(),
                                    b'/' => format_type.has_date(),
                                    _ => false,
                                } || format_type == FormatType::Fraction;
                                if keeps_backslash {
                                    format.extend_from_slice(self.text(index));
                                } else {
                                    format.push(escaped);
                                }
                            }
                            _ => format.extend_from_slice(self.text(index)),
                        }
                        let unquoted = self
                            .symbols
                            .get_mut(index)
                            .is_some_and(|symbol| remove_quotes(&mut symbol.text));
                        if unquoted && is_currency {
                            self.mark_currency(format, format_start, target, symbol_start)?;
                            format_start = format.len();
                            symbol_start = self.text(target).len();
                            if target != index {
                                symbol_start += self.text(index).len();
                            }
                        }
                        if target != index {
                            self.merge_into(target, index);
                        }
                        index += 1;
                        if index >= self.count() || self.kind(index) != SymbolKind::String {
                            break;
                        }
                    }
                    if is_currency && format_start < format.len() {
                        self.mark_currency(format, format_start, target, symbol_start)?;
                    }
                    continue;
                }
                SymbolKind::Currency => {
                    format.extend_from_slice(self.text(index));
                    if let Some(symbol) = self.symbols.get_mut(index) {
                        remove_quotes(&mut symbol.text);
                    }
                }
                SymbolKind::Empty => {}
                _ => format.extend_from_slice(self.text(index)),
            }
            index += 1;
        }
        Ok(())
    }

    fn mark_currency(
        &self,
        format: &mut Text,
        format_start: usize,
        symbol: usize,
        symbol_start: usize,
    ) -> Result<(), FormatError> {
        let literal = text::tail(self.text(symbol), symbol_start)?;
        if let Some(offset) = literal
            .iter()
            .position(|&byte| byte == locale::CURRENCY_SYMBOL)
        {
            let position = format_start + offset;
            match position.cmp(&format.len()) {
                std::cmp::Ordering::Less => format[position] = locale::CURRENCY_SYMBOL,
                std::cmp::Ordering::Equal => format.push(locale::CURRENCY_SYMBOL),
                std::cmp::Ordering::Greater => return Err(FormatError),
            }
        }
        Ok(())
    }
}
