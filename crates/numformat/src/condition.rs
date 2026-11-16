//! Parser for the number of a format condition such as `[>=100]`.
//!
//! Conditions go through a general purpose input recogniser that also knows dates, times,
//! currencies, percentages and fractions. Only plain and scientific numbers are accepted
//! as limits, but the other kinds still have to be recognised: they decide between "not a number"
//! (the limit becomes 0) and an input the recogniser fails on ([`FormatError`]), which are date
//! and time parts that do not fit a 32-bit integer.
//!
//! The input is split into alternating runs of digits and other characters, at most
//! [`MAX_STRINGS`] of them. The strings before, between and after the numbers are then matched
//! against what may stand there: signs, separators, month and day names, currency and so on.

use crate::FormatError;
use crate::locale;
use crate::math::pow;
use crate::text::Text;

const MAX_STRINGS: usize = 20;
const MAX_INPUT_LENGTH: usize = 308;

#[derive(Clone, Copy, PartialEq, Eq)]
enum InputType {
    Undefined,
    Date,
    Time,
    DateTime,
    Currency,
    Number,
    Scientific,
    Fraction,
    Percent,
}

/// Where the decimal separator was found.
#[derive(Clone, Copy, PartialEq, Eq)]
enum DecimalPosition {
    None,
    Start,
    Middle,
    End,
}

/// Position of a month name relative to the numbers.
#[derive(Clone, Copy, PartialEq, Eq)]
enum MonthPosition {
    None,
    Start,
    Middle,
    End,
}

pub(crate) fn parse_limit(text: &[u8]) -> Result<Option<f64>, FormatError> {
    if text.is_empty() || text.len() > MAX_INPUT_LENGTH {
        return Ok(None);
    }
    let mut scan = InputScan::new();
    if !scan.recognise(&text.to_ascii_uppercase())? {
        return Ok(None);
    }
    scan.number()
}

/// The state of one recognition.
///
/// `numbers` holds the indices into `strings` of the digit runs, `grouped_number` the index of a
/// digit run that absorbed following `,ddd` groups (0 if none). `month` is negative for an
/// abbreviated name, `time_position` the one-based index of the first string with a time
/// separator, and `open_parenthesis` is set while a `(` waits for its `)`.
struct InputScan {
    strings: Vec<Text>,
    is_number: [bool; MAX_STRINGS],
    numbers: Vec<usize>,
    grouped_number: usize,
    sign: i32,
    month: i32,
    month_position: MonthPosition,
    time_position: usize,
    decimal_position: DecimalPosition,
    open_parenthesis: bool,
    exponent_sign: i32,
    thousand_count: usize,
    scanned_type: InputType,
    may_be_iso8601: Option<bool>,
}

fn is_digit(byte: u8) -> bool {
    byte.is_ascii_digit()
}

fn contains_at(text: &[u8], position: usize, what: &[u8]) -> bool {
    !what.is_empty()
        && text
            .get(position..)
            .is_some_and(|rest| rest.starts_with(what))
}

fn skip_char(byte: u8, text: &[u8], position: &mut usize) -> bool {
    if text.get(*position) == Some(&byte) {
        *position += 1;
        true
    } else {
        false
    }
}

fn skip_blanks(text: &[u8], position: &mut usize) {
    while text.get(*position) == Some(&b' ') {
        *position += 1;
    }
}

fn skip_string(what: &[u8], text: &[u8], position: &mut usize) -> bool {
    if contains_at(text, *position, what) {
        *position += what.len();
        true
    } else {
        false
    }
}

fn skip_decimal_separator(text: &[u8], position: &mut usize) -> bool {
    if *position < text.len() && text[*position..] == [locale::DECIMAL_SEPARATOR] {
        *position += 1;
        true
    } else {
        false
    }
}

fn skip_exponent_sign(text: &[u8], position: &mut usize) -> i32 {
    match text.get(*position) {
        Some(b'+') => {
            *position += 1;
            1
        }
        Some(b'-') => {
            *position += 1;
            -1
        }
        _ => 0,
    }
}

fn skip_date_separator(text: &[u8], position: &mut usize) -> bool {
    skip_char(locale::DATE_SEPARATOR, text, position)
        || skip_char(b'.', text, position)
        || skip_char(b'-', text, position)
}

fn skip_name(names: &[locale::Name], text: &[u8], position: &mut usize) -> i32 {
    let rest = text.get(*position..).unwrap_or_default();
    for (index, name) in names.iter().enumerate() {
        let number = index as i32 + 1;
        for (candidate, result) in [(name.full, number), (name.abbreviation, -number)] {
            let matches = rest
                .get(..candidate.len())
                .is_some_and(|start| start.eq_ignore_ascii_case(candidate.as_bytes()));
            if matches {
                *position += candidate.len();
                return result;
            }
        }
    }
    0
}

fn skip_month(text: &[u8], position: &mut usize) -> i32 {
    skip_name(&locale::MONTHS, text, position)
}

fn skip_currency(text: &[u8], position: &mut usize) -> bool {
    skip_char(locale::CURRENCY_SYMBOL, text, position)
}

fn skip_am_pm(text: &[u8], position: &mut usize) -> bool {
    skip_string(locale::TIME_AM, text, position) || skip_string(locale::TIME_PM, text, position)
}

fn to_int(digits: &[u8]) -> Result<i32, FormatError> {
    if digits.is_empty() {
        return Err(FormatError);
    }
    digits.iter().try_fold(0i32, |value, &digit| {
        value
            .checked_mul(10)
            .and_then(|value| value.checked_add(i32::from(digit - b'0')))
            .ok_or(FormatError)
    })
}

fn string_to_double(text: &[u8]) -> f64 {
    let mut integer = 0.0;
    let mut fraction = 0.0;
    let mut exponent = 0;
    let mut before_separator = true;
    // Digit by digit: the rounding errors this accumulates are part of the limit.
    for &byte in text {
        let digit = f64::from(i32::from(byte) - i32::from(b'0'));
        if byte == b'.' {
            before_separator = false;
        } else if before_separator {
            integer = integer * 10.0 + digit;
        } else {
            fraction = fraction * 10.0 + digit;
            exponent -= 1;
        }
    }
    if fraction != 0.0 {
        integer + fraction * pow(10.0, f64::from(exponent))
    } else {
        integer
    }
}

fn parse_double(text: &[u8]) -> Option<f64> {
    std::str::from_utf8(text).ok()?.parse().ok()
}

fn scientific_to_double(text: &[u8]) -> Option<f64> {
    let value = parse_double(text)?;
    let exponent_start = text.iter().position(|&byte| byte == b'E')?;
    let (mantissa, exponent) = text.split_at(exponent_start);
    if mantissa.iter().all(|&byte| byte == b'0' || byte == b'.') {
        return Some(0.0);
    }
    // A number that overflows or underflows is not a limit.
    if !value.is_finite() || value < f64::MIN_POSITIVE {
        return None;
    }
    // Underflow is judged before the exponent range limits the rounding, so the smallest normal
    // number can stand for one that underflows; scaled by a power of two the rounding shows.
    if value == f64::MIN_POSITIVE {
        const SCALE: u64 = 1 << 63;
        let mut scaled = multiply_decimal(mantissa, SCALE);
        scaled.extend_from_slice(exponent);
        if parse_double(&scaled)? < f64::MIN_POSITIVE * SCALE as f64 {
            return None;
        }
    }
    Some(value)
}

fn multiply_decimal(digits: &[u8], factor: u64) -> Text {
    let mut product = Text::with_capacity(digits.len() + 20);
    let mut carry = 0u128;
    for &byte in digits.iter().rev() {
        if byte == b'.' {
            product.push(byte);
            continue;
        }
        let value = u128::from(byte.wrapping_sub(b'0')) * u128::from(factor) + carry;
        product.push(b'0' + (value % 10) as u8);
        carry = value / 10;
    }
    while carry > 0 {
        product.push(b'0' + (carry % 10) as u8);
        carry /= 10;
    }
    product.reverse();
    product
}

impl InputScan {
    fn new() -> InputScan {
        InputScan {
            strings: Vec::new(),
            is_number: [false; MAX_STRINGS],
            numbers: Vec::new(),
            grouped_number: 0,
            sign: 0,
            month: 0,
            month_position: MonthPosition::None,
            time_position: 0,
            decimal_position: DecimalPosition::None,
            open_parenthesis: false,
            exponent_sign: 0,
            thousand_count: 0,
            scanned_type: InputType::Undefined,
            may_be_iso8601: None,
        }
    }

    fn string(&self, index: usize) -> &[u8] {
        self.strings.get(index).map_or(&[], Vec::as_slice)
    }

    fn number_string(&self, index: usize) -> &[u8] {
        self.numbers
            .get(index)
            .map_or(&[], |&string| self.string(string))
    }

    fn divide(&mut self, text: &[u8]) {
        let mut position = 0;
        while position < text.len() && self.strings.len() < MAX_STRINGS {
            let start = position;
            let number = text.get(position).is_some_and(|&byte| is_digit(byte));
            // A NUL ends a run without being consumed, so the table fills up with empty strings.
            while text
                .get(position)
                .is_some_and(|&byte| byte != 0 && is_digit(byte) == number)
            {
                position += 1;
            }
            let index = self.strings.len();
            let mut symbol = text[start..position].to_vec();
            if number {
                self.is_number[index] = true;
                self.numbers.push(index);
                if index >= MAX_STRINGS - 7
                    && self.grouped_number == 0
                    && skip_thousands(text, &mut position, &mut symbol)
                {
                    self.grouped_number = index;
                }
            }
            self.strings.push(symbol);
        }
    }

    fn sign_at(&mut self, text: &[u8], position: &mut usize) -> i32 {
        match text.get(*position) {
            Some(b'+') => {
                *position += 1;
                1
            }
            Some(b'(') => {
                self.open_parenthesis = true;
                *position += 1;
                -1
            }
            Some(b'-') => {
                *position += 1;
                -1
            }
            _ => 0,
        }
    }

    fn skip_thousand_separator(
        &self,
        text: &[u8],
        position: &mut usize,
        string_index: usize,
    ) -> bool {
        if text != [locale::THOUSAND_SEPARATOR]
            || string_index + 1 >= self.strings.len()
            || !self.is_number[string_index + 1]
        {
            return false;
        }
        if self.string(string_index + 1).len() == 3 || self.grouped_number == string_index + 1 {
            *position += 1;
            true
        } else {
            false
        }
    }

    fn may_be_iso8601(&mut self) -> Result<bool, FormatError> {
        if let Some(known) = self.may_be_iso8601 {
            return Ok(known);
        }
        let result = self.numbers.len() >= 3 && to_int(self.number_string(0))? > 31;
        self.may_be_iso8601 = Some(result);
        Ok(result)
    }

    fn scan_start_string(&mut self, index: usize) -> bool {
        let text = self.string(index).to_vec();
        let text = text.as_slice();
        let mut position = 0;
        skip_blanks(text, &mut position);
        self.sign = self.sign_at(text, &mut position);
        if self.sign != 0 {
            skip_blanks(text, &mut position);
        }
        if skip_decimal_separator(text, &mut position) {
            self.decimal_position = DecimalPosition::Start;
            skip_blanks(text, &mut position);
        } else if skip_currency(text, &mut position) {
            self.scanned_type = InputType::Currency;
            skip_blanks(text, &mut position);
            if self.sign == 0 {
                self.sign = self.sign_at(text, &mut position);
                if self.sign != 0 {
                    skip_blanks(text, &mut position);
                }
            }
        } else {
            self.month = skip_month(text, &mut position);
            if self.month != 0 {
                self.scanned_type = InputType::Date;
                self.month_position = MonthPosition::Start;
                if self.month < 0 {
                    skip_char(b'.', text, &mut position);
                }
                skip_blanks(text, &mut position);
            } else {
                let day_of_week = skip_name(&locale::DAYS_FROM_SUNDAY, text, &mut position);
                if day_of_week != 0 {
                    self.scanned_type = InputType::Date;
                    if position < text.len() {
                        if day_of_week < 0 {
                            skip_char(b'.', text, &mut position);
                        } else {
                            skip_blanks(text, &mut position);
                            skip_string(
                                locale::LONG_DATE_DAY_OF_WEEK_SEPARATOR,
                                text,
                                &mut position,
                            );
                        }
                        skip_blanks(text, &mut position);
                        self.month = skip_month(text, &mut position);
                        if self.month != 0 {
                            self.month_position = MonthPosition::Start;
                            if self.month < 0 {
                                skip_char(b'.', text, &mut position);
                            }
                            skip_blanks(text, &mut position);
                        }
                    }
                }
            }
        }
        position >= text.len()
    }

    fn scan_mid_string(&mut self, index: usize) -> Result<bool, FormatError> {
        use InputType::{Currency, Date, DateTime, Fraction, Scientific, Time, Undefined};
        let text = self.string(index).to_vec();
        let text = text.as_slice();
        let mut position = 0;
        let number_count = self.numbers.len();

        skip_blanks(text, &mut position);
        if skip_decimal_separator(text, &mut position) {
            match self.decimal_position {
                DecimalPosition::Start | DecimalPosition::End => return Ok(false),
                DecimalPosition::Middle => {
                    // The decimal separator doubles as a date separator.
                    if !matches!(self.scanned_type, Undefined | Date | DateTime) {
                        return Ok(false);
                    }
                    if self.scanned_type == Undefined {
                        self.scanned_type = Date;
                    }
                    skip_blanks(text, &mut position);
                }
                DecimalPosition::None => {
                    self.decimal_position = DecimalPosition::Middle;
                    skip_blanks(text, &mut position);
                }
            }
        }

        if skip_char(b'/', text, &mut position) {
            if !matches!(self.scanned_type, Undefined | Date) {
                return Ok(false);
            } else if self.scanned_type != Date && number_count == 3 && index > 2 {
                skip_blanks(text, &mut position);
                self.scanned_type = Fraction;
            } else {
                position -= 1;
            }
        }

        if self.skip_thousand_separator(text, &mut position, index) {
            if !matches!(self.scanned_type, Undefined | Currency) {
                return Ok(false);
            }
            self.thousand_count += 1;
        }

        skip_blanks(text, &mut position);
        if skip_date_separator(text, &mut position) {
            if !matches!(self.scanned_type, Undefined | Date) {
                return Ok(false);
            }
            skip_blanks(text, &mut position);
            self.scanned_type = Date;
            let month = skip_month(text, &mut position);
            if self.month != 0 && month != 0 {
                return Ok(false);
            }
            if month != 0 {
                self.month = month;
                self.month_position = MonthPosition::Middle;
                if !(self.month < 0 && skip_char(b'.', text, &mut position))
                    && !skip_char(b'-', text, &mut position)
                {
                    skip_string(locale::LONG_DATE_MONTH_SEPARATOR, text, &mut position);
                }
                skip_blanks(text, &mut position);
            }
        }

        let month = skip_month(text, &mut position);
        if month != 0 {
            if self.month != 0 || !matches!(self.scanned_type, Undefined | Date) {
                return Ok(false);
            }
            self.scanned_type = Date;
            self.month = month;
            self.month_position = MonthPosition::Middle;
            if self.month < 0 {
                skip_char(b'.', text, &mut position);
            }
            skip_string(locale::LONG_DATE_MONTH_SEPARATOR, text, &mut position);
            skip_blanks(text, &mut position);
        }

        // The input is upper case already.
        if skip_char(b'E', text, &mut position) {
            if self.scanned_type != Undefined {
                return Ok(false);
            }
            skip_blanks(text, &mut position);
            self.scanned_type = Scientific;
            if self.thousand_count + 2 == number_count
                && self.decimal_position == DecimalPosition::Middle
            {
                self.decimal_position = DecimalPosition::End;
            }
            self.exponent_sign = skip_exponent_sign(text, &mut position);
            skip_blanks(text, &mut position);
        }

        if skip_char(locale::TIME_SEPARATOR, text, &mut position) {
            if self.decimal_position != DecimalPosition::None {
                if !matches!(self.scanned_type, Date | DateTime) {
                    return Ok(false);
                }
                if self.scanned_type == Date {
                    self.decimal_position = DecimalPosition::None;
                }
            }
            if matches!(self.scanned_type, Date | DateTime) && number_count > 3 {
                skip_blanks(text, &mut position);
                self.scanned_type = DateTime;
            } else if !matches!(self.scanned_type, Undefined | Time) {
                return Ok(false);
            } else {
                skip_blanks(text, &mut position);
                self.scanned_type = Time;
            }
            if self.time_position == 0 {
                self.time_position = index + 1;
            }
        }

        if position < text.len() && self.scanned_type == Date {
            if self.month_position == MonthPosition::Start {
                if skip_string(locale::LONG_DATE_DAY_SEPARATOR, text, &mut position) {
                    skip_blanks(text, &mut position);
                }
            } else if index == 5 && position == 0 && text == b"T" && self.may_be_iso8601()? {
                position += 1;
            }
        }

        Ok(position >= text.len())
    }

    fn scan_end_string(&mut self, index: usize) -> bool {
        use InputType::{Currency, Date, DateTime, Percent, Time, Undefined};
        let text = self.string(index).to_vec();
        let text = text.as_slice();
        let mut position = 0;
        let number_count = self.numbers.len();

        skip_blanks(text, &mut position);
        if skip_decimal_separator(text, &mut position) {
            match self.decimal_position {
                DecimalPosition::Start | DecimalPosition::End => return false,
                DecimalPosition::Middle => {
                    if !matches!(self.scanned_type, Undefined | Date | DateTime) {
                        return false;
                    }
                    if self.scanned_type == Undefined {
                        self.scanned_type = Date;
                    }
                    skip_blanks(text, &mut position);
                }
                DecimalPosition::None => {
                    self.decimal_position = DecimalPosition::End;
                    skip_blanks(text, &mut position);
                }
            }
        }

        if self.sign == 0 && self.scanned_type != Date {
            self.sign = self.sign_at(text, &mut position);
            if self.open_parenthesis {
                return false;
            }
        }

        skip_blanks(text, &mut position);
        if self.open_parenthesis && skip_char(b')', text, &mut position) {
            self.open_parenthesis = false;
            skip_blanks(text, &mut position);
        }

        if skip_currency(text, &mut position) {
            if self.scanned_type != Undefined {
                return false;
            }
            skip_blanks(text, &mut position);
            self.scanned_type = Currency;
            if self.sign == 0 {
                self.sign = self.sign_at(text, &mut position);
                skip_blanks(text, &mut position);
                if self.open_parenthesis {
                    return false;
                }
            }
            if self.open_parenthesis && skip_char(b')', text, &mut position) {
                self.open_parenthesis = false;
                skip_blanks(text, &mut position);
            }
        }

        if skip_char(b'%', text, &mut position) {
            if self.scanned_type != Undefined {
                return false;
            }
            skip_blanks(text, &mut position);
            self.scanned_type = Percent;
        }

        if skip_char(locale::TIME_SEPARATOR, text, &mut position) {
            if self.decimal_position != DecimalPosition::None {
                return false;
            }
            if self.scanned_type == Date && number_count > 2 {
                skip_blanks(text, &mut position);
                self.scanned_type = DateTime;
            } else if !matches!(self.scanned_type, Undefined | Time) {
                return false;
            } else {
                skip_blanks(text, &mut position);
                self.scanned_type = Time;
            }
            if self.time_position == 0 {
                self.time_position = self.strings.len();
            }
        }

        if skip_date_separator(text, &mut position) {
            if !matches!(self.scanned_type, Undefined | Date) {
                return false;
            }
            skip_blanks(text, &mut position);
            self.scanned_type = Date;
            let month = skip_month(text, &mut position);
            if self.month != 0 && month != 0 {
                return false;
            }
            if month != 0 {
                self.month = month;
                self.month_position = MonthPosition::End;
                if self.month < 0 {
                    skip_char(b'.', text, &mut position);
                }
                skip_blanks(text, &mut position);
            }
        }

        let month = skip_month(text, &mut position);
        if month != 0 {
            if self.month != 0 || !matches!(self.scanned_type, Undefined | Date) {
                return false;
            }
            self.scanned_type = Date;
            self.month = month;
            self.month_position = MonthPosition::End;
            if self.month < 0 {
                skip_char(b'.', text, &mut position);
            }
            skip_blanks(text, &mut position);
        }

        let before_am_pm = position;
        if skip_am_pm(text, &mut position) {
            if !matches!(self.scanned_type, Undefined | Time | DateTime) {
                return false;
            }
            // "6.78am" is not a time: keep the suffix.
            if self.scanned_type != Time
                && self.decimal_position == DecimalPosition::Middle
                && number_count == 2
            {
                position = before_am_pm;
            } else {
                skip_blanks(text, &mut position);
                if self.scanned_type != DateTime {
                    self.scanned_type = Time;
                }
            }
        }

        if self.open_parenthesis && skip_char(b')', text, &mut position) {
            if self.scanned_type != Currency {
                return false;
            }
            self.open_parenthesis = false;
            skip_blanks(text, &mut position);
        }

        if position < text.len() && matches!(self.scanned_type, Date | DateTime) {
            let before_day_name = position;
            if skip_string(locale::LONG_DATE_DAY_OF_WEEK_SEPARATOR, text, &mut position) {
                skip_blanks(text, &mut position);
            }
            let day_of_week = skip_name(&locale::DAYS_FROM_SUNDAY, text, &mut position);
            if day_of_week == 0 {
                position = before_day_name;
            } else if position < text.len() {
                if day_of_week < 0 {
                    skip_char(b'.', text, &mut position);
                }
                skip_blanks(text, &mut position);
            }
        }

        position >= text.len()
    }

    fn skip_number(&self, string_index: &mut usize, number_index: &mut usize) -> bool {
        if *string_index < self.strings.len() && self.is_number[*string_index] {
            *number_index += 1;
            *string_index += 1;
            true
        } else {
            false
        }
    }

    fn recognise(&mut self, text: &[u8]) -> Result<bool, FormatError> {
        use InputType::{Date, Fraction, Number, Scientific, Time, Undefined};
        self.divide(text);
        let string_count = self.strings.len();
        let number_count = self.numbers.len();
        if string_count >= MAX_STRINGS {
            return Ok(false);
        }
        // Without a number the input is text or a boolean, neither of which is a limit.
        if number_count == 0 {
            return Ok(false);
        }

        let mut i = 0;
        let mut j = 0;
        if !self.skip_number(&mut i, &mut j) {
            if !self.scan_start_string(i) {
                return Ok(false);
            }
            i += 1;
            if number_count >= 3 && self.decimal_position == DecimalPosition::Start {
                return Ok(false);
            }
        }
        self.skip_number(&mut i, &mut j);
        match number_count {
            1 => {}
            2 => {
                if !self.scan_mid_string(i)? {
                    return Ok(false);
                }
                i += 1;
                self.skip_number(&mut i, &mut j);
            }
            3 => {
                if !self.scan_mid_string(i)? {
                    return Ok(false);
                }
                i += 1;
                if self.scanned_type == Scientific {
                    return Ok(false);
                }
                self.skip_number(&mut i, &mut j);
                if !self.scan_mid_string(i)? {
                    return Ok(false);
                }
                i += 1;
                self.skip_number(&mut i, &mut j);
            }
            _ => {
                if !self.scan_mid_string(i)? {
                    return Ok(false);
                }
                i += 1;
                // Group separators first, then the remaining parts of a long date, time or number.
                let mut previous_thousand_count = usize::MAX;
                while previous_thousand_count != self.thousand_count && j < number_count - 1 {
                    previous_thousand_count = self.thousand_count;
                    if !self.scan_inner_mid_string(&mut i, &mut j)? {
                        return Ok(false);
                    }
                }
                if matches!(self.scanned_type, Date | Time | Undefined) {
                    let remaining = (number_count - 1).saturating_sub(j);
                    for _ in 0..remaining {
                        if !self.scan_inner_mid_string(&mut i, &mut j)? {
                            return Ok(false);
                        }
                    }
                }
                self.skip_number(&mut i, &mut j);
            }
        }
        if i < string_count && !self.scan_end_string(i) {
            return Ok(false);
        }
        if number_count >= 3
            && self.scanned_type == Fraction
            && self.decimal_position != DecimalPosition::None
        {
            return Ok(false);
        }
        if self.scanned_type == Undefined {
            self.scanned_type = Number;
        }
        Ok(true)
    }

    fn scan_inner_mid_string(
        &mut self,
        string_index: &mut usize,
        number_index: &mut usize,
    ) -> Result<bool, FormatError> {
        if self.scanned_type == InputType::Scientific {
            return Ok(false);
        }
        self.skip_number(string_index, number_index);
        if *string_index < self.strings.len() && !self.scan_mid_string(*string_index)? {
            return Ok(false);
        }
        *string_index += 1;
        Ok(true)
    }

    fn check_time(&self, first: usize, count: usize) -> Result<(), FormatError> {
        let decimal_in_middle = self.decimal_position == DecimalPosition::Middle;
        let mut index = first;
        if !(decimal_in_middle && (count == 3 || count == 2)) && index - first < count {
            to_int(self.number_string(index))?;
            index += 1;
        }
        if !(decimal_in_middle && count == 2) && index - first < count {
            to_int(self.number_string(index))?;
            index += 1;
        }
        if index - first < count {
            to_int(self.number_string(index))?;
        }
        Ok(())
    }

    fn check_date(&mut self) -> Result<Option<usize>, FormatError> {
        let valid = match (self.numbers.len(), self.month_position) {
            (0, position) => (position != MonthPosition::None).then_some(0),
            (1, MonthPosition::None | MonthPosition::Start) => Some(1),
            (1, _) => None,
            (2, MonthPosition::End) => None,
            (2, _) => Some(2),
            (_, MonthPosition::None) => {
                let mut count = 3;
                if self.time_position > 1 {
                    let first_time_number = self
                        .numbers
                        .iter()
                        .position(|&string| string == self.time_position - 2);
                    count = first_time_number.unwrap_or(count);
                }
                self.may_be_iso8601()?;
                Some(count)
            }
            (_, MonthPosition::Start) => Some(2),
            (_, _) => None,
        };
        Ok(valid)
    }

    fn number(&mut self) -> Result<Option<f64>, FormatError> {
        use InputType::{Currency, Date, DateTime, Number, Percent, Scientific, Time};
        let number_count = self.numbers.len();
        let has_month = self.month != 0;
        let has_decimals = self.decimal_position != DecimalPosition::None;
        if self.open_parenthesis || (self.sign != 0 && matches!(self.scanned_type, Date | DateTime))
        {
            return Ok(None);
        }
        let plausible = match self.scanned_type {
            Percent | Currency | Number => {
                number_count
                    == match self.decimal_position {
                        DecimalPosition::Start => 1,
                        DecimalPosition::Middle => self.thousand_count + 2,
                        _ => self.thousand_count + 1,
                    }
            }
            Scientific => {
                number_count
                    == match self.decimal_position {
                        DecimalPosition::Start => 2,
                        DecimalPosition::Middle => self.thousand_count + 3,
                        _ => self.thousand_count + 2,
                    }
            }
            Date => number_count <= if has_month { 2 } else { 3 },
            Time => number_count <= if has_decimals { 4 } else { 3 },
            DateTime => number_count <= 5 + usize::from(!has_month) + usize::from(has_decimals),
            _ => true,
        };
        if !plausible {
            return Ok(None);
        }

        match self.scanned_type {
            Number | Scientific => {}
            Time => {
                self.check_time(0, number_count)?;
                return Ok(None);
            }
            Date => {
                self.check_date()?;
                return Ok(None);
            }
            DateTime => {
                if let Some(date_numbers) = self.check_date()? {
                    self.check_time(date_numbers, number_count.wrapping_sub(date_numbers))?;
                }
                return Ok(None);
            }
            _ => return Ok(None),
        }

        let scientific = self.scanned_type == Scientific;
        let mut digits: Text = Vec::new();
        if self.decimal_position == DecimalPosition::Start {
            digits.extend_from_slice(b"0.");
        }
        let mut index = 0;
        while index <= self.thousand_count {
            digits.extend_from_slice(self.number_string(index));
            index += 1;
        }
        if self.decimal_position == DecimalPosition::Middle && index < number_count {
            digits.push(b'.');
            let stop = number_count - usize::from(scientific);
            while index < stop {
                digits.extend_from_slice(self.number_string(index));
                index += 1;
            }
        }
        let mut value = if scientific {
            digits.push(b'E');
            if self.exponent_sign == -1 {
                digits.push(b'-');
            }
            digits.extend_from_slice(self.number_string(number_count - 1));
            match scientific_to_double(&digits) {
                Some(value) => value,
                None => return Ok(None),
            }
        } else {
            string_to_double(&digits)
        };
        if self.sign < 0 {
            value = -value;
        }
        Ok(value.is_finite().then_some(value))
    }
}

fn skip_thousands(text: &[u8], position: &mut usize, symbol: &mut Text) -> bool {
    let mut found = false;
    let mut here = *position;
    let mut group_digits = None;
    while let Some(&byte) = text.get(here).filter(|&&byte| byte != 0) {
        match group_digits {
            None if byte == locale::THOUSAND_SEPARATOR => group_digits = Some(0),
            Some(count) if is_digit(byte) => {
                symbol.push(byte);
                group_digits = Some(count + 1).filter(|&count| count < 3);
                found |= group_digits.is_none();
            }
            _ => {
                *position = here;
                return found;
            }
        }
        here += 1;
    }
    // Only a group cut short by the end of the input is put back.
    if let Some(count) = group_digits {
        symbol.truncate(symbol.len() - count);
        here -= count + 1;
    }
    *position = here;
    found
}

#[cfg(test)]
mod tests {
    use super::*;

    fn limit(text: &str) -> Result<Option<f64>, FormatError> {
        parse_limit(text.as_bytes())
    }

    #[test]
    fn plain_and_scientific_numbers_are_limits() {
        assert_eq!(limit("100"), Ok(Some(100.0)));
        assert_eq!(limit("-1.5"), Ok(Some(-1.5)));
        assert_eq!(limit(".5"), Ok(Some(0.5)));
        assert_eq!(limit("1,000"), Ok(Some(1000.0)));
        assert_eq!(limit("1e3"), Ok(Some(1000.0)));
        assert_eq!(limit("1.5E-2"), Ok(Some(0.015)));
        assert_eq!(limit("(5)"), Ok(Some(-5.0)));
        assert_eq!(limit("0e999"), Ok(Some(0.0)));
    }

    #[test]
    fn other_inputs_are_not_limits() {
        for text in [
            "", "abc", "true", "$5", "5%", "1/2/2020", "10:30", "1 2/3", "(5", "1e999", "1e-999",
        ] {
            assert_eq!(limit(text), Ok(None), "{text}");
        }
    }

    #[test]
    fn limits_that_round_up_to_the_smallest_normal_number_underflow() {
        let smallest = Ok(Some(f64::MIN_POSITIVE));
        assert_eq!(limit("2.2250738585072014E-308"), smallest);
        assert_eq!(limit("2.2250738585072013E-308"), smallest);
        assert_eq!(limit("2.22507385850720125958E-308"), smallest);
        assert_eq!(limit("2.22507385850720125957E-308"), Ok(None));
        assert_eq!(limit("2.2250738585072012E-308"), Ok(None));
        assert_eq!(limit("0.00022250738585072012E-304"), Ok(None));
        assert_eq!(limit("22250738585072013E-324"), smallest);
    }

    #[test]
    fn multiply_decimal_keeps_the_decimal_point() {
        assert_eq!(multiply_decimal(b"12.5", 4), b"50.0");
        assert_eq!(multiply_decimal(b"0.999", 1000), b"999.000");
        assert_eq!(multiply_decimal(b"9", u64::MAX), b"166020696663385964535");
    }

    #[test]
    fn oversized_date_and_time_parts_fail() {
        assert_eq!(limit("99999999999-1-1"), Err(FormatError));
        assert_eq!(limit("99999999999:00"), Err(FormatError));
    }
}
