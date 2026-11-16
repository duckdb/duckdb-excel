//! Locale data of the formatter. Only the en-US locale exists: format codes cannot select another.

pub(crate) const DECIMAL_SEPARATOR: u8 = b'.';
pub(crate) const THOUSAND_SEPARATOR: u8 = b',';
pub(crate) const DATE_SEPARATOR: u8 = b'/';
pub(crate) const TIME_SEPARATOR: u8 = b':';
pub(crate) const TIME_100SEC_SEPARATOR: u8 = b'.';

pub(crate) const LONG_DATE_DAY_OF_WEEK_SEPARATOR: &[u8] = b", ";
pub(crate) const LONG_DATE_DAY_SEPARATOR: &[u8] = b", ";
pub(crate) const LONG_DATE_MONTH_SEPARATOR: &[u8] = b" ";

pub(crate) const CURRENCY_SYMBOL: u8 = b'$';
pub(crate) const CURRENCY_BANK_SYMBOL: &[u8] = b"USD";

pub(crate) const TIME_AM: &[u8] = b"AM";
pub(crate) const TIME_PM: &[u8] = b"PM";

// Shown where a value is out of range for its format.
pub(crate) const ERROR_STRING: &[u8] = b"###";

/// Abbreviated and full names.
pub(crate) struct Name {
    pub(crate) abbreviation: &'static str,
    pub(crate) full: &'static str,
}

const fn name(abbreviation: &'static str, full: &'static str) -> Name {
    Name { abbreviation, full }
}

pub(crate) const MONTHS: [Name; 12] = [
    name("Jan", "January"),
    name("Feb", "February"),
    name("Mar", "March"),
    name("Apr", "April"),
    name("May", "May"),
    name("Jun", "June"),
    name("Jul", "July"),
    name("Aug", "August"),
    name("Sep", "September"),
    name("Oct", "October"),
    name("Nov", "November"),
    name("Dec", "December"),
];

pub(crate) const DAYS_FROM_SUNDAY: [Name; 7] = [
    name("Sun", "Sunday"),
    name("Mon", "Monday"),
    name("Tue", "Tuesday"),
    name("Wed", "Wednesday"),
    name("Thu", "Thursday"),
    name("Fri", "Friday"),
    name("Sat", "Saturday"),
];

pub(crate) const QUARTERS: [Name; 4] = [
    name("Q1", "1st quarter"),
    name("Q2", "2nd quarter"),
    name("Q3", "3rd quarter"),
    name("Q4", "4th quarter"),
];

// The calendar never leaves its first era, which has an abbreviation but no full name.
pub(crate) const ERA: Name = name("BC", "");
