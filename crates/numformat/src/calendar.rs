//! Gregorian date arithmetic on Excel serial numbers.
//!
//! Serial numbers count days from 1899-12-30. Results are clamped to 0000-01-01 and 9999-12-31, and
//! years are 16-bit quantities whose wrap-around is part of the results at those limits. Weeks are
//! numbered like ISO weeks: they start on Monday and the first week of a year has at least four
//! days.

use crate::math::to_i64;

const DAYS_IN_MONTH: [u16; 12] = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
const MAX_DAYS: i64 = 3636532;

// Days from the Unix epoch to the serial number origin; date formats add it to their value first.
pub(crate) const EPOCH_OFFSET: f64 = -25569.0;

fn is_leap_year(year: u16) -> bool {
    (year.is_multiple_of(4) && !year.is_multiple_of(100)) || year.is_multiple_of(400)
}

fn days_in_month(month: u16, year: u16) -> u16 {
    let days = DAYS_IN_MONTH
        .get(usize::from(month.wrapping_sub(1)))
        .copied()
        .unwrap_or(0);
    if month == 2 && is_leap_year(year) {
        days + 1
    } else {
        days
    }
}

fn date_to_days(day: u16, month: u16, year: u16) -> i64 {
    let previous_year = i32::from(year) - 1;
    let mut days = (u64::from(year).wrapping_sub(1)).wrapping_mul(365) as i64;
    days = days.wrapping_add(i64::from(
        previous_year / 4 - previous_year / 100 + previous_year / 400,
    ));
    for month in 1..month {
        days += i64::from(days_in_month(month, year));
    }
    days + i64::from(day)
}

fn days_to_date(days: i64) -> Date {
    let mut correction = 0i64;
    let mut year;
    let mut remaining;
    loop {
        year = (days / 365 - correction) as u16;
        let previous_year = i32::from(year) - 1;
        remaining = days.wrapping_sub((u64::from(year).wrapping_sub(1)).wrapping_mul(365) as i64);
        remaining -= i64::from(previous_year / 4 - previous_year / 100 + previous_year / 400);
        if remaining < 1 {
            correction += 1;
        } else if remaining > 365 && (remaining != 366 || !is_leap_year(year)) {
            correction -= 1;
        } else {
            break;
        }
    }
    let mut month = 1u16;
    while month < 12 && remaining as u64 > u64::from(days_in_month(month, year)) {
        remaining -= i64::from(days_in_month(month, year));
        month += 1;
    }
    Date::new(remaining as u16, month, year)
}

/// A calendar date packed as `day + 100 * month + 10000 * year`.
#[derive(Clone, Copy)]
struct Date(u32);

impl Date {
    fn new(day: u16, month: u16, year: u16) -> Date {
        Date(u32::from(day % 100) + u32::from(month % 100) * 100 + u32::from(year % 10000) * 10000)
    }

    fn day(self) -> u16 {
        (self.0 % 100) as u16
    }

    fn month(self) -> u16 {
        ((self.0 / 100) % 100) as u16
    }

    fn year(self) -> u16 {
        (self.0 / 10000) as u16
    }

    fn days(self) -> i64 {
        date_to_days(self.day(), self.month(), self.year())
    }

    fn add_days(self, days: i64) -> Date {
        let total = self.days().wrapping_add(days);
        if total > MAX_DAYS {
            Date::new(31, 12, 9999)
        } else if total <= 0 {
            Date(101)
        } else {
            days_to_date(total)
        }
    }

    fn previous(self) -> Date {
        let days = self.days();
        if days > 1 {
            days_to_date(days - 1)
        } else {
            self
        }
    }

    fn day_of_week(self) -> u16 {
        ((self.days().wrapping_sub(1)) as u64 % 7) as u16
    }

    fn day_of_year(self) -> u16 {
        (1..self.month()).fold(self.day(), |day, month| {
            day.wrapping_add(days_in_month(month, self.year()))
        })
    }

    fn week_of_year(self) -> u16 {
        const MIN_DAYS_IN_FIRST_WEEK: i32 = 4;
        let mut date = self;
        // Week 53 defers to the Sunday ending the week, which lies in the next year at most once.
        for _ in 0..4 {
            let first_weekday = i32::from(Date::new(1, 1, date.year()).day_of_week());
            let day_of_year = i32::from(date.day_of_year() as i16) - 1;
            let mut week = if first_weekday < MIN_DAYS_IN_FIRST_WEEK {
                1
            } else if first_weekday == MIN_DAYS_IN_FIRST_WEEK {
                53
            } else if first_weekday == MIN_DAYS_IN_FIRST_WEEK + 1 {
                if is_leap_year(Date::new(1, 1, date.year().wrapping_sub(1)).year()) {
                    53
                } else {
                    52
                }
            } else {
                52
            };
            if week == 1 || day_of_year + first_weekday > 6 {
                if week == 1 {
                    week += (day_of_year + first_weekday) / 7;
                } else {
                    week = (day_of_year + first_weekday) / 7;
                }
                if week == 53 {
                    let sunday = date.days() + 6 - i64::from(date.day_of_week());
                    date = days_to_date(sunday);
                    continue;
                }
            }
            return week as u16;
        }
        53
    }
}

/// The calendar fields that date formats can display, for the date a number of days after the
/// Unix epoch: a serial number plus [`EPOCH_OFFSET`].
#[derive(Clone, Copy, Debug)]
pub(crate) struct CalendarDate {
    pub(crate) day: u16,
    pub(crate) month_index: u16,
    pub(crate) year: u16,
    pub(crate) days_since_sunday: u16,
    pub(crate) week_of_year: u16,
}

impl CalendarDate {
    pub(crate) fn from_epoch_days(epoch_days: f64) -> CalendarDate {
        let date = add_fractional_days(Date::new(30, 12, 1899), epoch_days - EPOCH_OFFSET);
        CalendarDate {
            day: date.day(),
            month_index: date.month().wrapping_sub(1),
            year: date.year(),
            days_since_sunday: (date.day_of_week() + 1) % 7,
            week_of_year: date.week_of_year(),
        }
    }
}

fn add_fractional_days(date: Date, days: f64) -> Date {
    let (whole, fraction) = if days < 0.0 {
        let whole = days.ceil();
        (whole, if whole <= days { 0.0 } else { days - whole })
    } else {
        let whole = days.floor();
        (whole, if whole >= days { 0.0 } else { days - whole })
    };
    let date = date.add_days(to_i64(whole));
    if fraction == 0.0 {
        return date;
    }
    // Times count in hundredths of a second; a negative time of day belongs to the previous date.
    let milliseconds = to_i64(fraction * 86_400_000.0) as i32;
    if milliseconds < 0 && milliseconds.wrapping_neg() / 10 != 0 {
        date.previous()
    } else {
        date
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn fields(serial: f64) -> (u16, u16, u16, u16, u16) {
        let date = CalendarDate::from_epoch_days(serial + EPOCH_OFFSET);
        (
            date.year,
            date.month_index + 1,
            date.day,
            date.days_since_sunday,
            date.week_of_year,
        )
    }

    #[test]
    fn serial_origin_and_leap_day() {
        assert_eq!(fields(0.0), (1899, 12, 30, 6, 52));
        assert_eq!(fields(1.0), (1899, 12, 31, 0, 52));
        assert_eq!(fields(60.0), (1900, 2, 28, 3, 9));
        assert_eq!(fields(61.0), (1900, 3, 1, 4, 9));
        assert_eq!(fields(45000.75), (2023, 3, 15, 3, 11));
    }

    #[test]
    fn negative_fractions_fall_on_the_previous_day() {
        assert_eq!(fields(-0.5), (1899, 12, 29, 5, 52));
        assert_eq!(fields(-1.0), (1899, 12, 29, 5, 52));
    }

    #[test]
    fn out_of_range_serials_are_clamped() {
        assert_eq!(fields(2958465.0), (9999, 12, 31, 5, 52));
        assert_eq!(fields(1e7), (9999, 12, 31, 5, 52));
        assert_eq!(fields(-1e7).0, 0);
        assert_eq!(fields(f64::NAN).0, 0);
        assert_eq!(fields(1e300).0, 0);
    }
}
