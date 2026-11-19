//! Excel's serial numbers for dates and times: a count of days with the time of day as the fraction.
//!
//! A workbook counts from 1900 or from 1904. In the 1900 system serial number 1 is 1900-01-01 and 60 is a
//! 1900-02-29 that never was, so from 61 on the count is the days since 1899-12-30, which is also where the
//! numbers below 1 are counted from.

use std::fmt;

const MICROS_PER_DAY: i64 = 86_400_000_000;
const DAYS_FROM_1899_12_30_TO_1970: i64 = 25569;
const DAYS_FROM_1904_TO_1970: i64 = 24107;

/// The days, since 1970-01-01, whose serial number is one less than their distance from 1899-12-30:
/// 1900-01-01 to 1900-02-28.
const JANUARY_AND_FEBRUARY_1900: std::ops::Range<i64> = -25567..-25508;

/// No type has a value this many days from serial number zero, and below it the day arithmetic is exact.
const MAX_SERIAL_DAYS: f64 = 4_294_967_296.0;

/// How far a time of day may be from a whole second or millisecond and still be taken for it, in
/// microseconds. Serial numbers up to 65,536 are at most 0.63 microseconds apart, and a time that really is a
/// microsecond off is at least 0.68 away.
const NOISE_MICROS: f64 = 0.65;

/// The first TIMESTAMP DuckDB has, 290309-12-22 (BC) 00:00:00, and the last, 294247-01-10 04:00:54.775806.
const TIMESTAMP_RANGE: std::ops::RangeInclusive<i128> =
    -9_223_372_022_400_000_000..=(i64::MAX as i128 - 1);

/// The first DATE DuckDB has, 5877642-06-25 (BC), and the last, 5881580-07-10.
const DATE_RANGE: std::ops::RangeInclusive<i128> = -(i32::MAX as i128 - 1)..=(i32::MAX as i128 - 1);

/// What DuckDB stores for the DATE and the TIMESTAMP `infinity`; `-infinity` is the negative of it.
const INFINITE_DATE: i32 = i32::MAX;
const INFINITE_TIMESTAMP: i64 = i64::MAX;

/// The day a workbook counts its serial numbers from.
#[derive(Debug, Default, Clone, Copy, PartialEq, Eq)]
pub enum DateSystem {
    #[default]
    Since1900,
    /// Serial number 0 is 1904-01-01.
    Since1904,
}

/// Why a serial number is not a value of a type.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum SerialError {
    NotFinite,
    OutOfRange,
}

impl fmt::Display for SerialError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(match self {
            SerialError::NotFinite => "the value is not a finite number",
            SerialError::OutOfRange => "the value is out of range",
        })
    }
}

fn time_of_day_micros(serial: f64, day: f64) -> i64 {
    let exact = (serial - day) * MICROS_PER_DAY as f64;
    // Above 65,536 neighbouring serial numbers are further apart than the noise; the factor covers the
    // rounding of the product.
    let step = (serial.abs().next_up() - serial.abs()) * MICROS_PER_DAY as f64;
    let tolerance = (step * 1.01).max(NOISE_MICROS);
    for unit in [1e6, 1e3] {
        let rounded = (exact / unit).round() * unit;
        if (rounded - exact).abs() <= tolerance {
            return rounded as i64;
        }
    }
    exact.round() as i64
}

fn serial_to_micros(serial: f64, system: DateSystem) -> Result<i128, SerialError> {
    if !serial.is_finite() {
        return Err(SerialError::NotFinite);
    }
    let day = serial.floor();
    if day.abs() > MAX_SERIAL_DAYS {
        return Err(SerialError::OutOfRange);
    }
    let micros_per_day = i128::from(MICROS_PER_DAY);
    let since_zero = day as i128 * micros_per_day + i128::from(time_of_day_micros(serial, day));
    let days_to_1970 = match system {
        DateSystem::Since1904 => DAYS_FROM_1904_TO_1970,
        // The rounded time decides, so that a number a hair below 60 is 60.
        DateSystem::Since1900 if (micros_per_day..60 * micros_per_day).contains(&since_zero) => {
            DAYS_FROM_1899_12_30_TO_1970 - 1
        }
        DateSystem::Since1900 => DAYS_FROM_1899_12_30_TO_1970,
    };
    Ok(since_zero - i128::from(days_to_1970) * micros_per_day)
}

/// The TIMESTAMP, in microseconds since 1970-01-01, of a serial number: its day and its time of day rounded
/// to the microsecond, where a whole second or millisecond within the noise of a DOUBLE is that one. An
/// infinite number is the infinite TIMESTAMP of its sign.
pub fn serial_to_timestamp(serial: f64, system: DateSystem) -> Result<i64, SerialError> {
    if serial.is_infinite() {
        return Ok(if serial > 0.0 {
            INFINITE_TIMESTAMP
        } else {
            -INFINITE_TIMESTAMP
        });
    }
    let micros = serial_to_micros(serial, system)?;
    if !TIMESTAMP_RANGE.contains(&micros) {
        return Err(SerialError::OutOfRange);
    }
    i64::try_from(micros).map_err(|_| SerialError::OutOfRange)
}

/// The DATE, in days since 1970-01-01, of a serial number: the day of its [`serial_to_timestamp`]. Serial
/// numbers 59 and 60 of the 1900 system are both 1900-02-28.
pub fn serial_to_date(serial: f64, system: DateSystem) -> Result<i32, SerialError> {
    if serial.is_infinite() {
        return Ok(if serial > 0.0 {
            INFINITE_DATE
        } else {
            -INFINITE_DATE
        });
    }
    let days = serial_to_micros(serial, system)?.div_euclid(i128::from(MICROS_PER_DAY));
    if !DATE_RANGE.contains(&days) {
        return Err(SerialError::OutOfRange);
    }
    i32::try_from(days).map_err(|_| SerialError::OutOfRange)
}

/// The TIME, in microseconds since midnight, of a serial number: the time of day of its
/// [`serial_to_timestamp`], except that exactly 1 is 24:00:00, which is written as that number.
pub fn serial_to_time(serial: f64) -> Result<i64, SerialError> {
    if serial == 1.0 {
        return Ok(MICROS_PER_DAY);
    }
    let micros = serial_to_micros(serial, DateSystem::Since1900)?;
    i64::try_from(micros.rem_euclid(i128::from(MICROS_PER_DAY)))
        .map_err(|_| SerialError::OutOfRange)
}

fn days_from_serial_zero_to_1970(epoch_day: i64, system: DateSystem) -> i64 {
    match system {
        DateSystem::Since1904 => DAYS_FROM_1904_TO_1970,
        DateSystem::Since1900 if JANUARY_AND_FEBRUARY_1900.contains(&epoch_day) => {
            DAYS_FROM_1899_12_30_TO_1970 - 1
        }
        DateSystem::Since1900 => DAYS_FROM_1899_12_30_TO_1970,
    }
}

/// The serial number of a TIMESTAMP in microseconds since 1970-01-01, infinite for an infinite TIMESTAMP.
/// In the 1900 system 1899-12-31 has none and shares those of 1900-01-01.
pub fn epoch_micros_to_serial(micros: i64, system: DateSystem) -> f64 {
    if micros.unsigned_abs() == INFINITE_TIMESTAMP.unsigned_abs() {
        return f64::INFINITY.copysign(micros as f64);
    }
    let epoch_day = micros.div_euclid(MICROS_PER_DAY);
    let day = epoch_day + days_from_serial_zero_to_1970(epoch_day, system);
    // The whole day is added last, which leaves a single rounding that reading the number back undoes.
    day as f64 + micros.rem_euclid(MICROS_PER_DAY) as f64 / MICROS_PER_DAY as f64
}

/// The serial number of a DATE in days since 1970-01-01, infinite for an infinite DATE. In the 1900 system
/// 1899-12-31 has none and shares that of 1900-01-01.
pub fn epoch_days_to_serial(days: i32, system: DateSystem) -> f64 {
    if days.unsigned_abs() == INFINITE_DATE.unsigned_abs() {
        return f64::INFINITY.copysign(f64::from(days));
    }
    let epoch_day = i64::from(days);
    (epoch_day + days_from_serial_zero_to_1970(epoch_day, system)) as f64
}

/// The fraction of a day a TIME in microseconds since midnight is.
pub fn time_micros_to_serial(micros: i64) -> f64 {
    micros as f64 / MICROS_PER_DAY as f64
}

#[cfg(test)]
mod tests {
    use super::*;
    use DateSystem::{Since1900, Since1904};

    const DAY: i64 = MICROS_PER_DAY;
    const SECOND: i64 = 1_000_000;

    /// The days since 1970-01-01 of a day of the Gregorian calendar.
    fn days(year: i64, month: i64, day: i64) -> i64 {
        let year = if month <= 2 { year - 1 } else { year };
        let era = year.div_euclid(400);
        let year_of_era = year.rem_euclid(400);
        let day_of_year = (153 * (month + if month > 2 { -3 } else { 9 }) + 2) / 5 + day - 1;
        let day_of_era = year_of_era * 365 + year_of_era / 4 - year_of_era / 100 + day_of_year;
        era * 146_097 + day_of_era - 719_468
    }

    fn timestamp(serial: f64) -> i64 {
        serial_to_timestamp(serial, Since1900).unwrap()
    }

    fn time_serial(hours: i64, minutes: i64, seconds: i64, micros: i64) -> f64 {
        ((hours * 3600 + minutes * 60 + seconds) * SECOND + micros) as f64 / DAY as f64
    }

    #[test]
    fn serial_numbers_of_timestamps_dates_and_times() {
        assert_eq!(epoch_micros_to_serial(0, Since1900), 25569.0);
        assert_eq!(
            epoch_micros_to_serial(1_710_505_845_123_000, Since1900),
            45366.52135559028
        );
        assert_eq!(epoch_micros_to_serial(-DAY / 2, Since1900), 25568.5);
        assert_eq!(
            epoch_micros_to_serial(i64::MAX - 1, Since1900),
            106777560.16730064
        );
        assert_eq!(epoch_days_to_serial(0, Since1900), 25569.0);
        assert_eq!(epoch_days_to_serial(19797, Since1900), 45366.0);
        assert_eq!(epoch_days_to_serial(-25569, Since1900), 0.0);
        assert_eq!(epoch_days_to_serial(i32::MAX - 1, Since1900), 2147509215.0);
        assert_eq!(time_micros_to_serial(0), 0.0);
        assert_eq!(time_micros_to_serial(29_700_000_000), 0.34375);
        assert_eq!(time_micros_to_serial(DAY), 1.0);
    }

    #[test]
    fn serial_numbers_from_61_on_count_from_1899_12_30() {
        assert_eq!(timestamp(25569.0), 0);
        assert_eq!(timestamp(25570.5), DAY + DAY / 2);
        assert_eq!(timestamp(44197.0), days(2021, 1, 1) * DAY);
        assert_eq!(timestamp(61.0), days(1900, 3, 1) * DAY);
        assert_eq!(timestamp(2_958_465.0), days(9999, 12, 31) * DAY);
        assert_eq!(
            serial_to_date(45000.0, Since1900),
            Ok(days(2023, 3, 15) as i32)
        );
        assert_eq!(
            serial_to_date(45000.99, Since1900),
            Ok(days(2023, 3, 15) as i32)
        );
    }

    #[test]
    fn serial_numbers_below_61_are_the_days_excel_shows() {
        assert_eq!(timestamp(1.0), days(1900, 1, 1) * DAY);
        assert_eq!(timestamp(31.0), days(1900, 1, 31) * DAY);
        assert_eq!(timestamp(59.0), days(1900, 2, 28) * DAY);
        assert_eq!(timestamp(59.75), days(1900, 2, 28) * DAY + 3 * DAY / 4);
        // The day that never was is the last day before it.
        assert_eq!(timestamp(60.0), days(1900, 2, 28) * DAY);
        assert_eq!(timestamp(60.5), days(1900, 2, 28) * DAY + DAY / 2);
        for (serial, day) in [(1.0, 1), (2.5, 2), (31.0, 31)] {
            assert_eq!(
                serial_to_date(serial, Since1900),
                Ok(days(1900, 1, day) as i32)
            );
        }
        assert_eq!(
            serial_to_date(60.0, Since1900),
            Ok(days(1900, 2, 28) as i32)
        );
    }

    #[test]
    fn serial_numbers_below_1_count_from_1899_12_30() {
        assert_eq!(timestamp(0.0), days(1899, 12, 30) * DAY);
        assert_eq!(timestamp(0.5), days(1899, 12, 30) * DAY + DAY / 2);
        assert_eq!(timestamp(-1.0), days(1899, 12, 29) * DAY);
        assert_eq!(timestamp(-0.25), days(1899, 12, 29) * DAY + 3 * DAY / 4);
        assert_eq!(serial_to_date(-100.0, Since1900), Ok(-25669));
    }

    #[test]
    fn the_rounded_time_decides_the_day() {
        assert_eq!(timestamp(59.999_999_999_999), timestamp(60.0));
        assert_eq!(timestamp(0.999_999_999_999), timestamp(1.0));
        assert_eq!(timestamp(60.999_999_999_999), timestamp(61.0));
        // The last millisecond of 1900-02-28 is still that day.
        let serial = 59.0 + time_serial(23, 59, 59, 999_000);
        assert_eq!(timestamp(serial), days(1900, 3, 1) * DAY - 1000);
    }

    #[test]
    fn the_1904_system_counts_from_1904() {
        let timestamp = |serial: f64| serial_to_timestamp(serial, Since1904).unwrap();
        assert_eq!(timestamp(0.0), days(1904, 1, 1) * DAY);
        assert_eq!(timestamp(1.0), days(1904, 1, 2) * DAY);
        assert_eq!(timestamp(59.5), days(1904, 2, 29) * DAY + DAY / 2);
        assert_eq!(timestamp(60.0), days(1904, 3, 1) * DAY);
        assert_eq!(timestamp(43538.5), days(2023, 3, 15) * DAY + DAY / 2);
        assert_eq!(timestamp(-1.0), days(1903, 12, 31) * DAY);
        assert_eq!(
            serial_to_date(45000.0 - 1462.0, Since1904),
            serial_to_date(45000.0, Since1900)
        );
    }

    #[test]
    fn the_time_of_day_is_rounded_to_the_microsecond() {
        let day = days(2024, 3, 15) * DAY;
        let noon = day + DAY / 2;
        for (hours, minutes, seconds, micros) in [
            (12, 34, 56, 789_000),
            (12, 34, 56, 789_123),
            (0, 0, 0, 1000),
            (0, 0, 0, 999_000),
            (0, 0, 0, 1),
            (23, 59, 59, 999_999),
            (7, 7, 7, 700_007),
        ] {
            let serial = 45366.0 + time_serial(hours, minutes, seconds, micros);
            let expected = day + (hours * 3600 + minutes * 60 + seconds) * SECOND + micros;
            assert_eq!(timestamp(serial), expected, "{serial}");
        }
        // The DOUBLE next to a whole second or millisecond is that second or millisecond ...
        assert_eq!(timestamp(45366.5_f64.next_down()), noon);
        assert_eq!(timestamp(45366.5_f64.next_up()), noon);
        let below_the_millisecond = 45366.0 + time_serial(12, 34, 56, 789_000);
        assert_eq!(
            timestamp(below_the_millisecond.next_up()),
            timestamp(below_the_millisecond)
        );
        // ... and the microsecond next to it is not.
        assert_eq!(timestamp(45366.5 + 1e-6 / 86400.0), noon + 1);
        assert_eq!(timestamp(45366.5 - 1e-6 / 86400.0), noon - 1);
    }

    #[test]
    fn far_serial_numbers_are_the_roundest_time_they_can_stand_for() {
        let day = days(2100, 1, 1) * DAY;
        assert_eq!(timestamp(73051.5_f64.next_up()), day + DAY / 2);
        assert_eq!(
            timestamp(73051.0 + time_serial(12, 34, 56, 789_000)),
            day + 45_296_789_000
        );
        assert_eq!(
            timestamp(2_958_465.0 + time_serial(12, 34, 56, 789_000)),
            days(9999, 12, 31) * DAY + 45_296_789_000
        );
    }

    #[test]
    fn infinite_dates_and_timestamps_are_infinite_serial_numbers() {
        for system in [Since1900, Since1904] {
            assert_eq!(epoch_days_to_serial(i32::MAX, system), f64::INFINITY);
            assert_eq!(epoch_days_to_serial(-i32::MAX, system), f64::NEG_INFINITY);
            assert_eq!(epoch_micros_to_serial(i64::MAX, system), f64::INFINITY);
            assert_eq!(epoch_micros_to_serial(-i64::MAX, system), f64::NEG_INFINITY);
            assert_eq!(serial_to_date(f64::INFINITY, system), Ok(i32::MAX));
            assert_eq!(serial_to_date(f64::NEG_INFINITY, system), Ok(-i32::MAX));
            assert_eq!(serial_to_timestamp(f64::INFINITY, system), Ok(i64::MAX));
            assert_eq!(
                serial_to_timestamp(f64::NEG_INFINITY, system),
                Ok(-i64::MAX)
            );
        }
        // The values next to the infinite ones are the ends of the types.
        assert!(epoch_days_to_serial(-i32::MAX + 1, Since1900).is_finite());
        assert!(epoch_micros_to_serial(i64::MAX - 1, Since1900).is_finite());
        assert!(epoch_micros_to_serial(i64::MIN, Since1900).is_finite());
        assert!(epoch_days_to_serial(i32::MIN, Since1900).is_finite());
    }

    #[test]
    fn serial_numbers_that_are_no_numbers_are_errors() {
        for system in [Since1900, Since1904] {
            assert_eq!(
                serial_to_timestamp(f64::NAN, system),
                Err(SerialError::NotFinite)
            );
            assert_eq!(
                serial_to_date(f64::NAN, system),
                Err(SerialError::NotFinite)
            );
        }
        // A time of day is never infinite.
        for serial in [f64::NAN, f64::INFINITY, f64::NEG_INFINITY] {
            assert_eq!(serial_to_time(serial), Err(SerialError::NotFinite));
        }
        assert_eq!(
            SerialError::NotFinite.to_string(),
            "the value is not a finite number"
        );
        assert_eq!(
            SerialError::OutOfRange.to_string(),
            "the value is out of range"
        );
    }

    #[test]
    fn serial_numbers_out_of_range_are_errors() {
        let out_of_range = Err(SerialError::OutOfRange);
        for serial in [1e300, -1e300, 1e10, -1e10, 4_294_967_297.0] {
            assert_eq!(serial_to_timestamp(serial, Since1900), out_of_range);
            assert_eq!(
                serial_to_date(serial, Since1904),
                Err(SerialError::OutOfRange)
            );
            assert_eq!(serial_to_time(serial), Err(SerialError::OutOfRange));
        }
        // A billion days is no TIMESTAMP, but a DATE and a time of day.
        for serial in [-1e9, 1e9] {
            assert_eq!(serial_to_timestamp(serial, Since1900), out_of_range);
            assert!(serial_to_date(serial, Since1900).is_ok());
        }
        assert_eq!(serial_to_time(-1e9 + 0.25), Ok(DAY / 4));
        assert_eq!(serial_to_date(3e9, Since1900), Err(SerialError::OutOfRange));
        assert_eq!(
            serial_to_date(-3e9, Since1900),
            Err(SerialError::OutOfRange)
        );
    }

    #[test]
    fn the_ranges_end_where_the_types_end() {
        let first = *TIMESTAMP_RANGE.start() as i64;
        let first_serial = (first / DAY + 25569) as f64;
        assert_eq!(serial_to_timestamp(first_serial, Since1900), Ok(first));
        assert_eq!(
            serial_to_timestamp(first_serial - 1.0, Since1900),
            Err(SerialError::OutOfRange)
        );
        let last_serial = (i64::MAX / DAY + 25569) as f64;
        assert_eq!(
            serial_to_timestamp(last_serial, Since1900),
            Ok(i64::MAX / DAY * DAY)
        );
        assert_eq!(
            serial_to_timestamp(last_serial + 1.0, Since1900),
            Err(SerialError::OutOfRange)
        );

        let first = f64::from(-i32::MAX + 1);
        assert_eq!(
            serial_to_date(first + 25569.0, Since1900),
            Ok(-i32::MAX + 1)
        );
        assert_eq!(
            serial_to_date(first + 25568.0, Since1900),
            Err(SerialError::OutOfRange)
        );
        let last = f64::from(i32::MAX - 1);
        assert_eq!(serial_to_date(last + 25569.0, Since1900), Ok(i32::MAX - 1));
        assert_eq!(
            serial_to_date(last + 25570.0, Since1900),
            Err(SerialError::OutOfRange)
        );
    }

    #[test]
    fn times_are_the_time_of_day() {
        assert_eq!(serial_to_time(0.0), Ok(0));
        assert_eq!(serial_to_time(0.25), Ok(DAY / 4));
        assert_eq!(serial_to_time(45000.75), Ok(3 * DAY / 4));
        assert_eq!(serial_to_time(59.5), Ok(DAY / 2));
        assert_eq!(serial_to_time(-0.25), Ok(3 * DAY / 4));
        assert_eq!(
            serial_to_time(time_serial(12, 34, 56, 789_000)),
            Ok(45_296_789_000)
        );
        assert_eq!(serial_to_time(0.999_999_99), Ok(DAY - 864));
        assert_eq!(serial_to_time(0.999_999_999_999_9), Ok(0));
    }

    #[test]
    fn only_the_serial_number_1_is_midnight_at_the_end_of_the_day() {
        assert_eq!(serial_to_time(1.0), Ok(DAY));
        assert_eq!(serial_to_time(2.0), Ok(0));
        assert_eq!(serial_to_time(1.0_f64.next_up()), Ok(0));
        assert_eq!(serial_to_time(1.0_f64.next_down()), Ok(0));
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
    fn every_microsecond_of_four_centuries_reads_back_as_written() {
        // Serial numbers below 65,536 in size: 1720-07-25 to 2079-06-05.
        let first = days(1720, 7, 26) * DAY;
        let last = days(2079, 6, 5) * DAY;
        let day_without_serial_numbers = days(1899, 12, 31) * DAY..days(1900, 1, 1) * DAY;
        let mut next = sequence(0x9E37_79B9_7F4A_7C15);
        for _ in 0..200_000 {
            let micros = first + (next() % (last - first) as u64) as i64;
            if day_without_serial_numbers.contains(&micros) {
                continue;
            }
            let serial = epoch_micros_to_serial(micros, Since1900);
            assert_eq!(
                serial_to_timestamp(serial, Since1900),
                Ok(micros),
                "{micros}"
            );
        }
        for micros in [
            first,
            last - 1,
            0,
            1,
            -1,
            999_999,
            1_710_505_845_123_456,
            days(1900, 1, 1) * DAY,
            days(1900, 1, 1) * DAY + 1,
            days(1900, 3, 1) * DAY - 1,
            days(1900, 3, 1) * DAY,
            days(1899, 12, 31) * DAY - 1,
            days(1899, 12, 30) * DAY,
            days(1899, 12, 30) * DAY + 1,
        ] {
            let serial = epoch_micros_to_serial(micros, Since1900);
            assert_eq!(
                serial_to_timestamp(serial, Since1900),
                Ok(micros),
                "{micros}"
            );
        }
    }

    #[test]
    fn every_millisecond_of_any_year_reads_back_as_written() {
        let first = days(1, 1, 1) * DAY;
        let last = days(9999, 12, 31) * DAY + DAY;
        let mut next = sequence(0x2545_F491_4F6C_DD1D);
        for _ in 0..200_000 {
            let micros = first + (next() % (last - first) as u64) as i64;
            let millisecond = micros - micros.rem_euclid(1000);
            let serial = epoch_micros_to_serial(millisecond, Since1900);
            assert_eq!(
                serial_to_timestamp(serial, Since1900),
                Ok(millisecond),
                "{millisecond}"
            );
        }
    }

    #[test]
    fn dates_and_times_read_back_as_written() {
        for days in [
            -25569,
            0,
            19797,
            -30000,
            3_000_000,
            i32::MAX - 1,
            -i32::MAX + 1,
        ] {
            let serial = epoch_days_to_serial(days, Since1900);
            assert_eq!(serial_to_date(serial, Since1900), Ok(days), "{days}");
        }
        let mut next = sequence(0xD1B5_4A32_D192_ED03);
        for _ in 0..200_000 {
            let micros = (next() % (DAY as u64 + 1)) as i64;
            let serial = time_micros_to_serial(micros);
            assert_eq!(serial_to_time(serial), Ok(micros), "{micros}");
        }
        for micros in [0, 1, 29_700_000_000, DAY - 1, DAY] {
            let serial = time_micros_to_serial(micros);
            assert_eq!(serial_to_time(serial), Ok(micros), "{micros}");
        }
    }

    #[test]
    fn the_1904_system_is_written_as_it_is_read() {
        assert_eq!(
            epoch_days_to_serial(days(1904, 1, 1) as i32, Since1904),
            0.0
        );
        assert_eq!(
            epoch_days_to_serial(days(2023, 3, 15) as i32, Since1904),
            43538.0
        );
        assert_eq!(
            epoch_days_to_serial(days(1900, 1, 1) as i32, Since1904),
            -1460.0
        );
        assert_eq!(
            epoch_micros_to_serial(days(2023, 3, 15) * DAY + DAY / 2, Since1904),
            43538.5
        );
        assert_eq!(
            epoch_micros_to_serial(days(1903, 12, 31) * DAY + DAY / 4, Since1904),
            -0.75
        );
        let mut next = sequence(0x1234_5678_9ABC_DEF1);
        let first = days(1850, 1, 1) * DAY;
        let last = days(2079, 1, 1) * DAY;
        for _ in 0..100_000 {
            let micros = first + (next() % (last - first) as u64) as i64;
            let serial = epoch_micros_to_serial(micros, Since1904);
            assert_eq!(
                serial_to_timestamp(serial, Since1904),
                Ok(micros),
                "{micros}"
            );
            let day = micros.div_euclid(DAY) as i32;
            let serial = epoch_days_to_serial(day, Since1904);
            assert_eq!(serial_to_date(serial, Since1904), Ok(day), "{day}");
        }
    }

    #[test]
    fn days_of_january_and_february_1900_are_written_as_excel_counts_them() {
        assert_eq!(
            epoch_days_to_serial(days(1900, 1, 1) as i32, Since1900),
            1.0
        );
        assert_eq!(
            epoch_days_to_serial(days(1900, 1, 31) as i32, Since1900),
            31.0
        );
        assert_eq!(
            epoch_days_to_serial(days(1900, 2, 28) as i32, Since1900),
            59.0
        );
        assert_eq!(
            epoch_days_to_serial(days(1900, 3, 1) as i32, Since1900),
            61.0
        );
        assert_eq!(
            epoch_days_to_serial(days(1899, 12, 30) as i32, Since1900),
            0.0
        );
        assert_eq!(
            epoch_micros_to_serial(days(1900, 1, 1) * DAY, Since1900),
            1.0
        );
        assert_eq!(
            epoch_micros_to_serial(days(1900, 1, 1) * DAY + DAY / 2, Since1900),
            1.5
        );
        assert_eq!(
            epoch_micros_to_serial(days(1900, 2, 28) * DAY + DAY / 4, Since1900),
            59.25
        );
        assert_eq!(
            epoch_micros_to_serial(days(1900, 3, 1) * DAY, Since1900),
            61.0
        );
        assert_eq!(
            epoch_micros_to_serial(days(1900, 3, 1) * DAY - DAY / 4, Since1900),
            59.75
        );
        assert_eq!(
            epoch_micros_to_serial(days(1899, 12, 30) * DAY + DAY / 2, Since1900),
            0.5
        );
    }

    #[test]
    fn every_day_but_one_reads_back_as_written() {
        let day_without_a_serial = days(1899, 12, 31) as i32;
        for day in (days(1890, 1, 1) as i32)..(days(1910, 1, 1) as i32) {
            let read = serial_to_date(epoch_days_to_serial(day, Since1900), Since1900);
            let expected = if day == day_without_a_serial {
                day + 1
            } else {
                day
            };
            assert_eq!(read, Ok(expected), "{day}");

            let noon = i64::from(day) * DAY + DAY / 2;
            let read = serial_to_timestamp(epoch_micros_to_serial(noon, Since1900), Since1900);
            let expected = if day == day_without_a_serial {
                noon + DAY
            } else {
                noon
            };
            assert_eq!(read, Ok(expected), "{day}");
        }
        for micros in [
            days(1900, 1, 1) * DAY,
            days(1900, 3, 1) * DAY - 1,
            days(1900, 3, 1) * DAY,
        ] {
            let read = serial_to_timestamp(epoch_micros_to_serial(micros, Since1900), Since1900);
            assert_eq!(read, Ok(micros), "{micros}");
        }
    }
}
