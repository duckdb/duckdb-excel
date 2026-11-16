//! Excel's serial numbers for dates and times: days since 1899-12-30, with the time of day as the fraction.

use duckdb_neo::Result;

use crate::error::conversion;

const SECONDS_PER_DAY: f64 = 86400.0;
const MICROS_PER_SECOND: f64 = 1_000_000.0;
const MICROS_PER_DAY: i64 = 86_400_000_000;
const DAYS_BETWEEN_1900_AND_1970: f64 = 25569.0;

/// The microseconds since 1970-01-01 of a serial number. A serial number within a millisecond of a whole
/// second is that second; one outside the range of a timestamp is the nearest end of the range.
pub fn serial_to_epoch_micros(serial: f64) -> i64 {
    let mut seconds = serial * SECONDS_PER_DAY;
    if (seconds - seconds.round()).abs() < 1e-3 {
        seconds = seconds.round();
    }
    let micros = (seconds - DAYS_BETWEEN_1900_AND_1970 * SECONDS_PER_DAY) * MICROS_PER_SECOND;
    if micros.is_nan() {
        return i64::MIN;
    }
    // The cast saturates, which is the clamping wanted here.
    micros as i64
}

/// The DuckDB DATE, in days since 1970-01-01, of a TIMESTAMP in microseconds; infinities stay infinite.
pub fn epoch_micros_to_date(micros: i64) -> i32 {
    if micros == i64::MAX {
        return i32::MAX;
    }
    if micros == -i64::MAX {
        return -i32::MAX;
    }
    let negative = i64::from(micros < 0);
    ((micros + negative) / MICROS_PER_DAY - negative) as i32
}

/// The DuckDB TIME, in microseconds since midnight, of a TIMESTAMP in microseconds.
pub fn epoch_micros_to_time(micros: i64) -> Result<i64> {
    if micros == i64::MAX || micros == -i64::MAX {
        return Err(conversion("Can't get TIME of infinite TIMESTAMP"));
    }
    Ok(micros.rem_euclid(MICROS_PER_DAY))
}

/// The serial number of a TIMESTAMP in microseconds since 1970-01-01.
pub fn epoch_micros_to_serial(micros: i64) -> f64 {
    let seconds = micros as f64 / MICROS_PER_SECOND;
    seconds / SECONDS_PER_DAY + DAYS_BETWEEN_1900_AND_1970
}

/// The serial number of a DATE in days since 1970-01-01.
pub fn epoch_days_to_serial(days: i32) -> f64 {
    (i64::from(days) + DAYS_BETWEEN_1900_AND_1970 as i64) as f64
}

/// The fraction of a day a TIME in microseconds since midnight is.
pub fn time_micros_to_serial(micros: i64) -> f64 {
    micros as f64 / MICROS_PER_DAY as f64
}

#[cfg(test)]
mod tests {
    use super::*;

    const DAY: i64 = MICROS_PER_DAY;

    #[test]
    fn serial_numbers_of_timestamps_dates_and_times() {
        assert_eq!(epoch_micros_to_serial(0), 25569.0);
        assert_eq!(
            epoch_micros_to_serial(1_710_505_845_123_000),
            45366.521355590274
        );
        assert_eq!(epoch_micros_to_serial(-DAY / 2), 25568.5);
        assert_eq!(epoch_micros_to_serial(i64::MAX), 106777560.16730064);
        assert_eq!(epoch_micros_to_serial(-i64::MAX), -106726422.16730064);
        assert_eq!(epoch_days_to_serial(0), 25569.0);
        assert_eq!(epoch_days_to_serial(19797), 45366.0);
        assert_eq!(epoch_days_to_serial(-25569), 0.0);
        assert_eq!(epoch_days_to_serial(i32::MAX), 2147509216.0);
        assert_eq!(epoch_days_to_serial(-i32::MAX), -2147458078.0);
        assert_eq!(time_micros_to_serial(0), 0.0);
        assert_eq!(time_micros_to_serial(29_700_000_000), 0.34375);
        assert_eq!(time_micros_to_serial(DAY), 1.0);
    }

    #[test]
    fn serial_numbers_read_back_as_what_was_written() {
        for micros in [0, 1, 999_999, 1_710_505_845_123_456, -86_400_000_001] {
            let read = serial_to_epoch_micros(epoch_micros_to_serial(micros));
            assert!(
                (read - micros).abs() <= 1000,
                "{micros} came back as {read}"
            );
        }
    }

    #[test]
    fn serial_numbers() {
        assert_eq!(serial_to_epoch_micros(25569.0), 0);
        assert_eq!(serial_to_epoch_micros(25570.5), DAY + DAY / 2);
        assert_eq!(serial_to_epoch_micros(0.0), -25569 * DAY);
        assert_eq!(serial_to_epoch_micros(44197.0), 1_609_459_200_000_000);
        assert_eq!(serial_to_epoch_micros(-1.0), -25570 * DAY);
    }

    #[test]
    fn serial_numbers_snap_to_whole_seconds() {
        // 12:34:56 as Excel stores it is not exactly a whole number of seconds.
        let serial = 44197.0 + (12.0 * 3600.0 + 34.0 * 60.0 + 56.0) / 86400.0;
        assert_eq!(
            serial_to_epoch_micros(serial),
            1_609_459_200_000_000 + 45_296_000_000
        );
        assert_eq!(serial_to_epoch_micros(25569.0 + 0.0004 / 86400.0), 0);
        assert_eq!(serial_to_epoch_micros(25569.0 - 0.0004 / 86400.0), 0);
        let quarter_second = serial_to_epoch_micros(25569.0 + 0.25 / 86400.0);
        assert!((249_999..=250_001).contains(&quarter_second));
    }

    #[test]
    fn serial_numbers_out_of_range_are_clamped() {
        assert_eq!(serial_to_epoch_micros(1e300), i64::MAX);
        assert_eq!(serial_to_epoch_micros(f64::INFINITY), i64::MAX);
        assert_eq!(serial_to_epoch_micros(-1e300), i64::MIN);
        assert_eq!(serial_to_epoch_micros(f64::NEG_INFINITY), i64::MIN);
        assert_eq!(serial_to_epoch_micros(f64::NAN), i64::MIN);
    }

    #[test]
    fn dates_round_towards_the_past() {
        assert_eq!(epoch_micros_to_date(0), 0);
        assert_eq!(epoch_micros_to_date(DAY - 1), 0);
        assert_eq!(epoch_micros_to_date(DAY), 1);
        assert_eq!(epoch_micros_to_date(-1), -1);
        assert_eq!(epoch_micros_to_date(-DAY), -1);
        assert_eq!(epoch_micros_to_date(-DAY - 1), -2);
        assert_eq!(epoch_micros_to_date(i64::MAX), i32::MAX);
        assert_eq!(epoch_micros_to_date(-i64::MAX), -i32::MAX);
        assert_eq!(epoch_micros_to_date(i64::MIN), -106_751_992);
    }

    #[test]
    fn times_are_the_time_of_day() {
        assert_eq!(epoch_micros_to_time(0).unwrap(), 0);
        assert_eq!(epoch_micros_to_time(DAY + 5).unwrap(), 5);
        assert_eq!(epoch_micros_to_time(-1).unwrap(), DAY - 1);
        assert_eq!(epoch_micros_to_time(-DAY).unwrap(), 0);
        assert_eq!(epoch_micros_to_time(i64::MIN).unwrap(), 71_945_224_192);
        for infinite in [i64::MAX, -i64::MAX] {
            let error = epoch_micros_to_time(infinite).unwrap_err();
            assert_eq!(error.message, "Can't get TIME of infinite TIMESTAMP");
        }
    }
}
