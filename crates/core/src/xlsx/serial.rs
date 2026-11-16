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

#[cfg(test)]
mod tests {
    use super::*;

    const DAY: i64 = MICROS_PER_DAY;

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
