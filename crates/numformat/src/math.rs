//! Floating point helpers of the formatter: decimal scaling, rounding and digit generation.
//!
//! [`double_to_string`] is not a shortest-round-trip printer: it scales the value into `[1, 10)`,
//! extracts at most 15 significant digits with small corrective offsets and pads with zeros. It
//! writes no sign, uses `.` as separator and gives the exponent three digits (`d.ddE+ddd`). Its
//! digits are what users see, so every step keeps its exact order of operations, and `log10` and
//! `pow` come from the platform's C library because their last bit can show in the result.
//!
//! Negative values are not negated before the digits are taken: each "digit" is then a character
//! below `0`. Text formats print negative numbers that way, so the function keeps code units wide
//! until the end, where such a character could otherwise pass for `0`, `.` or the terminating NUL.
//!
//! The integer conversions give what x86-64 gives for NaN and out-of-range values, which reach
//! the output for NaN and infinite inputs: the smallest integer for the signed types; for `u64`
//! 2^63 for NaN, wrap-around for negative values and 0 from 2^64 on.

use std::hint::black_box;

const POWERS_OF_TEN: [f64; 16] = [
    1e1, 1e2, 1e3, 1e4, 1e5, 1e6, 1e7, 1e8, 1e9, 1e10, 1e11, 1e12, 1e13, 1e14, 1e15, 1e16,
];
const NEGATIVE_POWERS_OF_TEN: [f64; 16] = [
    1e-1, 1e-2, 1e-3, 1e-4, 1e-5, 1e-6, 1e-7, 1e-8, 1e-9, 1e-10, 1e-11, 1e-12, 1e-13, 1e-14, 1e-15,
    1e-16,
];
const CORRECTION: [f64; 16] = [
    0.0, 9e-1, 9e-2, 9e-3, 9e-4, 9e-5, 9e-6, 9e-7, 9e-8, 9e-9, 9e-10, 9e-11, 9e-12, 9e-13, 9e-14,
    9e-15,
];
const ROUNDING: [f64; 16] = [
    5.0e+0, 0.5e+0, 0.5e-1, 0.5e-2, 0.5e-3, 0.5e-4, 0.5e-5, 0.5e-6, 0.5e-7, 0.5e-8, 0.5e-9,
    0.5e-10, 0.5e-11, 0.5e-12, 0.5e-13, 0.5e-14,
];

pub(crate) fn pow(base: f64, exponent: f64) -> f64 {
    // The opaque base keeps the compiler from substituting another function for the call.
    black_box(base).powf(exponent)
}

pub(crate) fn pow10(exponent: i32) -> f64 {
    let table = if exponent < 0 {
        &NEGATIVE_POWERS_OF_TEN
    } else {
        &POWERS_OF_TEN
    };
    match exponent.unsigned_abs() {
        0 => 1.0,
        magnitude @ 1..=16 => table[magnitude as usize - 1],
        _ => pow(10.0, f64::from(exponent)),
    }
}

pub(crate) fn to_i32(value: f64) -> i32 {
    if value > -2147483649.0 && value < 2147483648.0 {
        value as i32
    } else {
        i32::MIN
    }
}

pub(crate) fn to_i64(value: f64) -> i64 {
    if (-9223372036854775808.0..9223372036854775808.0).contains(&value) {
        value as i64
    } else {
        i64::MIN
    }
}

pub(crate) fn to_u64(value: f64) -> u64 {
    if value >= 9223372036854775808.0 {
        (to_i64(value - 9223372036854775808.0) as u64) ^ (1 << 63)
    } else {
        to_i64(value) as u64
    }
}

pub(crate) fn decimal_exponent(value: f64) -> i32 {
    to_i32(value.log10().floor())
}

pub(crate) fn round_half_up_corrected(value: f64, decimals: i32) -> f64 {
    if value == 0.0 {
        return value;
    }
    let mut value = value;
    let mut factor = 0.0;
    if decimals != 0 {
        if !(-20..=20).contains(&decimals) || value > f64::MAX / 1e20 {
            return value;
        }
        factor = pow10(decimals);
        value *= factor;
    }
    let exponent = if value > 0.0 {
        decimal_exponent(value)
    } else {
        0
    };
    let index = match 15i32.wrapping_sub(exponent) {
        index if index > 15 => 15,
        index if index <= 1 => 0,
        index => index as usize,
    };
    value = (value + 0.5 + CORRECTION[index]).floor();
    if decimals != 0 {
        value /= factor;
    }
    value
}

pub(crate) fn integer_digits(value: f64) -> i64 {
    if !(1e-7..=1e7).contains(&value) {
        return to_i64(value.log10().floor()).wrapping_add(1);
    }
    let mut value = value;
    let mut digits = 1;
    while value < 1.0 {
        value *= 10.0;
        digits -= 1;
    }
    while value >= 10.0 {
        value /= 10.0;
        digits += 1;
    }
    digits
}

#[derive(Clone, Copy, PartialEq, Eq)]
pub(crate) enum Notation {
    Exponential,
    Fixed,
}

pub(crate) fn double_to_string(
    value: f64,
    notation: Notation,
    decimals: i32,
    erase_trailing_zeros: bool,
) -> Vec<u8> {
    const ZERO: i32 = b'0' as i32;
    const NINE: i32 = b'9' as i32;
    const POINT: i32 = b'.' as i32;

    if value.is_nan() {
        return b"NaN".to_vec();
    }
    if value.is_infinite() {
        return b"INF".to_vec();
    }
    let fixed = notation == Notation::Fixed;
    let mut value = value;
    let mut exponent = 0;
    if value > 0.0 {
        exponent = decimal_exponent(value);
        value /= pow10(exponent);
    }

    let mut digits = decimals.wrapping_add(1);
    if fixed {
        digits = digits.wrapping_add(exponent);
    }
    if digits >= 0 {
        value += ROUNDING[digits.min(15) as usize];
        if value >= 10.0 {
            value = 1.0;
            exponent += 1;
            if fixed {
                digits += 1;
            }
        }
    }

    let mut buffer: Vec<i32> = Vec::with_capacity(32);
    let mut has_point = false;
    let mut point_position;
    if fixed {
        if exponent < 0 {
            buffer.push(ZERO);
            if decimals > 0 {
                buffer.push(POINT);
                has_point = true;
            }
            let leading_zeros = if digits <= 0 { decimals } else { -exponent - 1 };
            buffer.extend(std::iter::repeat_n(ZERO, leading_zeros.max(0) as usize));
            point_position = 0;
        } else {
            point_position = exponent + 1;
        }
    } else {
        point_position = 1;
    }

    if digits > 0 {
        let mut index = 0usize;
        loop {
            if index < 15 {
                let digit = if digits - 1 == 0 && index > 0 && index < 14 {
                    to_i32((value + CORRECTION[15 - index]).floor())
                } else {
                    to_i32(value + 1e-15)
                };
                if digit >= 10 {
                    if buffer.is_empty() {
                        buffer.push(ZERO + 1);
                        if !fixed {
                            buffer.push(POINT);
                            exponent += 1;
                            has_point = true;
                        }
                        buffer.push(ZERO);
                    } else {
                        for position in (0..buffer.len()).rev() {
                            let unit = buffer[position];
                            if unit == POINT {
                                continue;
                            }
                            if unit != NINE {
                                buffer[position] = unit.wrapping_add(1);
                                break;
                            }
                            buffer[position] = ZERO;
                            if position == 0 {
                                if fixed {
                                    buffer.insert(0, ZERO + 1);
                                } else {
                                    buffer[0] = ZERO + 1;
                                    exponent += 1;
                                }
                            }
                        }
                        buffer.push(ZERO);
                    }
                    value = 0.0;
                } else {
                    buffer.push(digit.wrapping_add(ZERO));
                    value = (value - f64::from(digit)) * 10.0;
                }
            } else {
                buffer.push(ZERO);
            }
            digits -= 1;
            if digits == 0 {
                break;
            }
            if point_position != 0 {
                point_position -= 1;
                if point_position == 0 {
                    buffer.push(POINT);
                    has_point = true;
                }
            }
            index += 1;
        }
    }

    if !has_point && fixed {
        while point_position > 1 {
            point_position -= 1;
            buffer.push(ZERO);
        }
    }

    if erase_trailing_zeros && has_point && !buffer.is_empty() {
        while buffer.last() == Some(&ZERO) {
            buffer.pop();
        }
        if buffer.last() == Some(&POINT) {
            buffer.pop();
        }
    }

    if !fixed {
        if buffer.is_empty() {
            buffer.push(ZERO + 1);
        }
        buffer.push(i32::from(b'E'));
        if exponent < 0 {
            exponent = -exponent;
            buffer.push(i32::from(b'-'));
        } else {
            buffer.push(i32::from(b'+'));
        }
        buffer.push(ZERO + exponent / 100);
        exponent %= 100;
        buffer.push(ZERO + exponent / 10);
        buffer.push(ZERO + exponent % 10);
    }
    // A "digit" of a negative value that comes out as NUL ends the result.
    buffer
        .iter()
        .take_while(|&&unit| unit != 0)
        .map(|&unit| unit as u8)
        .collect()
}

#[cfg(test)]
mod tests {
    use super::*;

    fn fixed(value: f64, decimals: i32) -> String {
        String::from_utf8(double_to_string(value, Notation::Fixed, decimals, false)).unwrap()
    }

    fn exponential(value: f64, decimals: i32) -> String {
        String::from_utf8(double_to_string(
            value,
            Notation::Exponential,
            decimals,
            false,
        ))
        .unwrap()
    }

    #[test]
    fn pow10_uses_exact_constants_near_one() {
        assert_eq!(pow10(0), 1.0);
        assert_eq!(pow10(16), 1e16);
        assert_eq!(pow10(-16), 1e-16);
        assert_eq!(pow10(-324), 0.0);
    }

    #[test]
    fn integer_conversions_of_unrepresentable_values() {
        assert_eq!(to_i32(f64::NAN), i32::MIN);
        assert_eq!(to_i32(-2147483648.9), i32::MIN);
        assert_eq!(to_i32(2147483648.0), i32::MIN);
        assert_eq!(to_i64(f64::INFINITY), i64::MIN);
        assert_eq!(to_u64(f64::NAN), 1 << 63);
        assert_eq!(to_u64(-2.5), u64::MAX - 1);
        assert_eq!(to_u64(1.8446744073709552e19), 0);
        assert_eq!(to_u64(9223372036854775808.0), 1 << 63);
    }

    #[test]
    fn round_half_up_corrected_compensates_representation_errors() {
        assert_eq!(round_half_up_corrected(2.5, 0), 3.0);
        assert_eq!(round_half_up_corrected(0.125, 2), 0.13);
        assert_eq!(round_half_up_corrected(1.005, 2), 1.01);
        assert_eq!(round_half_up_corrected(1.25, 21), 1.25);
    }

    #[test]
    fn integer_digits_counts_from_the_decimal_point() {
        assert_eq!(integer_digits(1234.5), 4);
        assert_eq!(integer_digits(1.0), 1);
        assert_eq!(integer_digits(0.001), -2);
        assert_eq!(integer_digits(1e20), 21);
        assert_eq!(integer_digits(f64::INFINITY), i64::MIN + 1);
    }

    #[test]
    fn double_to_string_fixed() {
        assert_eq!(fixed(1234.567, 2), "1234.57");
        assert_eq!(fixed(0.0, 2), "0.00");
        assert_eq!(fixed(0.005, 2), "0.01");
        assert_eq!(fixed(999.995, 2), "1000.00");
        assert_eq!(fixed(0.7, 0), "1");
        assert_eq!(fixed(1e20, 1), "100000000000000000000.0");
        assert_eq!(fixed(0.12345678901234568, 20), "0.12345678901234600000");
    }

    #[test]
    fn double_to_string_exponential() {
        assert_eq!(exponential(12345.678, 2), "1.23E+004");
        assert_eq!(exponential(0.000999, 1), "1.0E-003");
        assert_eq!(exponential(0.0, 0), "0E+000");
        assert_eq!(exponential(5e-324, 2), "1.00E-323");
    }

    #[test]
    fn double_to_string_erases_trailing_zeros() {
        let erased = |value| double_to_string(value, Notation::Fixed, 2, true);
        assert_eq!(erased(1.5), b"1.5");
        assert_eq!(erased(2.0), b"2");
        assert_eq!(erased(100.0), b"100");
    }

    #[test]
    fn double_to_string_of_non_finite_values() {
        assert_eq!(fixed(f64::NAN, 2), "NaN");
        assert_eq!(fixed(f64::INFINITY, 2), "INF");
        assert_eq!(exponential(f64::NEG_INFINITY, 2), "INF");
    }

    #[test]
    fn double_to_string_of_negative_values_is_not_a_number() {
        assert_eq!(double_to_string(-5.0, Notation::Fixed, 2, true), b",.'&");
        assert_eq!(double_to_string(-1234.5, Notation::Fixed, 2, true), b"^.,&");
        assert_eq!(
            double_to_string(-1e20, Notation::Exponential, 2, false),
            b"0.00E+000"
        );
        assert_eq!(double_to_string(-48.5, Notation::Fixed, 2, true), b"");
    }
}
