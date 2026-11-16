//! Byte strings with checked positions.
//!
//! Format codes travel through the formatter as opaque bytes, so multi-byte characters are never
//! interpreted. A position outside a string is a [`FormatError`] instead of a panic.

use crate::FormatError;

pub(crate) type Text = Vec<u8>;

pub(crate) fn at(text: &[u8], position: usize) -> Result<u8, FormatError> {
    text.get(position).copied().ok_or(FormatError)
}

pub(crate) fn insert(text: &mut Text, position: usize, piece: &[u8]) -> Result<(), FormatError> {
    if position > text.len() {
        return Err(FormatError);
    }
    text.splice(position..position, piece.iter().copied());
    Ok(())
}

pub(crate) fn insert_byte(text: &mut Text, position: usize, byte: u8) -> Result<(), FormatError> {
    insert(text, position, &[byte])
}

/// Removes up to `count` bytes starting at `position`.
pub(crate) fn erase(text: &mut Text, position: usize, count: usize) -> Result<(), FormatError> {
    if position > text.len() {
        return Err(FormatError);
    }
    let end = position.saturating_add(count).min(text.len());
    text.drain(position..end);
    Ok(())
}

/// The part of `text` starting at `position`.
pub(crate) fn tail(text: &[u8], position: usize) -> Result<&[u8], FormatError> {
    text.get(position..).ok_or(FormatError)
}

/// Removes leading `byte`s but always leaves the last byte of the text in place.
pub(crate) fn erase_leading(text: &mut Text, byte: u8) {
    let Some(last) = text.len().checked_sub(1) else {
        return;
    };
    let count = text
        .iter()
        .position(|&other| other != byte)
        .unwrap_or(last)
        .min(last);
    text.drain(..count);
}

pub(crate) fn erase_all(text: &mut Text, byte: u8) {
    text.retain(|&other| other != byte);
}

/// Whether `text` starts with `prefix`, ignoring ASCII case in `text`; `prefix` must be upper case.
pub(crate) fn starts_with_upper(text: &[u8], prefix: &[u8]) -> bool {
    text.len() >= prefix.len()
        && text
            .iter()
            .zip(prefix)
            .all(|(byte, expected)| byte.to_ascii_uppercase() == *expected)
}

/// A negative count stands for "to the end of the text".
pub(crate) fn count_from(difference: i32) -> usize {
    usize::try_from(difference).unwrap_or(usize::MAX)
}
