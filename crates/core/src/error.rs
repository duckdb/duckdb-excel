//! Constructors for the errors the extension raises, each with the error class DuckDB reports it under, and the
//! conversions between DuckDB errors and `std::io` errors.

use std::{fmt, io};

use duckdb_neo::{Error, error::DuckDBError};

fn error(code: DuckDBError, message: impl Into<String>) -> Error {
    Error {
        code,
        message: message.into(),
    }
}

pub fn binder(message: impl Into<String>) -> Error {
    error(DuckDBError::QueryBinder, message)
}

pub fn invalid_input(message: impl Into<String>) -> Error {
    error(DuckDBError::InputInvalid, message)
}

pub fn io(message: impl Into<String>) -> Error {
    error(DuckDBError::IoGeneral, message)
}

pub fn conversion(message: impl Into<String>) -> Error {
    error(DuckDBError::TypeConversion, message)
}

pub fn internal(message: impl Into<String>) -> Error {
    error(DuckDBError::RuntimeInternal, message)
}

/// The message of an error DuckDB reported, without the error class DuckDB puts in front of it.
pub fn raw_message(error: &Error) -> &str {
    let message = error.message.as_str();
    match message.split_once(" Error: ") {
        Some((class, raw))
            if !class.is_empty()
                && class
                    .bytes()
                    .all(|byte| byte.is_ascii_alphabetic() || byte == b' ') =>
        {
            raw
        }
        _ => message,
    }
}

/// Makes an error DuckDB reported fit to be raised again: DuckDB puts the error class in front once more.
pub fn reraise(error: Error) -> Error {
    Error {
        code: error.code,
        message: raw_message(&error).to_owned(),
    }
}

/// A DuckDB error on its way through an interface that speaks `std::io::Error`.
#[derive(Debug)]
struct CarriedError(Error);

impl fmt::Display for CarriedError {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(&self.0.message)
    }
}

impl std::error::Error for CarriedError {}

/// Wraps a DuckDB error so that [`from_io`] gets it back unchanged.
pub fn to_io(error: Error) -> io::Error {
    io::Error::other(CarriedError(error))
}

/// The DuckDB error an I/O error carries, or an IO error with the message of the I/O error.
pub fn from_io(error: &io::Error) -> Error {
    match error
        .get_ref()
        .and_then(|inner| inner.downcast_ref::<CarriedError>())
    {
        Some(CarriedError(carried)) => carried.clone(),
        None => io(error.to_string()),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn raw_message_drops_the_error_class() {
        let reported = io("IO Error: Cannot open file \"a.xlsx\": No such file or directory");
        assert_eq!(
            raw_message(&reported),
            "Cannot open file \"a.xlsx\": No such file or directory"
        );
        let reported = invalid_input("Invalid Input Error: nope");
        assert_eq!(raw_message(&reported), "nope");
    }

    #[test]
    fn raw_message_keeps_messages_without_a_class() {
        for message in [
            "plain",
            " Error: no class",
            "Sheet \"a Error: b\" not found",
            "x: y Error: z",
        ] {
            assert_eq!(raw_message(&binder(message)), message);
        }
    }

    #[test]
    fn io_round_trip_keeps_code_and_message() {
        let carried = from_io(&to_io(binder("the message")));
        assert_eq!(carried.code, DuckDBError::QueryBinder);
        assert_eq!(carried.message, "the message");
    }

    #[test]
    fn foreign_io_errors_become_io_errors() {
        let foreign = io::Error::new(io::ErrorKind::InvalidData, "broken archive");
        let converted = from_io(&foreign);
        assert_eq!(converted.code, DuckDBError::IoGeneral);
        assert_eq!(converted.message, "broken archive");
    }
}
