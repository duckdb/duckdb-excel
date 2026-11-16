//! One parser for each kind of part of a workbook; all of them are handlers of [`crate::xlsx::xml::XmlReader`].

pub mod content_types;
pub mod header_sniffer;
pub mod range_sniffer;
pub mod relationships;
pub mod row_collector;
pub mod shared_strings;
pub mod styles;
pub mod styles_append;
pub mod workbook;
pub mod workbook_append;
pub mod worksheet;
