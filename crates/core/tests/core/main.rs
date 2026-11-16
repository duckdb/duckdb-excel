//! The integration tests of the crate, which all run against an in-process DuckDB.

mod support;

mod casts;
mod copy_from;
mod copy_to;
mod copy_to_append;
mod copy_to_bytes;
mod file;
mod fixtures;
mod keywords;
mod malformed;
mod read_options;
mod replacement_scan;
mod scan;
mod sniffing;
mod text;
