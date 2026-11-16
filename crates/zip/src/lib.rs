//! A small ZIP archive reader and writer, covering what XLSX packages need.
//!
//! The reader walks the central directory and streams one entry at a time; the writer only ever
//! appends to its output, so it works on files that cannot seek.
//!
//! Errors are [`std::io::Error`]s: `InvalidData` for malformed archives, `Unsupported` for
//! features that are not handled and `InvalidInput` for calls made in the wrong state.

mod directory;
mod format;
mod reader;
mod writer;

pub use reader::ZipReader;
pub use writer::ZipWriter;

/// One entry of an archive's central directory.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct ZipEntry {
    pub name: String,
    pub compressed_size: u64,
    pub uncompressed_size: u64,
    pub crc32: u32,
}

impl ZipEntry {
    /// Directories are entries whose name ends with a slash.
    pub fn is_directory(&self) -> bool {
        self.name.ends_with('/')
    }
}
