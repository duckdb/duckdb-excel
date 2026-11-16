//! A small ZIP archive reader and writer, covering what XLSX packages need.
//!
//! The reader walks the central directory and streams one entry at a time; the writer only ever
//! appends to its output, so it works on files that cannot seek.

use std::io::{self, Read, Seek, Write};

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

/// Reads an archive, one open entry at a time.
#[derive(Debug)]
pub struct ZipReader<R> {
    inner: R,
}

impl<R: Read + Seek> ZipReader<R> {
    /// Reads the central directory of the archive in `inner`.
    pub fn new(inner: R) -> io::Result<Self> {
        Ok(ZipReader { inner })
    }

    /// The entries in central directory order.
    pub fn entries(&self) -> &[ZipEntry] {
        unimplemented!()
    }

    /// The index of the last entry called `name`: producers that repeat a name mean the last one.
    pub fn find(&self, name: &str) -> Option<usize> {
        let _ = name;
        unimplemented!()
    }

    /// Opens the entry at `index` for reading through [`Read`]; only one entry can be open.
    pub fn open_entry(&mut self, index: usize) -> io::Result<()> {
        let _ = index;
        unimplemented!()
    }

    pub fn is_entry_open(&self) -> bool {
        unimplemented!()
    }

    /// The uncompressed size of the open entry, as the central directory states it.
    pub fn entry_len(&self) -> u64 {
        unimplemented!()
    }

    /// The number of uncompressed bytes read from the open entry.
    pub fn entry_pos(&self) -> u64 {
        unimplemented!()
    }

    /// Whether the open entry has been read up to its stated size.
    pub fn is_done(&self) -> bool {
        unimplemented!()
    }

    /// Closes the open entry; its checksum is only verified when it was read to the end.
    pub fn close_entry(&mut self) -> io::Result<()> {
        unimplemented!()
    }

    pub fn into_inner(self) -> R {
        self.inner
    }
}

/// Reads the uncompressed bytes of the open entry.
impl<R: Read + Seek> Read for ZipReader<R> {
    fn read(&mut self, buffer: &mut [u8]) -> io::Result<usize> {
        let _ = buffer;
        unimplemented!()
    }
}

/// Writes an archive front to back, without ever seeking in its output.
#[derive(Debug)]
pub struct ZipWriter<W: Write> {
    inner: W,
}

impl<W: Write> ZipWriter<W> {
    pub fn new(inner: W) -> Self {
        ZipWriter { inner }
    }

    /// Adds an empty directory entry; `name` must end with a slash.
    pub fn add_directory(&mut self, name: &str) -> io::Result<()> {
        let _ = name;
        unimplemented!()
    }

    /// Starts a deflated entry that receives everything written through [`Write`] until [`Self::end_file`].
    pub fn begin_file(&mut self, name: &str) -> io::Result<()> {
        let _ = name;
        unimplemented!()
    }

    pub fn end_file(&mut self) -> io::Result<()> {
        unimplemented!()
    }

    /// Copies entry `index` of `source` as it is stored, without recompressing it.
    pub fn copy_entry<R: Read + Seek>(&mut self, source: &mut ZipReader<R>, index: usize) -> io::Result<()> {
        let _ = (source, index);
        unimplemented!()
    }

    /// Writes the central directory and returns the output.
    pub fn finish(self) -> io::Result<W> {
        Ok(self.inner)
    }
}

/// Writes into the entry started by [`ZipWriter::begin_file`].
impl<W: Write> Write for ZipWriter<W> {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        let _ = buffer;
        unimplemented!()
    }

    fn flush(&mut self) -> io::Result<()> {
        unimplemented!()
    }
}
