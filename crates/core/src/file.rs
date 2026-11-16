//! DuckDB files as `std::io` streams, so that every DuckDB file system can hold a workbook.
//!
//! Neither adapter buffers: `excel_zip` already reads and writes in large blocks.

use std::io::{self, Read, Seek, SeekFrom, Write};

use duckdb_neo::{
    Error, Result,
    file::{File, FileBuilder, FileSystem},
};

use crate::error::{self, reraise};

fn to_io(error: Error) -> io::Error {
    error::to_io(reraise(error))
}

/// The first file `pattern` matches, in the order the file system lists its matches, which need not be sorted.
pub fn first_match(file_system: &FileSystem<'_>, pattern: &str) -> Result<String> {
    let matches = file_system.glob(pattern).map_err(reraise)?;
    matches.into_iter().next().ok_or_else(|| {
        error::io(format!(
            "Cannot open file \"{pattern}\": No such file or directory"
        ))
    })
}

/// A DuckDB file opened for reading.
#[derive(Debug)]
pub struct FileReader {
    file: File,
    position: u64,
}

impl FileReader {
    pub fn open(file_system: &FileSystem<'_>, path: &str) -> Result<FileReader> {
        let file = FileBuilder::new(file_system, path)?
            .read()
            .open()
            .map_err(reraise)?;
        Ok(FileReader { file, position: 0 })
    }
}

impl Read for FileReader {
    fn read(&mut self, buffer: &mut [u8]) -> io::Result<usize> {
        let count = self.file.read_into(buffer).map_err(to_io)?;
        self.position += count as u64;
        Ok(count)
    }
}

impl Seek for FileReader {
    fn seek(&mut self, target: SeekFrom) -> io::Result<u64> {
        let position = match target {
            SeekFrom::Start(position) => Some(position),
            SeekFrom::End(offset) => self.file.size().map_err(to_io)?.checked_add_signed(offset),
            SeekFrom::Current(offset) => self.position.checked_add_signed(offset),
        };
        let position = position.and_then(|position| {
            usize::try_from(position)
                .ok()
                .map(|index| (position, index))
        });
        let Some((position, index)) = position else {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "cannot seek to a position outside the file",
            ));
        };
        self.file.seek(index).map_err(to_io)?;
        self.position = position;
        Ok(position)
    }
}

/// A DuckDB file opened for writing, created when it does not exist.
///
/// An existing file is written over from its start and keeps what lies beyond the bytes written.
#[derive(Debug)]
pub struct FileWriter {
    file: File,
}

impl FileWriter {
    pub fn create(file_system: &FileSystem<'_>, path: &str) -> Result<FileWriter> {
        let file = FileBuilder::new(file_system, path)?
            .write()
            .create()
            .open()
            .map_err(reraise)?;
        Ok(FileWriter { file })
    }

    /// Closes the file, reporting what a silent drop would hide.
    pub fn close(self) -> Result<()> {
        self.file.close().map_err(reraise)
    }
}

impl Write for FileWriter {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        self.file.write(buffer).map_err(to_io)
    }

    fn flush(&mut self) -> io::Result<()> {
        Ok(())
    }
}
