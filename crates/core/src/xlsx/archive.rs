//! A workbook as the ZIP archive it is, read through DuckDB's file system.

use std::io::{self, Read};

use duckdb_neo::{Result, file::FileSystem};
use excel_zip::ZipReader;

use crate::{
    error::from_io,
    file::FileReader,
    xlsx::xml::{XmlHandler, XmlReader},
};

/// The parts of a workbook, of which one at a time can be open for reading.
#[derive(Debug)]
pub struct Archive {
    zip: ZipReader<FileReader>,
}

impl Archive {
    pub fn open(file_system: &FileSystem<'_>, path: &str) -> Result<Archive> {
        let file = FileReader::open(file_system, path)?;
        let zip = ZipReader::new(file).map_err(|error| from_io(&error))?;
        Ok(Archive { zip })
    }

    /// Opens the part called `name` and returns whether the archive has it.
    pub fn try_open_entry(&mut self, name: &str) -> Result<bool> {
        match self.zip.find(name) {
            Some(index) => {
                self.zip
                    .open_entry(index)
                    .map_err(|error| from_io(&error))?;
                Ok(true)
            }
            None => Ok(false),
        }
    }

    /// Closes the open part; a part that was read to its end must match its checksum.
    pub fn close_entry(&mut self) -> Result<()> {
        self.zip.close_entry().map_err(|error| from_io(&error))
    }

    /// The uncompressed size of the open part.
    pub fn entry_len(&self) -> u64 {
        self.zip.entry_len()
    }

    /// Runs `handler` over the part called `name` and returns whether the archive has that part.
    pub fn parse_entry<H: XmlHandler>(&mut self, name: &str, handler: &mut H) -> Result<bool> {
        self.parse_entry_with_limit(name, None, handler)
    }

    /// As [`Self::parse_entry`], with a limit on the bytes of a single tag, comment or run of text.
    pub fn parse_entry_with_limit<H: XmlHandler>(
        &mut self,
        name: &str,
        limit: Option<usize>,
        handler: &mut H,
    ) -> Result<bool> {
        if !self.try_open_entry(name)? {
            return Ok(false);
        }
        XmlReader::new(&mut *self, limit).parse_all(handler)?;
        self.close_entry()?;
        Ok(true)
    }
}

impl Read for Archive {
    fn read(&mut self, buffer: &mut [u8]) -> io::Result<usize> {
        self.zip.read(buffer)
    }
}
