//! Appends entries to an output that cannot seek.

use std::fmt::Display;
use std::io::{self, Read, Seek, Write};

use flate2::{Compress, Compression, FlushCompress, Status};

use crate::ZipReader;
use crate::format::{
    CENTRAL_HEADER_SIGNATURE, DATA_DESCRIPTOR_SIGNATURE, END_SIGNATURE, FLAG_DATA_DESCRIPTOR,
    FLAG_UTF8, FLAGS_DEFLATE_OPTIONS, FLAGS_ENCRYPTED, LOCAL_HEADER_SIGNATURE, METHOD_DEFLATED,
    METHOD_STORED, ZIP64_MARKER_32, misuse, put_u16, put_u32, unsupported,
};

const OUTPUT_BUFFER_LEN: usize = 64 * 1024;
const INPUT_BUFFER_LEN: usize = 32 * 1024;
/// The backend's level 6 takes more than twice the time of level 5 for output of the same size.
const COMPRESSION_LEVEL: u32 = 5;
const ZIP_VERSION: u16 = 20;
const DOS_DATE_1980_01_01: u16 = (1 << 5) | 1;
const DOS_DIRECTORY_ATTRIBUTE: u32 = 0x10;
const MAX_ENTRIES: u16 = 0xFFFE;

/// Writes an archive front to back, without ever seeking in its output.
///
/// The archive stays within the limits of the original format, because Excel does not read
/// ZIP64: anything that would need it is refused. Once writing to the output has failed the
/// archive is incomplete and the writer must be dropped.
///
/// Small writes into a file are gathered before they reach the compressor, which is slow when
/// fed a few bytes at a time.
#[derive(Debug)]
pub struct ZipWriter<W: Write> {
    output: Output<W>,
    central_directory: Vec<u8>,
    entry_count: u16,
    deflater: Compress,
    pending: Vec<u8>,
    open: Option<OpenFile>,
}

/// The archive bytes on their way out, gathered into large writes.
#[derive(Debug)]
struct Output<W> {
    inner: W,
    buffer: Vec<u8>,
    offset: u64,
}

#[derive(Debug)]
struct OpenFile {
    header: Header,
    data_offset: u64,
    crc: crc32fast::Hasher,
}

#[derive(Debug)]
struct Header {
    name: Vec<u8>,
    flags: u16,
    method: u16,
    dos_time: u16,
    dos_date: u16,
    crc32: u32,
    compressed_size: u32,
    uncompressed_size: u32,
    version_made_by: u16,
    internal_attributes: u16,
    external_attributes: u32,
    offset: u32,
}

impl<W: Write> ZipWriter<W> {
    pub fn new(inner: W) -> Self {
        ZipWriter {
            output: Output {
                inner,
                buffer: Vec::with_capacity(OUTPUT_BUFFER_LEN),
                offset: 0,
            },
            central_directory: Vec::new(),
            entry_count: 0,
            deflater: Compress::new(Compression::new(COMPRESSION_LEVEL), false),
            pending: Vec::with_capacity(INPUT_BUFFER_LEN),
            open: None,
        }
    }

    /// Adds an empty directory entry; `name` must end with a slash.
    pub fn add_directory(&mut self, name: &str) -> io::Result<()> {
        if !name.ends_with('/') {
            return Err(misuse(format!(
                "ZIP directory name '{name}' must end with a slash"
            )));
        }
        let mut header = self.begin_entry(name.as_bytes(), !name.is_ascii())?;
        header.external_attributes = DOS_DIRECTORY_ATTRIBUTE;
        self.output.emit(&header.local())?;
        self.end_entry(&header);
        Ok(())
    }

    /// Starts a deflated entry that receives everything written through [`Write`] until [`Self::end_file`].
    pub fn begin_file(&mut self, name: &str) -> io::Result<()> {
        let mut header = self.begin_entry(name.as_bytes(), !name.is_ascii())?;
        header.flags |= FLAG_DATA_DESCRIPTOR;
        header.method = METHOD_DEFLATED;
        self.output.emit(&header.local())?;
        self.deflater.reset();
        self.open = Some(OpenFile {
            header,
            data_offset: self.output.offset,
            crc: crc32fast::Hasher::new(),
        });
        Ok(())
    }

    pub fn end_file(&mut self) -> io::Result<()> {
        self.compress_pending()?;
        let Some(OpenFile {
            mut header,
            data_offset,
            crc,
        }) = self.open.take()
        else {
            return Err(misuse("cannot end a ZIP entry when none is open"));
        };
        self.output.deflate(&mut self.deflater, &[], true)?;
        header.crc32 = crc.finalize();
        header.compressed_size =
            classic(self.output.offset - data_offset).ok_or_else(|| header.too_large())?;
        let mut descriptor = Vec::with_capacity(16);
        put_u32(&mut descriptor, DATA_DESCRIPTOR_SIGNATURE);
        put_u32(&mut descriptor, header.crc32);
        put_u32(&mut descriptor, header.compressed_size);
        put_u32(&mut descriptor, header.uncompressed_size);
        self.output.emit(&descriptor)?;
        self.end_entry(&header);
        Ok(())
    }

    /// Copies entry `index` of `source` as it is stored, without recompressing it.
    pub fn copy_entry<R: Read + Seek>(
        &mut self,
        source: &mut ZipReader<R>,
        index: usize,
    ) -> io::Result<()> {
        let (entry, record) = source.stored_entry(index)?;
        if record.flags & FLAGS_ENCRYPTED != 0 {
            return Err(unsupported(format!(
                "ZIP entry '{}' is encrypted",
                entry.name
            )));
        }
        let name = record.raw_name.as_deref().unwrap_or(entry.name.as_bytes());
        let utf8 =
            record.flags & FLAG_UTF8 != 0 || (record.raw_name.is_none() && !entry.name.is_ascii());
        let mut header = self.begin_entry(name, utf8)?;
        header.flags |= record.flags & FLAGS_DEFLATE_OPTIONS;
        header.method = record.method;
        header.dos_time = record.dos_time;
        header.dos_date = record.dos_date;
        header.crc32 = entry.crc32;
        header.compressed_size =
            classic(entry.compressed_size).ok_or_else(|| header.too_large())?;
        header.uncompressed_size =
            classic(entry.uncompressed_size).ok_or_else(|| header.too_large())?;
        header.version_made_by = record.version_made_by;
        header.internal_attributes = record.internal_attributes;
        header.external_attributes = record.external_attributes;
        source.locate_raw(index)?;
        self.output.emit(&header.local())?;
        source.drain_raw(|bytes| self.output.emit(bytes))?;
        self.end_entry(&header);
        Ok(())
    }

    /// Writes the central directory and returns the output; fails while an entry is open.
    pub fn finish(mut self) -> io::Result<W> {
        if self.open.is_some() {
            return Err(misuse("cannot finish a ZIP archive while an entry is open"));
        }
        let directory_offset = classic(self.output.offset).ok_or_else(archive_too_large)?;
        let directory_len =
            classic(self.central_directory.len() as u64).ok_or_else(archive_too_large)?;
        self.output.emit(&self.central_directory)?;
        let mut end = Vec::with_capacity(22);
        put_u32(&mut end, END_SIGNATURE);
        put_u16(&mut end, 0);
        put_u16(&mut end, 0);
        put_u16(&mut end, self.entry_count);
        put_u16(&mut end, self.entry_count);
        put_u32(&mut end, directory_len);
        put_u32(&mut end, directory_offset);
        put_u16(&mut end, 0);
        self.output.emit(&end)?;
        self.output.drain()?;
        self.output.inner.flush()?;
        Ok(self.output.inner)
    }

    fn begin_entry(&self, name: &[u8], utf8: bool) -> io::Result<Header> {
        if self.open.is_some() {
            return Err(misuse("cannot start a ZIP entry while another one is open"));
        }
        if u16::try_from(name.len()).is_err() {
            return Err(misuse("ZIP entry name is longer than 65535 bytes"));
        }
        if self.entry_count >= MAX_ENTRIES {
            return Err(needs_zip64(format!(
                "the ZIP archive holds more than {MAX_ENTRIES} entries"
            )));
        }
        Ok(Header {
            name: name.to_vec(),
            flags: if utf8 { FLAG_UTF8 } else { 0 },
            method: METHOD_STORED,
            dos_time: 0,
            dos_date: DOS_DATE_1980_01_01,
            crc32: 0,
            compressed_size: 0,
            uncompressed_size: 0,
            version_made_by: ZIP_VERSION,
            internal_attributes: 0,
            external_attributes: 0,
            offset: classic(self.output.offset).ok_or_else(archive_too_large)?,
        })
    }

    fn end_entry(&mut self, header: &Header) {
        header.append_central(&mut self.central_directory);
        self.entry_count += 1;
    }

    fn compress_pending(&mut self) -> io::Result<()> {
        if let Some(open) = self.open.as_mut() {
            open.crc.update(&self.pending);
        }
        let result = self
            .output
            .deflate(&mut self.deflater, &self.pending, false);
        self.pending.clear();
        result
    }
}

impl<W: Write> Output<W> {
    fn deflate(
        &mut self,
        deflater: &mut Compress,
        mut input: &[u8],
        finish: bool,
    ) -> io::Result<()> {
        let flush = if finish {
            FlushCompress::Finish
        } else {
            FlushCompress::None
        };
        while finish || !input.is_empty() {
            if self.buffer.len() == self.buffer.capacity() {
                self.drain()?;
            }
            let (consumed_before, produced_before) = (deflater.total_in(), deflater.total_out());
            let status = deflater
                .compress_vec(input, &mut self.buffer, flush)
                .map_err(io::Error::other)?;
            let consumed = (deflater.total_in() - consumed_before) as usize;
            let produced = deflater.total_out() - produced_before;
            self.offset += produced;
            input = input.get(consumed..).unwrap_or_default();
            if status == Status::StreamEnd {
                break;
            }
            if consumed == 0 && produced == 0 {
                return Err(io::Error::other(
                    "the deflate stream of a ZIP entry made no progress",
                ));
            }
        }
        Ok(())
    }

    fn emit(&mut self, bytes: &[u8]) -> io::Result<()> {
        if bytes.len() > self.buffer.capacity() - self.buffer.len() {
            self.drain()?;
        }
        if bytes.len() >= self.buffer.capacity() {
            self.inner.write_all(bytes)?;
        } else {
            self.buffer.extend_from_slice(bytes);
        }
        self.offset += bytes.len() as u64;
        Ok(())
    }

    fn drain(&mut self) -> io::Result<()> {
        let result = self.inner.write_all(&self.buffer);
        self.buffer.clear();
        result
    }
}

/// Writes into the entry started by [`ZipWriter::begin_file`].
impl<W: Write> Write for ZipWriter<W> {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        let Some(open) = self.open.as_mut() else {
            return Err(misuse(
                "cannot write to a ZIP archive without an open entry",
            ));
        };
        let header = &mut open.header;
        header.uncompressed_size =
            classic(u64::from(header.uncompressed_size).saturating_add(buffer.len() as u64))
                .ok_or_else(|| header.too_large())?;
        if buffer.len() > self.pending.capacity() - self.pending.len() {
            self.compress_pending()?;
        }
        if buffer.len() < self.pending.capacity() {
            self.pending.extend_from_slice(buffer);
        } else if let Some(open) = self.open.as_mut() {
            open.crc.update(buffer);
            self.output.deflate(&mut self.deflater, buffer, false)?;
        }
        Ok(buffer.len())
    }

    /// Pushes the archive bytes produced so far; input the compressor still holds follows later.
    fn flush(&mut self) -> io::Result<()> {
        self.compress_pending()?;
        self.output.drain()?;
        self.output.inner.flush()
    }
}

impl Header {
    fn too_large(&self) -> io::Error {
        needs_zip64(format!(
            "ZIP entry '{}' is larger than 4 GiB",
            String::from_utf8_lossy(&self.name)
        ))
    }

    fn local(&self) -> Vec<u8> {
        let mut out = Vec::with_capacity(30 + self.name.len());
        put_u32(&mut out, LOCAL_HEADER_SIGNATURE);
        self.append_common(&mut out);
        put_u16(&mut out, 0);
        out.extend_from_slice(&self.name);
        out
    }

    fn append_central(&self, out: &mut Vec<u8>) {
        put_u32(out, CENTRAL_HEADER_SIGNATURE);
        put_u16(out, self.version_made_by);
        self.append_common(out);
        put_u16(out, 0);
        put_u16(out, 0);
        put_u16(out, 0);
        put_u16(out, self.internal_attributes);
        put_u32(out, self.external_attributes);
        put_u32(out, self.offset);
        out.extend_from_slice(&self.name);
    }

    fn append_common(&self, out: &mut Vec<u8>) {
        put_u16(out, ZIP_VERSION);
        put_u16(out, self.flags);
        put_u16(out, self.method);
        put_u16(out, self.dos_time);
        put_u16(out, self.dos_date);
        put_u32(out, self.crc32);
        put_u32(out, self.compressed_size);
        put_u32(out, self.uncompressed_size);
        put_u16(out, self.name.len() as u16);
    }
}

fn classic(value: u64) -> Option<u32> {
    u32::try_from(value)
        .ok()
        .filter(|value| *value != ZIP64_MARKER_32)
}

fn archive_too_large() -> io::Error {
    needs_zip64("the ZIP archive is larger than 4 GiB")
}

fn needs_zip64(problem: impl Display) -> io::Error {
    unsupported(format!(
        "{problem}, which needs ZIP64 and that is not supported for writing"
    ))
}
