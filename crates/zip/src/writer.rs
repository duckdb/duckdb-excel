//! Appends entries to an output that cannot seek.

use std::fmt::Display;
use std::io::{self, Read, Seek, Write};

use flate2::{Compress, Compression, FlushCompress, Status};

use crate::ZipReader;
use crate::format::{
    CENTRAL_HEADER_SIGNATURE, DATA_DESCRIPTOR_SIGNATURE, END_SIGNATURE, FLAG_DATA_DESCRIPTOR,
    FLAG_UTF8, FLAGS_DEFLATE_OPTIONS, FLAGS_ENCRYPTED, LOCAL_HEADER_SIGNATURE, METHOD_DEFLATED,
    ZIP64_MARKER_32, misuse, put_u16, put_u32, unsupported,
};

const OUTPUT_BUFFER_LEN: usize = 64 * 1024;
const INPUT_BUFFER_LEN: usize = 32 * 1024;
const COMPRESSION_LEVEL: u32 = 6;
const ZIP_VERSION: u16 = 20;
const DOS_DATE_1980_01_01: u16 = (1 << 5) | 1;
const MAX_ENTRIES: u16 = 0xFFFE;

/// Writes an archive front to back, without ever seeking in its output.
///
/// The archive stays within the limits of the original format, because Excel does not read
/// ZIP64: anything that would need it is refused. A call that is refused before it writes
/// anything leaves the writer usable. Once a call has failed half-way the archive is broken
/// and every later call fails.
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
    crc: crc32fast::Hasher,
    open: Option<OpenFile>,
    broken: bool,
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
    version_needed: u16,
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
            crc: crc32fast::Hasher::new(),
            open: None,
            broken: false,
        }
    }

    /// Starts a deflated entry that receives everything written through [`Write`] until [`Self::end_file`].
    pub fn begin_file(&mut self, name: &str) -> io::Result<()> {
        let offset = self.next_entry_offset(name.as_bytes())?;
        let utf8 = if name.is_ascii() { 0 } else { FLAG_UTF8 };
        let header = Header {
            name: name.as_bytes().to_vec(),
            flags: FLAG_DATA_DESCRIPTOR | utf8,
            method: METHOD_DEFLATED,
            dos_time: 0,
            dos_date: DOS_DATE_1980_01_01,
            crc32: 0,
            compressed_size: 0,
            uncompressed_size: 0,
            version_made_by: ZIP_VERSION,
            version_needed: ZIP_VERSION,
            internal_attributes: 0,
            external_attributes: 0,
            offset,
        };
        self.commit(|writer| {
            writer.output.emit(&header.local())?;
            writer.deflater.reset();
            let data_offset = writer.output.offset;
            writer.open = Some(OpenFile {
                header,
                data_offset,
            });
            Ok(())
        })
    }

    pub fn end_file(&mut self) -> io::Result<()> {
        self.ensure_intact()?;
        let Some(OpenFile {
            mut header,
            data_offset,
        }) = self.open.take()
        else {
            return Err(misuse("cannot end a ZIP entry when none is open"));
        };
        self.commit(|writer| {
            writer.compress_pending()?;
            writer.output.deflate(&mut writer.deflater, &[], true)?;
            header.crc32 = std::mem::take(&mut writer.crc).finalize();
            header.compressed_size =
                classic(writer.output.offset - data_offset).ok_or_else(|| header.too_large())?;
            let mut descriptor = Vec::with_capacity(16);
            put_u32(&mut descriptor, DATA_DESCRIPTOR_SIGNATURE);
            put_u32(&mut descriptor, header.crc32);
            put_u32(&mut descriptor, header.compressed_size);
            put_u32(&mut descriptor, header.uncompressed_size);
            writer.output.emit(&descriptor)?;
            writer.end_entry(&header);
            Ok(())
        })
    }

    /// Copies entry `index` of `source` as it is stored, without recompressing it.
    pub fn copy_entry<R: Read + Seek>(
        &mut self,
        source: &mut ZipReader<R>,
        index: usize,
    ) -> io::Result<()> {
        self.ensure_intact()?;
        let (entry, record) = source.stored_entry(index)?;
        if record.flags & FLAGS_ENCRYPTED != 0 {
            return Err(unsupported(format!(
                "ZIP entry '{}' is encrypted",
                entry.name
            )));
        }
        let name = record.raw_name.as_deref().unwrap_or(entry.name.as_bytes());
        let offset = self.next_entry_offset(name)?;
        let too_large = || entry_too_large(&entry.name);
        let header = Header {
            name: name.to_vec(),
            flags: record.flags & (FLAGS_DEFLATE_OPTIONS | FLAG_UTF8),
            method: record.method,
            dos_time: record.dos_time,
            dos_date: record.dos_date,
            crc32: entry.crc32,
            compressed_size: classic(entry.compressed_size).ok_or_else(too_large)?,
            uncompressed_size: classic(entry.uncompressed_size).ok_or_else(too_large)?,
            version_made_by: record.version_made_by,
            version_needed: record.version_needed.max(ZIP_VERSION),
            internal_attributes: record.internal_attributes,
            external_attributes: record.external_attributes,
            offset,
        };
        source.locate_raw(index)?;
        self.commit(|writer| {
            writer.output.emit(&header.local())?;
            source.drain_raw(|bytes| writer.output.emit(bytes))?;
            writer.end_entry(&header);
            Ok(())
        })
    }

    /// Writes the central directory and returns the output; fails while an entry is open.
    pub fn finish(mut self) -> io::Result<W> {
        self.ensure_intact()?;
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

    fn ensure_intact(&self) -> io::Result<()> {
        if self.broken {
            return Err(io::Error::other(
                "the ZIP archive is broken because writing it failed earlier",
            ));
        }
        Ok(())
    }

    fn commit<T>(&mut self, step: impl FnOnce(&mut Self) -> io::Result<T>) -> io::Result<T> {
        let result = step(self);
        self.broken |= result.is_err();
        result
    }

    fn next_entry_offset(&self, name: &[u8]) -> io::Result<u32> {
        self.ensure_intact()?;
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
        classic(self.output.offset).ok_or_else(archive_too_large)
    }

    fn end_entry(&mut self, header: &Header) {
        header.append_central(&mut self.central_directory);
        self.entry_count += 1;
    }

    fn compress_pending(&mut self) -> io::Result<()> {
        self.crc.update(&self.pending);
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

    fn emit(&mut self, mut bytes: &[u8]) -> io::Result<()> {
        self.offset += bytes.len() as u64;
        loop {
            let space = self.buffer.capacity() - self.buffer.len();
            if self.buffer.is_empty() && bytes.len() >= space {
                return self.inner.write_all(bytes);
            }
            let (fitting, rest) = bytes.split_at(space.min(bytes.len()));
            self.buffer.extend_from_slice(fitting);
            if rest.is_empty() {
                return Ok(());
            }
            self.drain()?;
            bytes = rest;
        }
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
        self.ensure_intact()?;
        let Some(open) = self.open.as_mut() else {
            return Err(misuse(
                "cannot write to a ZIP archive without an open entry",
            ));
        };
        let header = &mut open.header;
        header.uncompressed_size =
            classic(u64::from(header.uncompressed_size).saturating_add(buffer.len() as u64))
                .ok_or_else(|| header.too_large())?;
        self.commit(|writer| {
            if buffer.len() > writer.pending.capacity() - writer.pending.len() {
                writer.compress_pending()?;
            }
            if buffer.len() < writer.pending.capacity() {
                writer.pending.extend_from_slice(buffer);
            } else {
                writer.crc.update(buffer);
                writer.output.deflate(&mut writer.deflater, buffer, false)?;
            }
            Ok(buffer.len())
        })
    }

    /// Pushes the archive bytes produced so far; input the compressor still holds follows later.
    fn flush(&mut self) -> io::Result<()> {
        self.ensure_intact()?;
        self.commit(|writer| {
            writer.compress_pending()?;
            writer.output.drain()?;
            writer.output.inner.flush()
        })
    }
}

impl Header {
    fn too_large(&self) -> io::Error {
        entry_too_large(&String::from_utf8_lossy(&self.name))
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
        put_u16(out, self.version_needed);
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

fn entry_too_large(name: &str) -> io::Error {
    needs_zip64(format!("ZIP entry '{name}' is larger than 4 GiB"))
}

fn archive_too_large() -> io::Error {
    needs_zip64("the ZIP archive is larger than 4 GiB")
}

fn needs_zip64(problem: impl Display) -> io::Error {
    unsupported(format!(
        "{problem}, which needs ZIP64 and that is not supported for writing"
    ))
}
