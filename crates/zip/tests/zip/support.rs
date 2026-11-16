//! Deterministic test data, hand-built archives and instrumented I/O.

use std::cell::RefCell;
use std::io::{self, Cursor, Read, Seek, SeekFrom, Write};
use std::rc::Rc;

use excel_zip::{ZipReader, ZipWriter};
use flate2::{Compress, Compression, FlushCompress, Status};

pub const FLAG_ENCRYPTED: u16 = 1;
pub const FLAG_DATA_DESCRIPTOR: u16 = 1 << 3;
pub const FLAG_UTF8: u16 = 1 << 11;
pub const METHOD_STORED: u16 = 0;
pub const METHOD_DEFLATED: u16 = 8;

/// Bytes that do not compress.
pub fn noise(len: usize, seed: u64) -> Vec<u8> {
    let mut state = seed.wrapping_mul(0x9E37_79B9_7F4A_7C15) | 1;
    let mut bytes = Vec::with_capacity(len + 8);
    while bytes.len() < len {
        state ^= state << 13;
        state ^= state >> 7;
        state ^= state << 17;
        bytes.extend_from_slice(&state.to_le_bytes());
    }
    bytes.truncate(len);
    bytes
}

/// Sheet-like XML that compresses well.
pub fn rows(len: usize) -> Vec<u8> {
    let mut bytes = Vec::with_capacity(len + 128);
    let mut row = 0u64;
    while bytes.len() < len {
        row += 1;
        let value = row.wrapping_mul(2_654_435_761) % 100_000;
        let line = format!(
            "<row r=\"{row}\"><c r=\"A{row}\" t=\"s\"><v>{}</v></c><c r=\"B{row}\"><v>{value}.5</v></c></row>\n",
            row % 977
        );
        bytes.extend_from_slice(line.as_bytes());
    }
    bytes.truncate(len);
    bytes
}

/// Writes an archive with this crate.
pub fn write_archive(files: &[(&str, &[u8])]) -> Vec<u8> {
    let mut writer = ZipWriter::new(Vec::new());
    for (name, content) in files {
        writer.begin_file(name).unwrap();
        writer.write_all(content).unwrap();
        writer.end_file().unwrap();
    }
    writer.finish().unwrap()
}

pub fn kind<T>(result: io::Result<T>) -> io::ErrorKind {
    match result {
        Ok(_) => panic!("expected an error"),
        Err(error) => error.kind(),
    }
}

pub fn open(bytes: &[u8]) -> io::Result<ZipReader<Cursor<&[u8]>>> {
    ZipReader::new(Cursor::new(bytes))
}

/// Reads one entry to its end in pieces of `chunk_len` and closes it, which checks the CRC.
pub fn read_entry<R: Read + Seek>(
    reader: &mut ZipReader<R>,
    index: usize,
    chunk_len: usize,
) -> io::Result<Vec<u8>> {
    reader.open_entry(index)?;
    let stated = reader.entry_len();
    let mut content = Vec::new();
    let mut chunk = vec![0; chunk_len];
    loop {
        let count = reader.read(&mut chunk)?;
        if count == 0 {
            break;
        }
        content.extend_from_slice(&chunk[..count]);
        assert!(content.len() as u64 <= stated);
    }
    assert_eq!(content.len() as u64, stated);
    reader.close_entry()?;
    Ok(content)
}

/// Reads every entry of an archive, in directory order.
pub fn read_archive(bytes: &[u8]) -> io::Result<Vec<(String, Vec<u8>)>> {
    let mut reader = open(bytes)?;
    let names: Vec<String> = reader
        .entries()
        .iter()
        .map(|entry| entry.name.clone())
        .collect();
    let mut files = Vec::new();
    for (index, name) in names.into_iter().enumerate() {
        files.push((name, read_entry(&mut reader, index, 8192)?));
    }
    Ok(files)
}

pub fn deflate(content: &[u8]) -> Vec<u8> {
    let mut compressor = Compress::new(Compression::new(6), false);
    let mut compressed = Vec::with_capacity(content.len() + content.len() / 8 + 64);
    let status = compressor
        .compress_vec(content, &mut compressed, FlushCompress::Finish)
        .unwrap();
    assert_eq!(status, Status::StreamEnd);
    compressed
}

pub fn crc32(content: &[u8]) -> u32 {
    crc32fast::hash(content)
}

/// One entry of a hand-built archive; the stated values need not match `data`.
#[derive(Clone)]
pub struct RawEntry {
    pub name: Vec<u8>,
    pub flags: u16,
    pub method: u16,
    pub dos_time: u16,
    pub dos_date: u16,
    pub crc32: u32,
    pub compressed_size: u64,
    pub uncompressed_size: u64,
    pub data: Vec<u8>,
    pub local_extra: Vec<u8>,
    pub descriptor: bool,
    pub zip64_sizes: bool,
    pub zip64_offset: bool,
    pub version_made_by: u16,
    pub version_needed: u16,
    pub external_attributes: u32,
}

impl RawEntry {
    pub fn stored(name: &str, content: &[u8]) -> Self {
        RawEntry {
            name: name.as_bytes().to_vec(),
            flags: 0,
            method: METHOD_STORED,
            dos_time: 0x6000,
            dos_date: 0x5A21,
            crc32: crc32(content),
            compressed_size: content.len() as u64,
            uncompressed_size: content.len() as u64,
            data: content.to_vec(),
            local_extra: Vec::new(),
            descriptor: false,
            zip64_sizes: false,
            zip64_offset: false,
            version_made_by: 20,
            version_needed: 20,
            external_attributes: 0,
        }
    }

    pub fn deflated(name: &str, content: &[u8]) -> Self {
        let data = deflate(content);
        RawEntry {
            method: METHOD_DEFLATED,
            compressed_size: data.len() as u64,
            data,
            ..RawEntry::stored(name, content)
        }
    }

    pub fn local_header(&self) -> Vec<u8> {
        let flags = self.flags
            | if self.descriptor {
                FLAG_DATA_DESCRIPTOR
            } else {
                0
            };
        let (crc32, compressed_size, uncompressed_size) = if self.descriptor {
            (0, 0, 0)
        } else {
            (
                self.crc32,
                self.compressed_size as u32,
                self.uncompressed_size as u32,
            )
        };
        let mut out = Vec::new();
        out.extend_from_slice(b"PK\x03\x04");
        out.extend_from_slice(&self.version_needed.to_le_bytes());
        out.extend_from_slice(&flags.to_le_bytes());
        out.extend_from_slice(&self.method.to_le_bytes());
        out.extend_from_slice(&self.dos_time.to_le_bytes());
        out.extend_from_slice(&self.dos_date.to_le_bytes());
        out.extend_from_slice(&crc32.to_le_bytes());
        out.extend_from_slice(&compressed_size.to_le_bytes());
        out.extend_from_slice(&uncompressed_size.to_le_bytes());
        out.extend_from_slice(&(self.name.len() as u16).to_le_bytes());
        out.extend_from_slice(&(self.local_extra.len() as u16).to_le_bytes());
        out.extend_from_slice(&self.name);
        out.extend_from_slice(&self.local_extra);
        out
    }

    pub fn data_descriptor(&self) -> Vec<u8> {
        let mut out = Vec::new();
        out.extend_from_slice(b"PK\x07\x08");
        out.extend_from_slice(&self.crc32.to_le_bytes());
        out.extend_from_slice(&(self.compressed_size as u32).to_le_bytes());
        out.extend_from_slice(&(self.uncompressed_size as u32).to_le_bytes());
        out
    }

    pub fn central_header(&self, header_offset: u64) -> Vec<u8> {
        let flags = self.flags
            | if self.descriptor {
                FLAG_DATA_DESCRIPTOR
            } else {
                0
            };
        let mut extra = Vec::new();
        let mut zip64 = Vec::new();
        let (mut compressed_size, mut uncompressed_size) =
            (self.compressed_size as u32, self.uncompressed_size as u32);
        let mut offset = header_offset as u32;
        if self.zip64_sizes {
            zip64.extend_from_slice(&self.uncompressed_size.to_le_bytes());
            zip64.extend_from_slice(&self.compressed_size.to_le_bytes());
            (compressed_size, uncompressed_size) = (u32::MAX, u32::MAX);
        }
        if self.zip64_offset {
            zip64.extend_from_slice(&header_offset.to_le_bytes());
            offset = u32::MAX;
        }
        if !zip64.is_empty() {
            extra.extend_from_slice(&0x7875u16.to_le_bytes());
            extra.extend_from_slice(&3u16.to_le_bytes());
            extra.extend_from_slice(&[1, 0, 0]);
            extra.extend_from_slice(&1u16.to_le_bytes());
            extra.extend_from_slice(&(zip64.len() as u16).to_le_bytes());
            extra.extend_from_slice(&zip64);
        }
        let mut out = Vec::new();
        out.extend_from_slice(b"PK\x01\x02");
        out.extend_from_slice(&self.version_made_by.to_le_bytes());
        out.extend_from_slice(&self.version_needed.to_le_bytes());
        out.extend_from_slice(&flags.to_le_bytes());
        out.extend_from_slice(&self.method.to_le_bytes());
        out.extend_from_slice(&self.dos_time.to_le_bytes());
        out.extend_from_slice(&self.dos_date.to_le_bytes());
        out.extend_from_slice(&self.crc32.to_le_bytes());
        out.extend_from_slice(&compressed_size.to_le_bytes());
        out.extend_from_slice(&uncompressed_size.to_le_bytes());
        out.extend_from_slice(&(self.name.len() as u16).to_le_bytes());
        out.extend_from_slice(&(extra.len() as u16).to_le_bytes());
        out.extend_from_slice(&0u16.to_le_bytes());
        out.extend_from_slice(&0u16.to_le_bytes());
        out.extend_from_slice(&0u16.to_le_bytes());
        out.extend_from_slice(&self.external_attributes.to_le_bytes());
        out.extend_from_slice(&offset.to_le_bytes());
        out.extend_from_slice(&self.name);
        out.extend_from_slice(&extra);
        out
    }
}

/// The records that close an archive whose directory of `count` entries sits at `offset`.
pub fn end_records(count: u64, len: u64, offset: u64, zip64: bool, comment: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    if zip64 {
        let zip64_end_offset = offset + len;
        out.extend_from_slice(b"PK\x06\x06");
        out.extend_from_slice(&44u64.to_le_bytes());
        out.extend_from_slice(&45u16.to_le_bytes());
        out.extend_from_slice(&45u16.to_le_bytes());
        out.extend_from_slice(&0u32.to_le_bytes());
        out.extend_from_slice(&0u32.to_le_bytes());
        out.extend_from_slice(&count.to_le_bytes());
        out.extend_from_slice(&count.to_le_bytes());
        out.extend_from_slice(&len.to_le_bytes());
        out.extend_from_slice(&offset.to_le_bytes());
        out.extend_from_slice(b"PK\x06\x07");
        out.extend_from_slice(&0u32.to_le_bytes());
        out.extend_from_slice(&zip64_end_offset.to_le_bytes());
        out.extend_from_slice(&1u32.to_le_bytes());
    }
    let (count, len, offset) = if zip64 {
        (u16::MAX, u32::MAX, u32::MAX)
    } else {
        (count as u16, len as u32, offset as u32)
    };
    out.extend_from_slice(b"PK\x05\x06");
    out.extend_from_slice(&0u16.to_le_bytes());
    out.extend_from_slice(&0u16.to_le_bytes());
    out.extend_from_slice(&count.to_le_bytes());
    out.extend_from_slice(&count.to_le_bytes());
    out.extend_from_slice(&len.to_le_bytes());
    out.extend_from_slice(&offset.to_le_bytes());
    out.extend_from_slice(&(comment.len() as u16).to_le_bytes());
    out.extend_from_slice(comment);
    out
}

/// A hand-built archive; `prefix` is put in front without adjusting any offset.
#[derive(Default)]
pub struct RawArchive {
    pub entries: Vec<RawEntry>,
    pub prefix: Vec<u8>,
    pub comment: Vec<u8>,
    pub zip64: bool,
}

impl RawArchive {
    pub fn of(entries: Vec<RawEntry>) -> Self {
        RawArchive {
            entries,
            ..RawArchive::default()
        }
    }

    pub fn build(&self) -> Vec<u8> {
        let mut body = Vec::new();
        let mut directory = Vec::new();
        for entry in &self.entries {
            directory.extend_from_slice(&entry.central_header(body.len() as u64));
            body.extend_from_slice(&entry.local_header());
            body.extend_from_slice(&entry.data);
            if entry.descriptor {
                body.extend_from_slice(&entry.data_descriptor());
            }
        }
        let (count, len, offset) = (
            self.entries.len() as u64,
            directory.len() as u64,
            body.len() as u64,
        );
        let mut out = self.prefix.clone();
        out.extend_from_slice(&body);
        out.extend_from_slice(&directory);
        out.extend_from_slice(&end_records(count, len, offset, self.zip64, &self.comment));
        out
    }
}

/// What a central directory record says, parsed without any tolerance.
#[derive(Debug, PartialEq, Eq)]
pub struct Central {
    pub name: Vec<u8>,
    pub version_made_by: u16,
    pub version_needed: u16,
    pub flags: u16,
    pub method: u16,
    pub dos_time: u16,
    pub dos_date: u16,
    pub crc32: u32,
    pub compressed_size: u32,
    pub uncompressed_size: u32,
    pub extra_len: u16,
    pub external_attributes: u32,
    pub header_offset: u32,
}

fn u16_at(bytes: &[u8], at: usize) -> u16 {
    u16::from_le_bytes(bytes[at..at + 2].try_into().unwrap())
}

fn u32_at(bytes: &[u8], at: usize) -> u32 {
    u32::from_le_bytes(bytes[at..at + 4].try_into().unwrap())
}

/// Parses the directory of a classic archive that has no comment.
pub fn central_directory(bytes: &[u8]) -> Vec<Central> {
    let end = bytes.len() - 22;
    assert_eq!(&bytes[end..end + 4], b"PK\x05\x06");
    let count = u16_at(bytes, end + 10);
    let mut at = u32_at(bytes, end + 16) as usize;
    assert_eq!(at + u32_at(bytes, end + 12) as usize, end);
    let mut records = Vec::new();
    for _ in 0..count {
        assert_eq!(&bytes[at..at + 4], b"PK\x01\x02");
        let name_len = u16_at(bytes, at + 28) as usize;
        let extra_len = u16_at(bytes, at + 30);
        let comment_len = u16_at(bytes, at + 32) as usize;
        records.push(Central {
            name: bytes[at + 46..at + 46 + name_len].to_vec(),
            version_made_by: u16_at(bytes, at + 4),
            version_needed: u16_at(bytes, at + 6),
            flags: u16_at(bytes, at + 8),
            method: u16_at(bytes, at + 10),
            dos_time: u16_at(bytes, at + 12),
            dos_date: u16_at(bytes, at + 14),
            crc32: u32_at(bytes, at + 16),
            compressed_size: u32_at(bytes, at + 20),
            uncompressed_size: u32_at(bytes, at + 24),
            extra_len,
            external_attributes: u32_at(bytes, at + 38),
            header_offset: u32_at(bytes, at + 42),
        });
        at += 46 + name_len + extra_len as usize + comment_len;
    }
    assert_eq!(at, end);
    records
}

/// What a watched file was asked for: where every read started and how much it returned.
#[derive(Default)]
pub struct ReadLog {
    pub reads: Vec<(u64, usize)>,
    pub seeks: usize,
}

impl ReadLog {
    pub fn bytes_read(&self) -> u64 {
        self.reads.iter().map(|(_, count)| *count as u64).sum()
    }
}

/// A file that records how it is read, in a log that outlives the reader owning the file.
pub struct Watched<R> {
    inner: R,
    limit: Option<usize>,
    failing_read: Option<usize>,
    log: Rc<RefCell<ReadLog>>,
}

impl<R> Watched<R> {
    pub fn new(inner: R) -> (Self, Rc<RefCell<ReadLog>>) {
        let log = Rc::new(RefCell::new(ReadLog::default()));
        let file = Watched {
            inner,
            limit: None,
            failing_read: None,
            log: Rc::clone(&log),
        };
        (file, log)
    }

    /// Hands out at most `limit` bytes per read, like a slow network file.
    pub fn trickling(inner: R, limit: usize) -> Self {
        Watched {
            limit: Some(limit),
            ..Watched::new(inner).0
        }
    }

    /// Fails the read with this index, counted from zero, and serves all others.
    pub fn failing(inner: R, failing_read: usize) -> Self {
        Watched {
            failing_read: Some(failing_read),
            ..Watched::new(inner).0
        }
    }
}

impl<R: Read + Seek> Read for Watched<R> {
    fn read(&mut self, buffer: &mut [u8]) -> io::Result<usize> {
        let mut log = self.log.borrow_mut();
        if self.failing_read == Some(log.reads.len()) {
            log.reads.push((self.inner.stream_position()?, 0));
            return Err(io::Error::new(
                io::ErrorKind::TimedOut,
                "the connection dropped",
            ));
        }
        let offset = self.inner.stream_position()?;
        let limit = self.limit.unwrap_or(buffer.len()).min(buffer.len());
        let count = self.inner.read(&mut buffer[..limit])?;
        log.reads.push((offset, count));
        Ok(count)
    }
}

impl<R: Seek> Seek for Watched<R> {
    fn seek(&mut self, position: SeekFrom) -> io::Result<u64> {
        self.log.borrow_mut().seeks += 1;
        self.inner.seek(position)
    }
}

/// A huge file of zeros with a few real pieces, to reach offsets beyond 4 GiB cheaply.
pub struct Sparse {
    len: u64,
    position: u64,
    pieces: Vec<(u64, Vec<u8>)>,
}

impl Sparse {
    pub fn new() -> Self {
        Sparse {
            len: 0,
            position: 0,
            pieces: Vec::new(),
        }
    }

    pub fn len(&self) -> u64 {
        self.len
    }

    pub fn append(&mut self, bytes: Vec<u8>) {
        let offset = self.len;
        self.len += bytes.len() as u64;
        self.pieces.push((offset, bytes));
    }

    pub fn append_zeros(&mut self, count: u64) {
        self.len += count;
    }
}

impl Read for Sparse {
    fn read(&mut self, buffer: &mut [u8]) -> io::Result<usize> {
        let count = (self.len - self.position.min(self.len)).min(buffer.len() as u64) as usize;
        let (start, end) = (self.position, self.position + count as u64);
        buffer[..count].fill(0);
        for (offset, bytes) in &self.pieces {
            let (from, to) = (start.max(*offset), end.min(*offset + bytes.len() as u64));
            if from < to {
                let piece = &bytes[(from - offset) as usize..(to - offset) as usize];
                buffer[(from - start) as usize..(to - start) as usize].copy_from_slice(piece);
            }
        }
        self.position = end;
        Ok(count)
    }
}

impl Seek for Sparse {
    fn seek(&mut self, position: SeekFrom) -> io::Result<u64> {
        self.position = match position {
            SeekFrom::Start(offset) => offset,
            SeekFrom::End(delta) => self.len.checked_add_signed(delta).unwrap(),
            SeekFrom::Current(delta) => self.position.checked_add_signed(delta).unwrap(),
        };
        Ok(self.position)
    }
}

/// An output that cannot seek and records the size of every write; one chosen write can fail.
#[derive(Default)]
pub struct Appended {
    pub bytes: Vec<u8>,
    pub writes: Vec<usize>,
    pub flushes: usize,
    pub failing_write: Option<usize>,
    attempts: usize,
}

impl Appended {
    /// Fails the write with this index, counted from zero, and accepts all others.
    pub fn failing(failing_write: usize) -> Self {
        Appended {
            failing_write: Some(failing_write),
            ..Appended::default()
        }
    }
}

impl Write for Appended {
    fn write(&mut self, buffer: &[u8]) -> io::Result<usize> {
        self.attempts += 1;
        if self.failing_write == Some(self.attempts - 1) {
            return Err(io::Error::new(
                io::ErrorKind::StorageFull,
                "the disk is full",
            ));
        }
        self.writes.push(buffer.len());
        self.bytes.extend_from_slice(buffer);
        Ok(buffer.len())
    }

    fn flush(&mut self) -> io::Result<()> {
        self.flushes += 1;
        Ok(())
    }
}
