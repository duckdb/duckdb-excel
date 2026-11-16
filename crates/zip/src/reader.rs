//! Streams the entries of an archive.

use std::io::{self, Read, Seek, SeekFrom};

use flate2::{Decompress, FlushDecompress, Status};

use crate::ZipEntry;
use crate::directory::{Record, read_directory};
use crate::format::{
    FLAGS_ENCRYPTED, Fields, LOCAL_HEADER_LEN, LOCAL_HEADER_SIGNATURE, METHOD_DEFLATED,
    METHOD_STORED, invalid, misuse, unsupported,
};

const BUFFER_LEN: usize = 64 * 1024;

/// Reads an archive, one open entry at a time.
#[derive(Debug)]
pub struct ZipReader<R> {
    source: Source<R>,
    entries: Vec<ZipEntry>,
    records: Vec<Record>,
    idle_inflater: Option<Decompress>,
    open: Option<OpenEntry>,
}

#[derive(Debug)]
struct OpenEntry {
    index: usize,
    len: u64,
    crc32: u32,
    pos: u64,
    crc: crc32fast::Hasher,
    decoder: Decoder,
}

#[derive(Debug)]
enum Decoder {
    Stored,
    Deflated { inflater: Decompress, ended: bool },
}

/// The compressed bytes of one entry, fetched from the file in large pieces.
#[derive(Debug)]
struct Source<R> {
    inner: R,
    file_len: u64,
    buffer: Vec<u8>,
    start: usize,
    end: usize,
    unread: u64,
}

impl<R: Read + Seek> ZipReader<R> {
    /// Reads the central directory of the archive in `inner`.
    pub fn new(mut inner: R) -> io::Result<Self> {
        let directory = read_directory(&mut inner)?;
        Ok(ZipReader {
            source: Source {
                inner,
                file_len: directory.file_len,
                buffer: Vec::new(),
                start: 0,
                end: 0,
                unread: 0,
            },
            entries: directory.entries,
            records: directory.records,
            idle_inflater: None,
            open: None,
        })
    }

    /// The entries in central directory order.
    pub fn entries(&self) -> &[ZipEntry] {
        &self.entries
    }

    /// The index of the last entry called `name`: producers that repeat a name mean the last one.
    pub fn find(&self, name: &str) -> Option<usize> {
        self.entries.iter().rposition(|entry| entry.name == name)
    }

    /// Opens the entry at `index` for reading through [`Read`]; only one entry can be open.
    pub fn open_entry(&mut self, index: usize) -> io::Result<()> {
        if self.open.is_some() {
            return Err(misuse("cannot open a ZIP entry while another one is open"));
        }
        let (entry, record) = self.stored_entry(index)?;
        let name = &entry.name;
        if record.flags & FLAGS_ENCRYPTED != 0 {
            return Err(unsupported(format!("ZIP entry '{name}' is encrypted")));
        }
        let deflated = match record.method {
            METHOD_STORED if entry.compressed_size == entry.uncompressed_size => false,
            METHOD_STORED => {
                return Err(invalid(format!(
                    "ZIP entry '{name}' is stored but its sizes differ"
                )));
            }
            METHOD_DEFLATED => true,
            method => {
                return Err(unsupported(format!(
                    "ZIP entry '{name}' uses unsupported compression method {method}"
                )));
            }
        };
        let (len, crc32) = (entry.uncompressed_size, entry.crc32);
        self.locate(index)?;
        let decoder = if deflated {
            let inflater = match self.idle_inflater.take() {
                Some(mut inflater) => {
                    inflater.reset(false);
                    inflater
                }
                None => Decompress::new(false),
            };
            Decoder::Deflated {
                inflater,
                ended: false,
            }
        } else {
            Decoder::Stored
        };
        self.open = Some(OpenEntry {
            index,
            len,
            crc32,
            pos: 0,
            crc: crc32fast::Hasher::new(),
            decoder,
        });
        Ok(())
    }

    pub fn is_entry_open(&self) -> bool {
        self.open.is_some()
    }

    /// The uncompressed size of the open entry, as the central directory states it.
    pub fn entry_len(&self) -> u64 {
        self.open.as_ref().map_or(0, |open| open.len)
    }

    /// The number of uncompressed bytes read from the open entry.
    pub fn entry_pos(&self) -> u64 {
        self.open.as_ref().map_or(0, |open| open.pos)
    }

    /// Whether the open entry has been read up to its stated size.
    pub fn is_done(&self) -> bool {
        self.entry_pos() >= self.entry_len()
    }

    /// Closes the open entry; its checksum is only verified when it was read to the end.
    pub fn close_entry(&mut self) -> io::Result<()> {
        let Some(open) = self.open.take() else {
            return Err(misuse("cannot close a ZIP entry when none is open"));
        };
        let complete = open.pos >= open.len;
        let name = entry_name(&self.entries, open.index);
        let mut surplus = Ok(false);
        if let Decoder::Deflated {
            mut inflater,
            ended,
        } = open.decoder
        {
            if complete && !ended {
                surplus = inflate(&mut self.source, &mut inflater, &mut [0], name)
                    .map(|(count, _)| count > 0);
            }
            self.idle_inflater = Some(inflater);
        }
        if !complete {
            return Ok(());
        }
        if surplus? {
            return Err(invalid(format!(
                "ZIP entry '{name}' holds more data than its stated size"
            )));
        }
        if open.crc.finalize() != open.crc32 {
            return Err(invalid(format!(
                "ZIP entry '{name}' does not match its checksum"
            )));
        }
        Ok(())
    }

    pub fn into_inner(self) -> R {
        self.source.inner
    }

    pub(crate) fn stored_entry(&self, index: usize) -> io::Result<(&ZipEntry, &Record)> {
        match (self.entries.get(index), self.records.get(index)) {
            (Some(entry), Some(record)) => Ok((entry, record)),
            _ => Err(misuse(format!("the ZIP archive has no entry {index}"))),
        }
    }

    pub(crate) fn locate_raw(&mut self, index: usize) -> io::Result<()> {
        if self.open.is_some() {
            return Err(misuse(
                "cannot copy a ZIP entry while the source has an open entry",
            ));
        }
        self.locate(index)
    }

    pub(crate) fn drain_raw(
        &mut self,
        mut sink: impl FnMut(&[u8]) -> io::Result<()>,
    ) -> io::Result<()> {
        while self.source.has_input()? {
            let chunk = self.source.available();
            let count = chunk.len();
            sink(chunk)?;
            self.source.consume(count);
        }
        Ok(())
    }

    fn locate(&mut self, index: usize) -> io::Result<()> {
        let (entry, record) = self.stored_entry(index)?;
        let name_len = record
            .raw_name
            .as_ref()
            .map_or(entry.name.len(), |name| name.len());
        let (header_offset, compressed_size) = (record.header_offset, entry.compressed_size);
        self.source.locate(header_offset, name_len, compressed_size)
    }
}

/// Reads the uncompressed bytes of the open entry.
impl<R: Read + Seek> Read for ZipReader<R> {
    fn read(&mut self, buffer: &mut [u8]) -> io::Result<usize> {
        let Some(open) = self.open.as_mut() else {
            return Err(misuse(
                "cannot read from a ZIP archive without an open entry",
            ));
        };
        let left = open.len.saturating_sub(open.pos);
        let wanted = usize::try_from(left).map_or(buffer.len(), |left| left.min(buffer.len()));
        let Some(output) = buffer.get_mut(..wanted).filter(|output| !output.is_empty()) else {
            return Ok(0);
        };
        let name = entry_name(&self.entries, open.index);
        let count = match &mut open.decoder {
            Decoder::Stored => self.source.read(output)?,
            Decoder::Deflated { ended: true, .. } => 0,
            Decoder::Deflated { inflater, ended } => {
                let (count, stream_ended) = inflate(&mut self.source, inflater, output, name)?;
                *ended = stream_ended;
                count
            }
        };
        let Some(produced) = output.get(..count).filter(|produced| !produced.is_empty()) else {
            return Err(invalid(format!(
                "ZIP entry '{name}' ends before its stated size"
            )));
        };
        open.crc.update(produced);
        open.pos += count as u64;
        Ok(count)
    }
}

fn entry_name(entries: &[ZipEntry], index: usize) -> &str {
    entries.get(index).map_or("", |entry| entry.name.as_str())
}

fn inflate<R: Read + Seek>(
    source: &mut Source<R>,
    inflater: &mut Decompress,
    output: &mut [u8],
    name: &str,
) -> io::Result<(usize, bool)> {
    loop {
        source.has_input()?;
        let (consumed_before, produced_before) = (inflater.total_in(), inflater.total_out());
        let status = inflater
            .decompress(source.available(), output, FlushDecompress::None)
            .map_err(|_| invalid(format!("ZIP entry '{name}' holds a corrupt deflate stream")))?;
        let consumed = (inflater.total_in() - consumed_before) as usize;
        let produced = (inflater.total_out() - produced_before) as usize;
        source.consume(consumed);
        let ended = status == Status::StreamEnd;
        if produced > 0 || ended || consumed == 0 {
            return Ok((produced, ended));
        }
    }
}

impl<R: Read + Seek> Source<R> {
    fn locate(
        &mut self,
        header_offset: u64,
        name_len: usize,
        compressed_size: u64,
    ) -> io::Result<()> {
        (self.start, self.end, self.unread) = (0, 0, 0);
        let outside = || invalid("ZIP entry lies outside the file");
        let header_len = LOCAL_HEADER_LEN as u64;
        let file_left = self
            .file_len
            .checked_sub(header_offset)
            .filter(|left| *left >= header_len)
            .ok_or_else(outside)?;
        let wanted = (header_len + name_len as u64)
            .saturating_add(compressed_size)
            .min(file_left);
        let wanted = wanted.min(BUFFER_LEN as u64) as usize;

        self.buffer.resize(BUFFER_LEN, 0);
        self.inner.seek(SeekFrom::Start(header_offset))?;
        let mut filled = 0;
        while filled < LOCAL_HEADER_LEN {
            let space = self.buffer.get_mut(filled..wanted).unwrap_or_default();
            filled += read_some(&mut self.inner, space)?;
        }

        let mut header = Fields::new(&self.buffer);
        if header.u32() != Some(LOCAL_HEADER_SIGNATURE) {
            return Err(invalid("ZIP entry does not start with a local file header"));
        }
        let lengths = header
            .bytes(LOCAL_HEADER_LEN - 8)
            .and_then(|_| Some((header.u16()?, header.u16()?)));
        let (local_name_len, local_extra_len) = lengths.ok_or_else(outside)?;
        let skip = LOCAL_HEADER_LEN + usize::from(local_name_len) + usize::from(local_extra_len);
        let data_offset = header_offset.checked_add(skip as u64).ok_or_else(outside)?;
        if data_offset
            .checked_add(compressed_size)
            .is_none_or(|data_end| data_end > self.file_len)
        {
            return Err(outside());
        }

        if let Some(buffered) = filled.checked_sub(skip) {
            let buffered = (buffered as u64).min(compressed_size);
            self.start = skip;
            self.end = skip + buffered as usize;
            self.unread = compressed_size - buffered;
        } else {
            self.inner.seek(SeekFrom::Start(data_offset))?;
            self.unread = compressed_size;
        }
        Ok(())
    }

    fn available(&self) -> &[u8] {
        self.buffer.get(self.start..self.end).unwrap_or_default()
    }

    fn consume(&mut self, count: usize) {
        self.start = self.start.saturating_add(count).min(self.end);
    }

    fn has_input(&mut self) -> io::Result<bool> {
        if self.start < self.end {
            return Ok(true);
        }
        if self.unread == 0 {
            return Ok(false);
        }
        let wanted = self.unread.min(self.buffer.len() as u64) as usize;
        let count = read_some(
            &mut self.inner,
            self.buffer.get_mut(..wanted).unwrap_or_default(),
        )?;
        (self.start, self.end) = (0, count);
        self.unread -= count as u64;
        Ok(true)
    }

    fn read(&mut self, output: &mut [u8]) -> io::Result<usize> {
        if self.start == self.end && output.len() >= self.buffer.len() {
            let wanted = self.unread.min(output.len() as u64) as usize;
            let Some(space) = output.get_mut(..wanted).filter(|space| !space.is_empty()) else {
                return Ok(0);
            };
            let count = read_some(&mut self.inner, space)?;
            self.unread -= count as u64;
            return Ok(count);
        }
        if !self.has_input()? {
            return Ok(0);
        }
        let available = self.available();
        let count = available.len().min(output.len());
        output[..count].copy_from_slice(&available[..count]);
        self.consume(count);
        Ok(count)
    }
}

fn read_some<R: Read>(inner: &mut R, buffer: &mut [u8]) -> io::Result<usize> {
    loop {
        match inner.read(buffer) {
            Ok(0) => return Err(invalid("unexpected end of file inside a ZIP entry")),
            Ok(count) => return Ok(count.min(buffer.len())),
            Err(error) if error.kind() == io::ErrorKind::Interrupted => {}
            Err(error) => return Err(error),
        }
    }
}
