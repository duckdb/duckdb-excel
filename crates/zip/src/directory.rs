//! Locates and parses the central directory.
//!
//! The end record is the last one whose comment runs exactly to the end of the file, so that a
//! comment holding the signature is not taken for it; a record followed by stray bytes is the
//! fallback.
//!
//! An archive with data put in front of it keeps its original offsets. Its directory is then
//! found right before the end record instead of where the end record says, and every entry
//! offset is off by the same amount.

use std::borrow::Cow;
use std::cmp::Ordering;
use std::io::{self, Read, Seek, SeekFrom};

use crate::ZipEntry;
use crate::format::{
    CENTRAL_HEADER_LEN, CENTRAL_HEADER_SIGNATURE, END_LEN, END_SIGNATURE, Fields, MAX_COMMENT_LEN,
    ZIP64_END_LEN, ZIP64_END_SIGNATURE, ZIP64_EXTRA_ID, ZIP64_LOCATOR_LEN, ZIP64_LOCATOR_SIGNATURE,
    ZIP64_MARKER_32, invalid, unsupported,
};

/// What the central directory says about an entry beyond the public [`ZipEntry`] fields.
///
/// `raw_name` holds the stored name only when it is not valid UTF-8, where [`ZipEntry::name`]
/// is a lossy rendering of it.
#[derive(Debug)]
pub(crate) struct Record {
    pub(crate) header_offset: u64,
    pub(crate) method: u16,
    pub(crate) flags: u16,
    pub(crate) dos_time: u16,
    pub(crate) dos_date: u16,
    pub(crate) version_made_by: u16,
    pub(crate) internal_attributes: u16,
    pub(crate) external_attributes: u32,
    pub(crate) raw_name: Option<Box<[u8]>>,
}

#[derive(Debug)]
pub(crate) struct Directory {
    pub(crate) file_len: u64,
    pub(crate) entries: Vec<ZipEntry>,
    pub(crate) records: Vec<Record>,
}

/// The last bytes of the file: the end records and, in small archives, the whole directory.
struct Tail {
    start: u64,
    bytes: Vec<u8>,
}

struct End {
    position: u64,
    disk: u32,
    directory_disk: u32,
    entries_on_disk: u64,
    entry_count: u64,
    directory_len: u64,
    directory_offset: u64,
}

pub(crate) fn read_directory<R: Read + Seek>(inner: &mut R) -> io::Result<Directory> {
    let file_len = inner.seek(SeekFrom::End(0))?;
    let tail_len = (END_LEN + MAX_COMMENT_LEN + ZIP64_LOCATOR_LEN) as u64;
    let tail_len = file_len.min(tail_len) as usize;
    let start = file_len - tail_len as u64;
    let mut bytes = vec![0; tail_len];
    inner.seek(SeekFrom::Start(start))?;
    read_exact(inner, &mut bytes)?;
    let tail = Tail { start, bytes };

    let end = read_end(inner, &tail)?;
    if end.disk != 0 || end.directory_disk != 0 || end.entries_on_disk != end.entry_count {
        return Err(unsupported("multi-disk ZIP archives are not supported"));
    }
    let (directory, shift) = load_directory(inner, &tail, &end)?;
    let (entries, records) = parse_entries(&directory, end.entry_count, shift)?;
    Ok(Directory {
        file_len,
        entries,
        records,
    })
}

fn read_exact<R: Read>(inner: &mut R, buffer: &mut [u8]) -> io::Result<()> {
    inner
        .read_exact(buffer)
        .map_err(|error| match error.kind() {
            io::ErrorKind::UnexpectedEof => invalid("unexpected end of file"),
            _ => error,
        })
}

impl Tail {
    fn load<R: Read + Seek>(
        &self,
        inner: &mut R,
        offset: u64,
        len: usize,
    ) -> io::Result<Cow<'_, [u8]>> {
        let buffered = offset
            .checked_sub(self.start)
            .and_then(|skip| usize::try_from(skip).ok())
            .and_then(|skip| self.bytes.get(skip..)?.get(..len));
        if let Some(bytes) = buffered {
            return Ok(Cow::Borrowed(bytes));
        }
        let mut bytes = Vec::new();
        bytes.try_reserve_exact(len).map_err(|_| {
            io::Error::new(
                io::ErrorKind::OutOfMemory,
                "the central directory does not fit in memory",
            )
        })?;
        bytes.resize(len, 0);
        inner.seek(SeekFrom::Start(offset))?;
        read_exact(inner, &mut bytes)?;
        Ok(Cow::Owned(bytes))
    }
}

fn find_end(tail: &[u8]) -> Option<usize> {
    let last = tail.len().checked_sub(END_LEN)?;
    let mut followed_by_more = None;
    for at in (0..=last).rev() {
        let mut fields = Fields::new(tail.get(at..).unwrap_or_default());
        if fields.u32() != Some(END_SIGNATURE) || fields.bytes(END_LEN - 6).is_none() {
            continue;
        }
        match fields
            .u16()
            .map(|comment_len| usize::from(comment_len).cmp(&fields.remaining()))
        {
            Some(Ordering::Equal) => return Some(at),
            Some(Ordering::Less) => followed_by_more = followed_by_more.or(Some(at)),
            _ => {}
        }
    }
    followed_by_more
}

fn read_end<R: Read + Seek>(inner: &mut R, tail: &Tail) -> io::Result<End> {
    let missing = || invalid("not a ZIP archive: the end of central directory record is missing");
    let at = find_end(&tail.bytes).ok_or_else(missing)?;
    if let Some(end) = read_zip64_end(inner, tail, at)? {
        return Ok(end);
    }
    let record = tail.bytes.get(at..).unwrap_or_default();
    parse_end(record, tail.start + at as u64).ok_or_else(missing)
}

fn parse_end(record: &[u8], position: u64) -> Option<End> {
    let mut fields = Fields::new(record);
    let _signature = fields.u32()?;
    Some(End {
        position,
        disk: fields.u16()?.into(),
        directory_disk: fields.u16()?.into(),
        entries_on_disk: fields.u16()?.into(),
        entry_count: fields.u16()?.into(),
        directory_len: fields.u32()?.into(),
        directory_offset: fields.u32()?.into(),
    })
}

fn read_zip64_end<R: Read + Seek>(
    inner: &mut R,
    tail: &Tail,
    end_at: usize,
) -> io::Result<Option<End>> {
    let Some(locator_at) = end_at.checked_sub(ZIP64_LOCATOR_LEN) else {
        return Ok(None);
    };
    let mut locator = Fields::new(tail.bytes.get(locator_at..end_at).unwrap_or_default());
    if locator.u32() != Some(ZIP64_LOCATOR_SIGNATURE) {
        return Ok(None);
    }
    let (Some(_disk), Some(position)) = (locator.u32(), locator.u64()) else {
        return Ok(None);
    };
    let locator_position = tail.start + locator_at as u64;
    if position
        .checked_add(ZIP64_END_LEN as u64)
        .is_none_or(|end| end > locator_position)
    {
        return Err(invalid(
            "the ZIP64 end of central directory record lies outside the file",
        ));
    }
    let record = tail.load(inner, position, ZIP64_END_LEN)?;
    match parse_zip64_end(&record, position) {
        Some(end) => Ok(Some(end)),
        None => Err(invalid(
            "the ZIP64 end of central directory record is corrupt",
        )),
    }
}

fn parse_zip64_end(record: &[u8], position: u64) -> Option<End> {
    let mut fields = Fields::new(record);
    if fields.u32()? != ZIP64_END_SIGNATURE {
        return None;
    }
    let _record_len = fields.u64()?;
    let _version_made_by = fields.u16()?;
    let _version_needed = fields.u16()?;
    Some(End {
        position,
        disk: fields.u32()?,
        directory_disk: fields.u32()?,
        entries_on_disk: fields.u64()?,
        entry_count: fields.u64()?,
        directory_len: fields.u64()?,
        directory_offset: fields.u64()?,
    })
}

fn load_directory<'a, R: Read + Seek>(
    inner: &mut R,
    tail: &'a Tail,
    end: &End,
) -> io::Result<(Cow<'a, [u8]>, u64)> {
    if end.entry_count == 0 {
        return Ok((Cow::Borrowed(&[]), 0));
    }
    let stated = (end.directory_offset, 0);
    let shifted = end
        .position
        .checked_sub(end.directory_len)
        .and_then(|start| {
            let shift = start.checked_sub(end.directory_offset)?;
            (shift > 0).then_some((start, shift))
        });
    for (start, shift) in [Some(stated), shifted].into_iter().flatten() {
        let Some(available) = end.position.checked_sub(start) else {
            continue;
        };
        let Ok(len) = usize::try_from(available.min(end.directory_len)) else {
            continue;
        };
        if len < CENTRAL_HEADER_LEN {
            continue;
        }
        let bytes = tail.load(inner, start, len)?;
        if bytes.starts_with(&CENTRAL_HEADER_SIGNATURE.to_le_bytes()) {
            return Ok((bytes, shift));
        }
    }
    Err(invalid(
        "the central directory of the ZIP archive is missing",
    ))
}

fn parse_entries(
    directory: &[u8],
    count: u64,
    shift: u64,
) -> io::Result<(Vec<ZipEntry>, Vec<Record>)> {
    let capacity = usize::try_from(count)
        .unwrap_or(usize::MAX)
        .min(directory.len() / CENTRAL_HEADER_LEN);
    let mut entries = Vec::with_capacity(capacity);
    let mut records = Vec::with_capacity(capacity);
    let mut fields = Fields::new(directory);
    for index in 0..count {
        let (entry, record) = parse_entry(&mut fields, shift).ok_or_else(|| {
            invalid(format!(
                "the central directory of the ZIP archive is corrupt at entry {index}"
            ))
        })?;
        entries.push(entry);
        records.push(record);
    }
    Ok((entries, records))
}

fn parse_entry(fields: &mut Fields<'_>, shift: u64) -> Option<(ZipEntry, Record)> {
    if fields.u32()? != CENTRAL_HEADER_SIGNATURE {
        return None;
    }
    let version_made_by = fields.u16()?;
    let _version_needed = fields.u16()?;
    let flags = fields.u16()?;
    let method = fields.u16()?;
    let dos_time = fields.u16()?;
    let dos_date = fields.u16()?;
    let crc32 = fields.u32()?;
    let mut compressed_size = u64::from(fields.u32()?);
    let mut uncompressed_size = u64::from(fields.u32()?);
    let name_len = usize::from(fields.u16()?);
    let extra_len = usize::from(fields.u16()?);
    let comment_len = usize::from(fields.u16()?);
    let _disk = fields.u16()?;
    let internal_attributes = fields.u16()?;
    let external_attributes = fields.u32()?;
    let mut header_offset = u64::from(fields.u32()?);
    let name = fields.bytes(name_len)?;
    let extra = fields.bytes(extra_len)?;
    let _comment = fields.bytes(comment_len)?;

    if let Some(mut zip64) = zip64_extra(extra).map(Fields::new) {
        for value in [
            &mut uncompressed_size,
            &mut compressed_size,
            &mut header_offset,
        ] {
            if *value == u64::from(ZIP64_MARKER_32) {
                *value = zip64.u64()?;
            }
        }
    }

    let (name, raw_name) = match String::from_utf8(name.to_vec()) {
        Ok(name) => (name, None),
        Err(error) => {
            let lossy = String::from_utf8_lossy(error.as_bytes()).into_owned();
            (lossy, Some(error.into_bytes().into_boxed_slice()))
        }
    };
    let entry = ZipEntry {
        name,
        compressed_size,
        uncompressed_size,
        crc32,
    };
    let record = Record {
        header_offset: header_offset.checked_add(shift)?,
        method,
        flags,
        dos_time,
        dos_date,
        version_made_by,
        internal_attributes,
        external_attributes,
        raw_name,
    };
    Some((entry, record))
}

fn zip64_extra(extra: &[u8]) -> Option<&[u8]> {
    let mut fields = Fields::new(extra);
    loop {
        let id = fields.u16()?;
        let len = fields.u16()?;
        let data = fields.bytes(len.into())?;
        if id == ZIP64_EXTRA_ID {
            return Some(data);
        }
    }
}
