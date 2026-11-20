//! Record layouts, limits and byte helpers shared by the reader and the writer.

use std::io;

pub(crate) const LOCAL_HEADER_SIGNATURE: u32 = 0x0403_4b50;
pub(crate) const CENTRAL_HEADER_SIGNATURE: u32 = 0x0201_4b50;
pub(crate) const END_SIGNATURE: u32 = 0x0605_4b50;
pub(crate) const ZIP64_END_SIGNATURE: u32 = 0x0606_4b50;
pub(crate) const ZIP64_LOCATOR_SIGNATURE: u32 = 0x0706_4b50;
pub(crate) const DATA_DESCRIPTOR_SIGNATURE: u32 = 0x0807_4b50;

pub(crate) const LOCAL_HEADER_LEN: usize = 30;
pub(crate) const CENTRAL_HEADER_LEN: usize = 46;
pub(crate) const END_LEN: usize = 22;
pub(crate) const ZIP64_END_LEN: usize = 56;
pub(crate) const ZIP64_LOCATOR_LEN: usize = 20;
pub(crate) const MAX_COMMENT_LEN: usize = 0xFFFF;

pub(crate) const METHOD_STORED: u16 = 0;
pub(crate) const METHOD_DEFLATED: u16 = 8;

pub(crate) const FLAGS_ENCRYPTED: u16 = (1 << 0) | (1 << 6) | (1 << 13);
pub(crate) const FLAGS_DEFLATE_OPTIONS: u16 = (1 << 1) | (1 << 2);
pub(crate) const FLAG_DATA_DESCRIPTOR: u16 = 1 << 3;
pub(crate) const FLAG_UTF8: u16 = 1 << 11;

pub(crate) const ZIP64_EXTRA_ID: u16 = 0x0001;
pub(crate) const ZIP64_MARKER_32: u32 = u32::MAX;

pub(crate) fn invalid(message: impl Into<String>) -> io::Error {
    io::Error::new(io::ErrorKind::InvalidData, message.into())
}

pub(crate) fn unsupported(message: impl Into<String>) -> io::Error {
    io::Error::new(io::ErrorKind::Unsupported, message.into())
}

pub(crate) fn misuse(message: impl Into<String>) -> io::Error {
    io::Error::new(io::ErrorKind::InvalidInput, message.into())
}

/// A cursor over the little-endian fields of a record; every accessor fails softly at the end.
pub(crate) struct Fields<'a>(&'a [u8]);

impl<'a> Fields<'a> {
    pub(crate) fn new(bytes: &'a [u8]) -> Self {
        Fields(bytes)
    }

    pub(crate) fn remaining(&self) -> usize {
        self.0.len()
    }

    pub(crate) fn bytes(&mut self, len: usize) -> Option<&'a [u8]> {
        let (head, rest) = self.0.split_at_checked(len)?;
        self.0 = rest;
        Some(head)
    }

    pub(crate) fn u16(&mut self) -> Option<u16> {
        let (head, rest) = self.0.split_first_chunk()?;
        self.0 = rest;
        Some(u16::from_le_bytes(*head))
    }

    pub(crate) fn u32(&mut self) -> Option<u32> {
        let (head, rest) = self.0.split_first_chunk()?;
        self.0 = rest;
        Some(u32::from_le_bytes(*head))
    }

    pub(crate) fn u64(&mut self) -> Option<u64> {
        let (head, rest) = self.0.split_first_chunk()?;
        self.0 = rest;
        Some(u64::from_le_bytes(*head))
    }
}

pub(crate) fn put_u16(out: &mut Vec<u8>, value: u16) {
    out.extend_from_slice(&value.to_le_bytes());
}

pub(crate) fn put_u32(out: &mut Vec<u8>, value: u32) {
    out.extend_from_slice(&value.to_le_bytes());
}
