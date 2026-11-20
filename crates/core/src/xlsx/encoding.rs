//! The encodings an XML part may be in, and the conversion of the ones that are not UTF-8 to it.

use std::{
    fmt,
    io::{self, Read},
};

const UTF8_BOM: &[u8] = b"\xEF\xBB\xBF";
const RAW_BUFFER_LEN: usize = 16 * 1024;

/// The most UTF-8 bytes one character of another encoding becomes.
const MAX_CHARACTER_LEN: usize = 4;

/// How the bytes of a document become characters.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Encoding {
    Utf8,
    Latin1,
    Utf16Le,
    Utf16Be,
}

/// What the encoding a document declares is to the encoding it is in.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Declared {
    Fits,
    Differs,
    Unknown,
}

impl Encoding {
    /// The most bytes of the start of a document it takes to tell its encoding.
    pub const DETECTION_LEN: usize = 4096;

    /// The encoding of the document that starts with `start`, and how many bytes of a byte order mark to pass
    /// over. `None` asks for more bytes, which a document that `is_whole` does not have.
    pub fn of_document(start: &[u8], is_whole: bool) -> Option<(Encoding, usize)> {
        let more = || is_whole.then_some((Encoding::Utf8, 0));
        match start {
            [] | [_] => return more(),
            [0xFF, 0xFE, ..] => return Some((Encoding::Utf16Le, 2)),
            [0xFE, 0xFF, ..] => return Some((Encoding::Utf16Be, 2)),
            [0, _, ..] => return Some((Encoding::Utf16Be, 0)),
            [_, 0, ..] => return Some((Encoding::Utf16Le, 0)),
            _ => {}
        }
        let (body, skipped) = match start.strip_prefix(UTF8_BOM) {
            Some(body) => (body, UTF8_BOM.len()),
            None if UTF8_BOM.starts_with(start) => return more(),
            None => (start, 0),
        };
        const DECLARATION_START: &[u8] = b"<?xml";
        let Some(declaration) = body.strip_prefix(DECLARATION_START) else {
            return if DECLARATION_START.starts_with(body) {
                more()
            } else {
                Some((Encoding::Utf8, 0))
            };
        };
        // `<?xml-stylesheet` and the like are no declaration.
        match declaration.first() {
            Some(byte) if is_whitespace(*byte) => {}
            Some(_) => return Some((Encoding::Utf8, 0)),
            None => return more(),
        }
        let Some(end) = declaration.windows(2).position(|pair| pair == b"?>") else {
            return more();
        };
        let declared = declared_encoding(declaration.get(..end).unwrap_or_default());
        if declared.is_some_and(|name| name.eq_ignore_ascii_case(b"iso-8859-1")) {
            Some((Encoding::Latin1, skipped))
        } else {
            Some((Encoding::Utf8, 0))
        }
    }

    /// Whether a document in this encoding may declare itself as `declared`.
    pub fn accepts(self, declared: &str) -> Declared {
        let is = |name: &str| declared.eq_ignore_ascii_case(name);
        let fits = match self {
            Encoding::Utf8 => is("utf-8") || is("us-ascii"),
            Encoding::Latin1 => is("iso-8859-1"),
            Encoding::Utf16Le => is("utf-16") || is("utf-16le"),
            Encoding::Utf16Be => is("utf-16") || is("utf-16be"),
        };
        let is_known = [
            "utf-8",
            "us-ascii",
            "iso-8859-1",
            "utf-16",
            "utf-16le",
            "utf-16be",
        ]
        .iter()
        .any(|name| is(name));
        match (fits, is_known) {
            (true, _) => Declared::Fits,
            (false, true) => Declared::Differs,
            (false, false) => Declared::Unknown,
        }
    }
}

fn is_whitespace(byte: u8) -> bool {
    matches!(byte, b' ' | b'\t' | b'\n' | b'\r')
}

fn declared_encoding(mut declaration: &[u8]) -> Option<&[u8]> {
    loop {
        declaration = declaration.trim_ascii_start();
        let name_len = declaration
            .iter()
            .position(|byte| *byte == b'=' || is_whitespace(*byte))?;
        let (name, rest) = declaration.split_at_checked(name_len)?;
        let rest = rest
            .trim_ascii_start()
            .strip_prefix(b"=")?
            .trim_ascii_start();
        let (quote, rest) = rest
            .split_first()
            .filter(|(quote, _)| matches!(quote, b'"' | b'\''))?;
        let value_len = rest.iter().position(|byte| byte == quote)?;
        let (value, rest) = rest.split_at_checked(value_len)?;
        if name == b"encoding" {
            return Some(value);
        }
        declaration = rest.get(1..)?;
    }
}

/// Marks the I/O error for bytes that are not text in the encoding of their document.
#[derive(Debug)]
pub struct MalformedText(pub &'static str);

impl fmt::Display for MalformedText {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str(self.0)
    }
}

impl std::error::Error for MalformedText {}

fn malformed(reason: &'static str) -> io::Error {
    io::Error::other(MalformedText(reason))
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Scheme {
    Latin1,
    Utf16 { is_big_endian: bool },
}

/// Reads a document that is not in UTF-8 and hands it on as UTF-8, a buffer at a time.
#[derive(Debug)]
pub struct Transcoder {
    scheme: Scheme,
    raw: Box<[u8]>,
    start: usize,
    end: usize,
    raw_bytes_read: u64,
    is_at_end: bool,
}

impl Transcoder {
    /// A transcoder for a document in `encoding` of which `start`, at most [`Encoding::DETECTION_LEN`]
    /// bytes, was read already; `None` for UTF-8, which needs none.
    pub fn new(encoding: Encoding, start: &[u8]) -> Option<Transcoder> {
        let scheme = match encoding {
            Encoding::Utf8 => return None,
            Encoding::Latin1 => Scheme::Latin1,
            Encoding::Utf16Le => Scheme::Utf16 {
                is_big_endian: false,
            },
            Encoding::Utf16Be => Scheme::Utf16 {
                is_big_endian: true,
            },
        };
        let mut raw = vec![0; RAW_BUFFER_LEN].into_boxed_slice();
        let end = start.len().min(raw.len());
        if let (Some(target), Some(source)) = (raw.get_mut(..end), start.get(..end)) {
            target.copy_from_slice(source);
        }
        Some(Transcoder {
            scheme,
            raw,
            start: 0,
            end,
            raw_bytes_read: 0,
            is_at_end: false,
        })
    }

    /// The number of bytes taken from the source so far, without the ones the transcoder was created with.
    pub fn raw_bytes_read(&self) -> u64 {
        self.raw_bytes_read
    }

    /// Fills `target`, which must have room for a character, with the UTF-8 of what comes next in `source`
    /// and returns the number of bytes written, which is 0 at the end of the document.
    pub fn read<R: Read>(&mut self, source: &mut R, target: &mut [u8]) -> io::Result<usize> {
        if target.len() < MAX_CHARACTER_LEN {
            return Err(io::Error::new(
                io::ErrorKind::InvalidInput,
                "no room for a character",
            ));
        }
        loop {
            let written = self.convert(target)?;
            if written > 0 {
                return Ok(written);
            }
            if self.is_at_end {
                return if self.start == self.end {
                    Ok(0)
                } else {
                    Err(malformed("the document ends in the middle of a character"))
                };
            }
            self.raw.copy_within(self.start..self.end, 0);
            self.end -= self.start;
            self.start = 0;
            let room = self.raw.get_mut(self.end..).unwrap_or_default();
            let count = loop {
                match source.read(room) {
                    Err(error) if error.kind() == io::ErrorKind::Interrupted => {}
                    result => break result?.min(room.len()),
                }
            };
            self.is_at_end = count == 0;
            self.end += count;
            self.raw_bytes_read += count as u64;
        }
    }

    fn convert(&mut self, target: &mut [u8]) -> io::Result<usize> {
        let mut written = 0;
        loop {
            let room = target.get_mut(written..).unwrap_or_default();
            if room.len() < MAX_CHARACTER_LEN {
                return Ok(written);
            }
            let raw = self.raw.get(self.start..self.end).unwrap_or_default();
            let Some((character, len)) = self.scheme.next_character(raw)? else {
                return Ok(written);
            };
            written += character.encode_utf8(room).len();
            self.start += len;
        }
    }
}

fn utf16_unit(raw: &[u8], is_big_endian: bool) -> Option<u16> {
    let pair = *raw.first_chunk::<2>()?;
    Some(if is_big_endian {
        u16::from_be_bytes(pair)
    } else {
        u16::from_le_bytes(pair)
    })
}

impl Scheme {
    fn next_character(self, raw: &[u8]) -> io::Result<Option<(char, usize)>> {
        let is_big_endian = match self {
            Scheme::Latin1 => return Ok(raw.first().map(|byte| (char::from(*byte), 1))),
            Scheme::Utf16 { is_big_endian } => is_big_endian,
        };
        let Some(unit) = utf16_unit(raw, is_big_endian) else {
            return Ok(None);
        };
        if let Some(character) = char::from_u32(u32::from(unit)) {
            return Ok(Some((character, 2)));
        }
        let Some(low) = raw
            .get(2..)
            .and_then(|rest| utf16_unit(rest, is_big_endian))
        else {
            return Ok(None);
        };
        let mut pair = char::decode_utf16([unit, low]);
        match (pair.next(), pair.next()) {
            (Some(Ok(character)), None) => Ok(Some((character, 4))),
            _ => Err(malformed("half a surrogate pair in UTF-16 text")),
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    /// Hands out its bytes `step` at a time.
    struct Trickle<'a> {
        bytes: &'a [u8],
        step: usize,
    }

    impl Read for Trickle<'_> {
        fn read(&mut self, target: &mut [u8]) -> io::Result<usize> {
            let count = self.step.min(self.bytes.len()).min(target.len());
            let (now, later) = self.bytes.split_at(count);
            target[..count].copy_from_slice(now);
            self.bytes = later;
            Ok(count)
        }
    }

    fn transcode(
        encoding: Encoding,
        bytes: &[u8],
        step: usize,
        target_len: usize,
    ) -> io::Result<String> {
        let split = bytes.len().min(3);
        let mut source = Trickle {
            bytes: &bytes[split..],
            step,
        };
        let mut transcoder = Transcoder::new(encoding, &bytes[..split]).expect("a transcoder");
        let mut target = vec![0; target_len];
        let mut utf8 = Vec::new();
        loop {
            let count = transcoder.read(&mut source, &mut target)?;
            if count == 0 {
                assert_eq!(
                    transcoder.raw_bytes_read(),
                    (bytes.len() - split) as u64,
                    "every byte is read"
                );
                return Ok(String::from_utf8(utf8).expect("UTF-8 from the transcoder"));
            }
            utf8.extend_from_slice(&target[..count]);
        }
    }

    fn utf16(text: &str, encoding: Encoding) -> Vec<u8> {
        text.encode_utf16()
            .flat_map(|unit| match encoding {
                Encoding::Utf16Be => unit.to_be_bytes(),
                _ => unit.to_le_bytes(),
            })
            .collect()
    }

    const TEXT: &str = "<r a=\"ß\">żółw \u{1F600}\u{10FFFF} 漢字 \u{FFFD}\u{D7FF}\u{E000} end</r>";

    #[test]
    fn utf16_becomes_utf8_however_it_is_cut_up() {
        for encoding in [Encoding::Utf16Le, Encoding::Utf16Be] {
            let bytes = utf16(TEXT, encoding);
            for step in [1, 2, 3, 5, 7, 4096] {
                for target_len in [4, 5, 7, 64, 4096] {
                    let text = transcode(encoding, &bytes, step, target_len).unwrap();
                    assert_eq!(text, TEXT, "{encoding:?} {step} {target_len}");
                }
            }
        }
    }

    #[test]
    fn latin1_becomes_utf8() {
        let bytes: Vec<u8> = (1..=255).collect();
        let expected: String = (1..=255u8).map(char::from).collect();
        for step in [1, 3, 4096] {
            for target_len in [4, 5, 4096] {
                let text = transcode(Encoding::Latin1, &bytes, step, target_len).unwrap();
                assert_eq!(text, expected);
            }
        }
    }

    #[test]
    fn long_documents_are_read_in_pieces() {
        let text = "aé漢\u{1F600}".repeat(40_000);
        for encoding in [Encoding::Utf16Le, Encoding::Utf16Be] {
            let bytes = utf16(&text, encoding);
            assert_eq!(
                transcode(encoding, &bytes, 10_000, 32 * 1024).unwrap(),
                text
            );
            assert_eq!(
                transcode(encoding, &bytes, 1 << 20, 32 * 1024).unwrap(),
                text
            );
        }
        let latin1: Vec<u8> = b"caf\xe9 ".repeat(40_000);
        assert_eq!(
            transcode(Encoding::Latin1, &latin1, 1 << 20, 32 * 1024).unwrap(),
            "café ".repeat(40_000)
        );
    }

    #[test]
    fn an_empty_document_is_empty() {
        for encoding in [Encoding::Utf16Le, Encoding::Utf16Be, Encoding::Latin1] {
            assert_eq!(transcode(encoding, b"", 1, 16).unwrap(), "");
        }
    }

    #[test]
    fn utf8_needs_no_transcoder_and_a_target_needs_room() {
        assert!(Transcoder::new(Encoding::Utf8, b"<r/>").is_none());
        let mut transcoder = Transcoder::new(Encoding::Latin1, b"abc").unwrap();
        let error = transcoder.read(&mut io::empty(), &mut [0; 3]).unwrap_err();
        assert_eq!(error.kind(), io::ErrorKind::InvalidInput);
    }

    fn failure(encoding: Encoding, bytes: &[u8]) -> String {
        for step in [1, 4096] {
            let error = transcode(encoding, bytes, step, 64).unwrap_err();
            let reason = error
                .get_ref()
                .and_then(|inner| inner.downcast_ref::<MalformedText>())
                .expect("a malformed text error");
            if step == 4096 {
                return reason.0.to_owned();
            }
        }
        unreachable!()
    }

    #[test]
    fn utf16_that_ends_in_a_character_is_an_error() {
        let ends_early = "the document ends in the middle of a character";
        assert_eq!(failure(Encoding::Utf16Le, b"a\0b"), ends_early);
        assert_eq!(failure(Encoding::Utf16Be, b"\0a\0"), ends_early);
        assert_eq!(failure(Encoding::Utf16Le, b"a\0\x3D\xD8"), ends_early);
        assert_eq!(failure(Encoding::Utf16Le, b"a\0\x3D\xD8\x00"), ends_early);
    }

    #[test]
    fn halves_of_surrogate_pairs_are_errors() {
        let half = "half a surrogate pair in UTF-16 text";
        assert_eq!(failure(Encoding::Utf16Le, b"\x3D\xD8a\0"), half);
        assert_eq!(failure(Encoding::Utf16Le, b"a\0\x00\xDEb\0"), half);
        assert_eq!(
            failure(Encoding::Utf16Be, b"\xD8\x3D\xD8\x3D\xDE\x00"),
            half
        );
        assert_eq!(failure(Encoding::Utf16Be, b"\xDE\x00\xD8\x3D"), half);
    }

    fn detect(start: &[u8]) -> Option<(Encoding, usize)> {
        Encoding::of_document(start, false)
    }

    #[test]
    fn utf16_is_told_by_its_first_two_bytes() {
        assert_eq!(detect(b"\xFF\xFE<\0"), Some((Encoding::Utf16Le, 2)));
        assert_eq!(detect(b"\xFE\xFF\0<"), Some((Encoding::Utf16Be, 2)));
        assert_eq!(detect(b"<\0?\0"), Some((Encoding::Utf16Le, 0)));
        assert_eq!(detect(b"\0<\0?"), Some((Encoding::Utf16Be, 0)));
        assert_eq!(detect(b"\0\0"), Some((Encoding::Utf16Be, 0)));
        assert_eq!(detect(b" \0"), Some((Encoding::Utf16Le, 0)));
        assert_eq!(detect(b"\xFF"), None);
        assert_eq!(detect(b"<"), None);
        assert_eq!(detect(b""), None);
    }

    #[test]
    fn latin1_is_told_by_the_declaration() {
        for declaration in [
            "<?xml version=\"1.0\" encoding=\"ISO-8859-1\"?>",
            "<?xml version='1.0' encoding='iso-8859-1' standalone='yes'?>",
            "<?xml\tversion = \"1.0\"\n encoding\r\n=\n'Iso-8859-1' ?>",
            "<?xml encoding=\"ISO-8859-1\"?>",
        ] {
            let document = format!("{declaration}<r/>");
            assert_eq!(
                detect(document.as_bytes()),
                Some((Encoding::Latin1, 0)),
                "{declaration}"
            );
            let with_mark = [UTF8_BOM, document.as_bytes()].concat();
            assert_eq!(detect(&with_mark), Some((Encoding::Latin1, 3)));
        }
    }

    #[test]
    fn everything_else_is_read_as_utf8() {
        for document in [
            "<r/>",
            "<?xml version=\"1.0\"?><r/>",
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?><r/>",
            "<?xml version=\"1.0\" encoding=\"windows-1252\"?><r/>",
            "<?xml version=\"1.0\" encoding=\"UTF-16\"?><r/>",
            "<?xml version=\"encoding='ISO-8859-1'\"?><r/>",
            "<?xml version=\"1.0\" encoding=ISO-8859-1?><r/>",
            "<?xml version=\"1.0\" encoding=\"ISO-8859-1?><r/>",
            "<?xml version=\"1.0\" encoding?><r/>",
            "<?xml?><r/>",
            "<?xmlencoding=\"ISO-8859-1\"?><r/>",
            "<?xml-stylesheet encoding=\"ISO-8859-1\"?><r/>",
            " <?xml version=\"1.0\" encoding=\"ISO-8859-1\"?><r/>",
            "<?pi encoding=\"ISO-8859-1\"?><r/>",
            "<!-- <?xml encoding=\"ISO-8859-1\"?> --><r/>",
            "\u{feff}<r/>",
            "é",
        ] {
            assert_eq!(
                detect(document.as_bytes()),
                Some((Encoding::Utf8, 0)),
                "{document}"
            );
        }
    }

    #[test]
    fn a_start_that_does_not_tell_asks_for_more() {
        let document = b"\xEF\xBB\xBF<?xml version=\"1.0\" encoding=\"ISO-8859-1\"?><r/>";
        let end = document.len() - "<r/>".len();
        for len in 0..end {
            assert_eq!(detect(&document[..len]), None, "{len}");
            assert_eq!(
                Encoding::of_document(&document[..len], true),
                Some((Encoding::Utf8, 0)),
                "{len}"
            );
        }
        assert_eq!(detect(&document[..end]), Some((Encoding::Latin1, 3)));
        assert_eq!(detect(b"\xEF\xBBx"), Some((Encoding::Utf8, 0)));
        assert_eq!(detect(b"<?xm>"), Some((Encoding::Utf8, 0)));
    }

    #[test]
    fn a_document_may_declare_the_encoding_it_is_in() {
        use Declared::{Differs, Fits, Unknown};
        for (encoding, declared, verdict) in [
            (Encoding::Utf8, "UTF-8", Fits),
            (Encoding::Utf8, "utf-8", Fits),
            (Encoding::Utf8, "US-ASCII", Fits),
            (Encoding::Utf8, "UTF-16", Differs),
            (Encoding::Utf8, "ISO-8859-1", Differs),
            (Encoding::Utf8, "utf8", Unknown),
            (Encoding::Utf8, "windows-1252", Unknown),
            (Encoding::Latin1, "iso-8859-1", Fits),
            (Encoding::Latin1, "UTF-8", Differs),
            (Encoding::Utf16Le, "UTF-16", Fits),
            (Encoding::Utf16Le, "utf-16le", Fits),
            (Encoding::Utf16Le, "UTF-16BE", Differs),
            (Encoding::Utf16Be, "UTF-16", Fits),
            (Encoding::Utf16Be, "UTF-16BE", Fits),
            (Encoding::Utf16Be, "UTF-16LE", Differs),
            (Encoding::Utf16Be, "UTF-8", Differs),
            (Encoding::Utf16Be, "latin1", Unknown),
        ] {
            assert_eq!(
                encoding.accepts(declared),
                verdict,
                "{encoding:?} {declared}"
            );
        }
    }
}
