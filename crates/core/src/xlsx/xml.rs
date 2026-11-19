//! Reads the XML parts of a workbook and hands them to a handler one element at a time.
//!
//! Handlers see what a non-validating SAX parser without namespace support would show them: element names
//! without their prefix, attribute names as written, and attribute values and character data with entities and
//! line ends resolved. Documents are in UTF-8, UTF-16 or ISO-8859-1; the latter two are converted to UTF-8 as
//! they are read. Entities a document declares itself are not known.

use std::{
    borrow::Cow,
    cell::Cell,
    fmt,
    io::{self, BufRead, Read},
};

use duckdb_neo::{Error, Result};
use quick_xml::{
    XmlVersion,
    events::{BytesStart, Event},
};

use crate::{
    error,
    xlsx::encoding::{Declared, Encoding, MalformedText, Transcoder},
};

const SOURCE_BUFFER_LEN: usize = 32 * 1024;

/// What a handler wants the reader to do after an element.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Flow {
    Continue,
    /// Return from [`XmlReader::parse`]; the next call goes on after this element. An empty element that
    /// pauses at its start is still shown its end first.
    Pause,
    /// Nothing after this element is of interest: the rest of the document is never read.
    Stop,
}

impl Flow {
    fn then(self, next: Flow) -> Flow {
        match (self, next) {
            (Flow::Stop, _) | (_, Flow::Stop) => Flow::Stop,
            (Flow::Pause, _) | (_, Flow::Pause) => Flow::Pause,
            (Flow::Continue, Flow::Continue) => Flow::Continue,
        }
    }
}

/// Why [`XmlReader::parse`] returned.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Parsed {
    Paused,
    /// The handler stopped the reader or the document ended.
    Done,
}

/// Receives the elements of a document in document order.
pub trait XmlHandler {
    /// `name` comes without its namespace prefix; an empty element is a start followed by an end.
    fn start_element(&mut self, name: &str, attributes: &Attributes<'_>) -> Result<Flow>;

    fn end_element(&mut self, name: &str) -> Result<Flow>;

    /// Whether [`Self::text`] wants the character data that comes next.
    fn collects_text(&self) -> bool {
        false
    }

    /// Receives character data in pieces, CDATA sections included, whitespace untouched.
    fn text(&mut self, _text: &str) -> Result<()> {
        Ok(())
    }

    /// The error for a tag, comment or run of text longer than the limit of the reader.
    fn oversized(&self, limit: usize) -> Error {
        error::io(format!(
            "XML parse error: a single tag, comment or text is longer than {limit} bytes"
        ))
    }
}

/// The attributes of a start tag.
#[derive(Debug)]
pub struct Attributes<'a> {
    tag: &'a BytesStart<'a>,
    position: u64,
    checked: Cell<bool>,
}

impl<'a> Attributes<'a> {
    const SMALL_TAG: usize = 8;

    /// The values of the attributes called `names`, unescaped and with whitespace normalized.
    /// Fails when any attribute of the tag is malformed, whether it was asked for or not.
    pub fn get<const N: usize>(&self, names: [&str; N]) -> Result<[Option<Cow<'a, str>>; N]> {
        if let Some(values) = self.get_plain(names) {
            self.checked.set(true);
            return Ok(values);
        }
        let mut values = [const { None }; N];
        self.for_each(|name, value| {
            if let Some(index) = names.iter().position(|wanted| *wanted == name) {
                values[index] = Some(value);
            }
        })?;
        Ok(values)
    }

    /// Every attribute of the tag in the order they are written, with its value as [`Self::get`] returns it.
    pub fn all(&self) -> Result<Vec<(&'a str, Cow<'a, str>)>> {
        let mut attributes = Vec::new();
        self.for_each(|name, value| attributes.push((name, value)))?;
        Ok(attributes)
    }

    fn for_each(&self, mut visit: impl FnMut(&'a str, Cow<'a, str>)) -> Result<()> {
        let mut seen = [""; Self::SMALL_TAG];
        let mut count = 0;
        // The duplicate check of quick-xml allocates for every tag, so tags with few attributes, which is
        // nearly all of them, are checked here.
        let mut attributes = self.tag.attributes();
        attributes.with_checks(false);
        for attribute in attributes {
            let attribute = attribute.map_err(|error| self.error(error))?;
            let name = attribute.key.0;
            if seen.iter().take(count).any(|known| *known == name) {
                return Err(self.error(format_args!("duplicate attribute '{name}'")));
            }
            if let Some(slot) = seen.get_mut(count) {
                *slot = name;
            }
            count += 1;
            if attribute.value.contains('<') {
                return Err(self.error("'<' in an attribute value"));
            }
            let value = attribute
                .normalized_value(XmlVersion::Implicit1_0)
                .map_err(|error| self.error(error))?;
            visit(name, value);
        }
        if count > Self::SMALL_TAG {
            for attribute in self.tag.attributes() {
                attribute.map_err(|error| self.error(error))?;
            }
        }
        self.checked.set(true);
        Ok(())
    }

    fn get_plain<const N: usize>(&self, names: [&str; N]) -> Option<[Option<Cow<'a, str>>; N]> {
        // Nearly every tag is written as `name="value"` pairs with nothing to unescape or normalize. Such a
        // tag is read here in one pass; `None` leaves anything else, errors included, to the full parser.
        let raw = self.tag.attributes_raw();
        let bytes = raw.as_bytes();
        let mut values = [const { None }; N];
        let mut seen = [""; Self::SMALL_TAG];
        let mut count = 0;
        let mut at = 0;
        loop {
            let name_start = at + run_len(bytes.get(at..)?, is_xml_whitespace);
            if name_start == bytes.len() {
                return Some(values);
            }
            if name_start == at && count > 0 {
                return None;
            }
            let name_end = name_start + run_len(bytes.get(name_start..)?, is_name_byte);
            let quote = match bytes.get(name_end..)? {
                [b'=', quote @ (b'"' | b'\''), ..] => *quote,
                _ => return None,
            };
            let value_start = name_end + 2;
            let is_plain =
                |byte: u8| byte != quote && !matches!(byte, b'<' | b'&' | b'\t' | b'\n' | b'\r');
            let value_end = value_start + run_len(bytes.get(value_start..)?, is_plain);
            if bytes.get(value_end) != Some(&quote) {
                return None;
            }
            let name = raw.get(name_start..name_end)?;
            if name.is_empty() || seen.iter().take(count).any(|known| *known == name) {
                return None;
            }
            *seen.get_mut(count)? = name;
            count += 1;
            if let Some(index) = names.iter().position(|wanted| *wanted == name) {
                values[index] = Some(Cow::Borrowed(raw.get(value_start..value_end)?));
            }
            at = value_end + 1;
        }
    }

    fn error(&self, error: impl fmt::Display) -> Error {
        parse_error(self.position, error)
    }
}

fn run_len(bytes: &[u8], belongs: impl Fn(u8) -> bool) -> usize {
    bytes.iter().take_while(|byte| belongs(**byte)).count()
}

fn parse_error(position: u64, error: impl fmt::Display) -> Error {
    error::io(format!("XML parse error at byte {position}: {error}"))
}

/// Marks the I/O error that reports markup or text over the limit of the source.
#[derive(Debug)]
struct Oversized;

impl fmt::Display for Oversized {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter.write_str("XML markup or text is too long")
    }
}

impl std::error::Error for Oversized {}

/// Where the bytes of a document come from once its encoding is known.
#[derive(Debug)]
enum Input {
    Unread,
    Utf8,
    Transcoded(Box<Transcoder>),
}

/// Buffers the document as UTF-8 and watches what quick-xml does not: it refuses to feed a single event more
/// than `limit` bytes, which keeps a document that never closes a tag or a text from being held in memory as a
/// whole, and it notes where the first control character that XML forbids is, which is cheaper to find in whole
/// buffers than in every event. For the same reason it notes whether the document has had a carriage return
/// or a `]` yet: without them no text has a line end to normalize or a `]]>` to refuse. `event_start` is where
/// in the document the event being read begins.
#[derive(Debug)]
struct Source<R> {
    inner: R,
    input: Input,
    encoding: Encoding,
    buffer: Box<[u8]>,
    start: usize,
    end: usize,
    bytes_read: u64,
    source_bytes_before: u64,
    first_control: Option<u64>,
    has_carriage_return: bool,
    has_closing_bracket: bool,
    event_start: u64,
    limit: usize,
}

impl<R: Read> Read for Source<R> {
    fn read(&mut self, target: &mut [u8]) -> io::Result<usize> {
        let available = self.fill_buf()?;
        let count = available.len().min(target.len());
        if let (Some(target), Some(source)) = (target.get_mut(..count), available.get(..count)) {
            target.copy_from_slice(source);
        }
        self.consume(count);
        Ok(count)
    }
}

impl<R: Read> Source<R> {
    #[cold]
    fn refill(&mut self) -> io::Result<()> {
        // Everything read so far has been taken, so this is how long the event has become.
        if self.bytes_read - self.event_start > self.limit as u64 {
            return Err(io::Error::other(Oversized));
        }
        let count = match &mut self.input {
            Input::Unread => None,
            Input::Utf8 => Some(read_some(&mut self.inner, &mut self.buffer)?),
            Input::Transcoded(transcoder) => {
                Some(transcoder.read(&mut self.inner, &mut self.buffer)?)
            }
        };
        let count = match count {
            Some(count) => count,
            None => self.read_start()?,
        };
        self.start = 0;
        self.end = count.min(self.buffer.len());
        let fresh = self.buffer.get(..self.end).unwrap_or_default();
        // Nearly every buffer has none of these bytes, which a single pass over it tells.
        if has_unusual_byte(fresh) {
            self.has_carriage_return |= fresh.contains(&b'\r');
            self.has_closing_bracket |= fresh.contains(&b']');
            if self.first_control.is_none() {
                self.first_control = fresh
                    .iter()
                    .position(|byte| is_control_character(*byte))
                    .map(|index| self.bytes_read + index as u64);
            }
        }
        self.bytes_read += self.end as u64;
        Ok(())
    }

    fn read_start(&mut self) -> io::Result<usize> {
        let mut filled = 0;
        let (encoding, skipped) = loop {
            let room = self
                .buffer
                .get_mut(filled..Encoding::DETECTION_LEN)
                .unwrap_or_default();
            let count = read_some(&mut self.inner, room)?;
            filled += count;
            let start = self.buffer.get(..filled).unwrap_or_default();
            if let Some(found) = Encoding::of_document(start, count == 0) {
                break found;
            }
        };
        self.encoding = encoding;
        let start = self.buffer.get(skipped..filled).unwrap_or_default();
        match Transcoder::new(encoding, start) {
            None => {
                self.input = Input::Utf8;
                Ok(filled)
            }
            Some(mut transcoder) => {
                let count = transcoder.read(&mut self.inner, &mut self.buffer)?;
                self.source_bytes_before = filled as u64;
                self.input = Input::Transcoded(Box::new(transcoder));
                Ok(count)
            }
        }
    }

    fn source_bytes_read(&self) -> u64 {
        match &self.input {
            Input::Transcoded(transcoder) => self.source_bytes_before + transcoder.raw_bytes_read(),
            Input::Unread | Input::Utf8 => self.bytes_read,
        }
    }
}

fn read_some<R: Read>(source: &mut R, target: &mut [u8]) -> io::Result<usize> {
    loop {
        match source.read(target) {
            Err(error) if error.kind() == io::ErrorKind::Interrupted => {}
            result => return Ok(result?.min(target.len())),
        }
    }
}

impl<R: Read> BufRead for Source<R> {
    #[inline]
    fn fill_buf(&mut self) -> io::Result<&[u8]> {
        // The parser asks several times for every event: refilling is kept apart so that this stays small
        // enough to be inlined into it.
        if self.start == self.end {
            self.refill()?;
        }
        Ok(self.buffer.get(self.start..self.end).unwrap_or_default())
    }

    #[inline]
    fn consume(&mut self, count: usize) {
        self.start = self.start.saturating_add(count).min(self.end);
    }
}

/// A pull parser: [`Self::parse`] runs until the handler pauses or stops it, and a paused reader carries on
/// where it was.
pub struct XmlReader<R: Read> {
    reader: quick_xml::Reader<Source<R>>,
    buffer: Vec<u8>,
    depth: usize,
    has_root: bool,
    at_start: bool,
    done: bool,
}

impl<R: Read> fmt::Debug for XmlReader<R> {
    fn fmt(&self, formatter: &mut fmt::Formatter<'_>) -> fmt::Result {
        formatter
            .debug_struct("XmlReader")
            .field("depth", &self.depth)
            .field("done", &self.done)
            .finish_non_exhaustive()
    }
}

impl<R: Read> XmlReader<R> {
    /// Reads the document in `source`; `limit` bounds the bytes of a single tag, comment or run of text.
    pub fn new(source: R, limit: Option<usize>) -> XmlReader<R> {
        let mut reader = quick_xml::Reader::from_reader(Source {
            inner: source,
            input: Input::Unread,
            encoding: Encoding::Utf8,
            buffer: vec![0; SOURCE_BUFFER_LEN].into_boxed_slice(),
            start: 0,
            end: 0,
            bytes_read: 0,
            source_bytes_before: 0,
            first_control: None,
            has_carriage_return: false,
            has_closing_bracket: false,
            event_start: 0,
            limit: limit.unwrap_or(usize::MAX),
        });
        reader.config_mut().check_comments = true;
        XmlReader {
            reader,
            buffer: Vec::new(),
            depth: 0,
            has_root: false,
            at_start: true,
            done: false,
        }
    }

    /// The number of bytes taken from the source so far.
    pub fn bytes_read(&self) -> u64 {
        self.reader.get_ref().source_bytes_read()
    }

    /// Feeds `handler` until it pauses or stops the reader, or until the document ends.
    /// After an error the reader is done.
    pub fn parse<H: XmlHandler>(&mut self, handler: &mut H) -> Result<Parsed> {
        if self.done {
            return Ok(Parsed::Done);
        }
        let parsed = self.parse_events(handler);
        self.done |= !matches!(parsed, Ok(Parsed::Paused));
        parsed
    }

    fn parse_events<H: XmlHandler>(&mut self, handler: &mut H) -> Result<Parsed> {
        let mut flow = Flow::Continue;
        loop {
            match flow {
                Flow::Continue => {}
                Flow::Pause => return Ok(Parsed::Paused),
                Flow::Stop => return Ok(Parsed::Done),
            }
            self.buffer.clear();
            self.reader.get_mut().event_start = self.reader.buffer_position();
            let event = match self.reader.read_event_into(&mut self.buffer) {
                Ok(event) => event,
                Err(error) => {
                    let position = self.reader.error_position();
                    let limit = self.reader.get_ref().limit;
                    return Err(read_error(error, position, limit, handler));
                }
            };
            let position = self.reader.buffer_position();
            let source = self.reader.get_ref();
            if source
                .first_control
                .is_some_and(|control| control < position)
            {
                return Err(parse_error(position, "invalid character"));
            }
            let (has_carriage_return, has_closing_bracket) =
                (source.has_carriage_return, source.has_closing_bracket);
            let at_start = std::mem::take(&mut self.at_start);
            flow = match event {
                Event::Start(tag) => {
                    let flow = start_element(handler, &tag, position, self.depth, self.has_root)?;
                    self.depth += 1;
                    self.has_root = true;
                    flow
                }
                Event::Empty(tag) => {
                    let flow = empty_element(handler, &tag, position, self.depth, self.has_root)?;
                    self.has_root = true;
                    flow
                }
                Event::End(tag) => {
                    self.depth = self.depth.saturating_sub(1);
                    handler.end_element(local_name(tag.name().0))?
                }
                Event::Text(text) => {
                    if self.depth == 0 {
                        if !text.bytes().all(is_xml_whitespace) {
                            return Err(parse_error(position, "text outside the document element"));
                        }
                    } else if has_closing_bracket
                        && text.as_bytes().contains(&b'>')
                        && text.contains("]]>")
                    {
                        return Err(parse_error(position, "']]>' in text"));
                    } else if handler.collects_text() {
                        if has_carriage_return {
                            handler.text(&text.xml10_content())?;
                        } else {
                            handler.text(&text)?;
                        }
                    }
                    Flow::Continue
                }
                Event::CData(text) => {
                    if self.depth == 0 {
                        return Err(parse_error(position, "text outside the document element"));
                    }
                    if handler.collects_text() {
                        handler.text(&text.xml10_content())?;
                    }
                    Flow::Continue
                }
                Event::GeneralRef(reference) => {
                    if self.depth == 0 {
                        return Err(parse_error(position, "text outside the document element"));
                    }
                    let mut utf8 = [0; 4];
                    let text = resolve_reference(&reference, &mut utf8)
                        .map_err(|error| parse_error(position, error))?;
                    if handler.collects_text() {
                        handler.text(text)?;
                    }
                    Flow::Continue
                }
                Event::Decl(declaration) => {
                    if !at_start {
                        return Err(parse_error(
                            position,
                            "XML declaration not at the start of the document",
                        ));
                    }
                    declaration
                        .version()
                        .map_err(|error| parse_error(position, error))?;
                    if let Some(declared) = declaration.encoding() {
                        let declared = declared.map_err(|error| parse_error(position, error))?;
                        match self.reader.get_ref().encoding.accepts(&declared) {
                            Declared::Fits => {}
                            Declared::Differs => {
                                return Err(parse_error(
                                    position,
                                    format_args!(
                                        "the document is not in the encoding '{declared}' it declares"
                                    ),
                                ));
                            }
                            Declared::Unknown => return Err(unsupported_encoding(&declared)),
                        }
                    }
                    Flow::Continue
                }
                Event::Eof => {
                    let empty = self.bytes_read() == 0;
                    if !empty && (self.depth > 0 || !self.has_root) {
                        return Err(parse_error(
                            position,
                            "the document ends before its last element is closed",
                        ));
                    }
                    return Ok(Parsed::Done);
                }
                Event::Comment(_) | Event::PI(_) | Event::DocType(_) => Flow::Continue,
            };
        }
    }

    /// Runs `handler` over the document until it stops the reader or the document ends.
    pub fn parse_all<H: XmlHandler>(&mut self, handler: &mut H) -> Result<()> {
        while self.parse(handler)? == Parsed::Paused {}
        Ok(())
    }
}

fn start_element<H: XmlHandler>(
    handler: &mut H,
    tag: &BytesStart<'_>,
    position: u64,
    depth: usize,
    has_root: bool,
) -> Result<Flow> {
    if depth == 0 && has_root {
        return Err(parse_error(position, "junk after document element"));
    }
    let name = tag.name().0;
    if !is_name(name) {
        return Err(parse_error(position, "invalid element name"));
    }
    let attributes = Attributes {
        tag,
        position,
        // A tag that is all name has no attribute to check.
        checked: Cell::new(tag.len() == name.len()),
    };
    let flow = handler.start_element(local_name(name), &attributes)?;
    if !attributes.checked.get() {
        attributes.get([])?;
    }
    Ok(flow)
}

fn empty_element<H: XmlHandler>(
    handler: &mut H,
    tag: &BytesStart<'_>,
    position: u64,
    depth: usize,
    has_root: bool,
) -> Result<Flow> {
    let start = start_element(handler, tag, position, depth, has_root)?;
    // A handler that pauses or stops at the start of an empty element has seen all of it.
    let end = handler.end_element(local_name(tag.name().0))?;
    Ok(start.then(end))
}

fn unsupported_encoding(encoding: &str) -> Error {
    error::io(format!(
        "XML parse error: unsupported encoding '{encoding}', only UTF-8, UTF-16 and ISO-8859-1 are supported"
    ))
}

fn read_error<H: XmlHandler>(
    error: quick_xml::Error,
    position: u64,
    limit: usize,
    handler: &H,
) -> Error {
    match &error {
        quick_xml::Error::Io(io_error) => {
            let inner = io_error.get_ref();
            if inner.is_some_and(|inner| inner.is::<Oversized>()) {
                handler.oversized(limit)
            } else if let Some(malformed) =
                inner.and_then(|inner| inner.downcast_ref::<MalformedText>())
            {
                parse_error(position, malformed)
            } else {
                error::from_io(io_error)
            }
        }
        _ => parse_error(position, error),
    }
}

fn local_name(name: &str) -> &str {
    name.bytes()
        .position(|byte| byte == b':')
        .and_then(|colon| name.get(colon + 1..))
        .unwrap_or(name)
}

fn is_xml_whitespace(byte: u8) -> bool {
    matches!(byte, b' ' | b'\t' | b'\n' | b'\r')
}

fn has_unusual_byte(bytes: &[u8]) -> bool {
    bytes
        .iter()
        .fold(false, |found, byte| found | is_unusual_byte(*byte))
}

fn is_unusual_byte(byte: u8) -> bool {
    ((byte < b' ') & (byte != b'\t') & (byte != b'\n')) | (byte == b']')
}

fn is_control_character(byte: u8) -> bool {
    (byte < b' ') & (byte != b'\t') & (byte != b'\n') & (byte != b'\r')
}

fn is_name_byte(byte: u8) -> bool {
    byte.is_ascii_alphanumeric() || matches!(byte, b'_' | b':' | b'-' | b'.') || !byte.is_ascii()
}

fn is_name(name: &str) -> bool {
    match name.as_bytes() {
        [first, rest @ ..] => {
            is_name_byte(*first)
                && !matches!(first, b'0'..=b'9' | b'-' | b'.')
                && rest.iter().copied().all(is_name_byte)
        }
        [] => false,
    }
}

fn resolve_reference<'a>(
    reference: &str,
    utf8: &'a mut [u8; 4],
) -> std::result::Result<&'a str, String> {
    let Some(number) = reference.strip_prefix('#') else {
        return match reference {
            "lt" => Ok("<"),
            "gt" => Ok(">"),
            "amp" => Ok("&"),
            "apos" => Ok("'"),
            "quot" => Ok("\""),
            _ => Err(format!("undefined entity '&{reference};'")),
        };
    };
    let code = match number.strip_prefix('x') {
        Some(hex) if hex.bytes().all(|byte| byte.is_ascii_hexdigit()) => {
            u32::from_str_radix(hex, 16).ok()
        }
        None if number.bytes().all(|byte| byte.is_ascii_digit()) => number.parse().ok(),
        _ => None,
    };
    let character = code
        .filter(|code| {
            matches!(code, 0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF)
        })
        .and_then(char::from_u32);
    match character {
        Some(character) => Ok(character.encode_utf8(utf8)),
        None => Err(format!(
            "reference to invalid character number '&{reference};'"
        )),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use duckdb_neo::error::DuckDBError;

    /// Writes down what it is shown and does what `flow` says after each element.
    #[derive(Default)]
    struct Recorder {
        events: Vec<String>,
        collect: bool,
        wanted: Vec<&'static str>,
        pause_at: Option<&'static str>,
        stop_at: Option<&'static str>,
    }

    impl Recorder {
        fn collecting() -> Recorder {
            Recorder {
                collect: true,
                ..Recorder::default()
            }
        }

        fn flow(&self, event: &str) -> Flow {
            if self.stop_at == Some(event) {
                Flow::Stop
            } else if self.pause_at == Some(event) {
                Flow::Pause
            } else {
                Flow::Continue
            }
        }
    }

    impl XmlHandler for Recorder {
        fn start_element(&mut self, name: &str, attributes: &Attributes<'_>) -> Result<Flow> {
            let mut event = format!("<{name}");
            if let [Some(first), second] = attributes.get(["a", "r:id"])? {
                event.push_str(&format!(" a={first:?}"));
                if let Some(second) = second {
                    event.push_str(&format!(" r:id={second:?}"));
                }
            }
            for wanted in &self.wanted {
                if let [Some(value)] = attributes.get([*wanted])? {
                    event.push_str(&format!(" {wanted}={value:?}"));
                }
            }
            let flow = self.flow(&event);
            self.events.push(event);
            Ok(flow)
        }

        fn end_element(&mut self, name: &str) -> Result<Flow> {
            let event = format!("</{name}");
            let flow = self.flow(&event);
            self.events.push(event);
            Ok(flow)
        }

        fn collects_text(&self) -> bool {
            self.collect
        }

        fn text(&mut self, text: &str) -> Result<()> {
            self.events.push(format!("{text:?}"));
            Ok(())
        }
    }

    fn record(xml: &str, mut recorder: Recorder) -> Result<Vec<String>> {
        XmlReader::new(xml.as_bytes(), None).parse_all(&mut recorder)?;
        Ok(recorder.events)
    }

    fn events(xml: &str) -> Vec<String> {
        record(xml, Recorder::collecting()).unwrap()
    }

    fn error(xml: &str) -> String {
        let error = record(xml, Recorder::collecting()).unwrap_err();
        assert_eq!(error.code, DuckDBError::IoGeneral, "{xml}");
        assert!(
            error.message.starts_with("XML parse error"),
            "{}",
            error.message
        );
        error.message
    }

    #[test]
    fn element_names_lose_their_prefix_and_attribute_names_do_not() {
        assert_eq!(
            events(r#"<x:a xmlns:x="u" a="1" r:id="rId1"><x:b a="2"/><c x:a="3"/></x:a>"#),
            [
                "<a a=\"1\" r:id=\"rId1\"",
                "<b a=\"2\"",
                "</b",
                "<c",
                "</c",
                "</a"
            ]
        );
    }

    #[test]
    fn attribute_values_are_unescaped_and_normalized() {
        assert_eq!(
            events("<r a='x &amp; &lt;y&gt; &quot;&apos;&#65;&#x42;'/>"),
            ["<r a=\"x & <y> \\\"'AB\"", "</r"]
        );
        assert_eq!(
            events("<r a=\"one\ttwo\nthree\r\nfour&#10;five\"/>"),
            ["<r a=\"one two three four\\nfive\"", "</r"]
        );
    }

    fn plain_values(tag: &str) -> Option<Vec<Option<String>>> {
        let tag = BytesStart::from_content(tag, 1);
        let attributes = Attributes {
            tag: &tag,
            position: 0,
            checked: Cell::new(false),
        };
        let values = attributes.get_plain(["r", "t"])?;
        Some(
            values
                .into_iter()
                .map(|value| value.map(|value| value.into_owned()))
                .collect(),
        )
    }

    #[test]
    fn tags_written_the_plain_way_are_read_in_one_pass() {
        let some = |value: &str| Some(value.to_owned());
        for (tag, values) in [
            ("c", [None, None]),
            ("c ", [None, None]),
            ("c r=\"A1\"", [some("A1"), None]),
            (
                "c t='s' x:y=\"\" r=\"B'2> é\" ",
                [some("B'2> é"), some("s")],
            ),
            ("c\n  r='\"'\tt=\"\"\r\n", [some("\""), some("")]),
        ] {
            assert_eq!(plain_values(tag), Some(values.to_vec()), "{tag:?}");
        }
    }

    #[test]
    fn other_tags_are_left_to_the_full_parser() {
        let many: String = (0..9).map(|index| format!(" a{index}=\"\"")).collect();
        for tag in [
            "c r = \"A1\"",
            "c r=\"A1\"t=\"s\"",
            "c r=\"a&amp;b\"",
            "c r=\"a\tb\"",
            "c r=\"a\nb\"",
            "c r=\"a\rb\"",
            "c r=\"<\"",
            "c r=\"1\" r=\"2\"",
            "c r=1",
            "c r",
            "c r=\"1",
            "c r=\"1'",
            "c =\"1\"",
            "c r$=\"1\"",
            &format!("c{many}"),
        ] {
            assert_eq!(plain_values(tag), None, "{tag:?}");
        }
    }

    #[test]
    fn attributes_read_the_same_however_the_tag_is_written() {
        for (plain, other) in [
            (r#"<r a="1" r:id="x"/>"#, r#"<r a = "1"  r:id= 'x' />"#),
            (r#"<r a="1" r:id="x y"/>"#, "<r a=\"&#49;\" r:id=\"x\ty\"/>"),
            (
                r#"<r a="é'" r:id=">"/>"#,
                r#"<r a="&#233;&apos;" r:id="&gt;"/>"#,
            ),
            (r#"<r r:id="2" a="1"/>"#, r#"<r r:id="2"a="1"/>"#),
            (
                r#"<r z="" a="" r:id=""/>"#,
                r#"<r a='' z='' b='' c='' d='' e='' f='' g='' r:id=''/>"#,
            ),
        ] {
            assert_eq!(events(plain), events(other), "{other}");
        }
    }

    #[test]
    fn text_keeps_whitespace_and_resolves_references() {
        assert_eq!(
            events("<r>  a &amp; b&#32;&#x41;<![CDATA[<c> & ]]>  </r>"),
            [
                "<r",
                "\"  a \"",
                "\"&\"",
                "\" b\"",
                "\" \"",
                "\"A\"",
                "\"<c> & \"",
                "\"  \"",
                "</r"
            ]
        );
    }

    #[test]
    fn line_ends_become_line_feeds() {
        assert_eq!(
            events("<r>a\r\nb\rc\nd</r>"),
            ["<r", "\"a\\nb\\nc\\nd\"", "</r"]
        );
        assert_eq!(
            events("<r><![CDATA[a\r\nb]]></r>"),
            ["<r", "\"a\\nb\"", "</r"]
        );
        assert_eq!(
            events("<r>a&#13;b</r>"),
            ["<r", "\"a\"", "\"\\r\"", "\"b\"", "</r"]
        );
    }

    #[test]
    fn line_ends_and_brackets_are_found_wherever_in_a_document_they_come() {
        let filler = "<a>text</a>".repeat(10_000);
        assert!(filler.len() > 3 * SOURCE_BUFFER_LEN);
        let late = format!("<r>{filler}<b>one\r\ntwo\rthree</b><c>] ]> ]]</c></r>");
        let found = events(&late);
        assert_eq!(
            found.get(found.len() - 7..).unwrap_or_default(),
            [
                "<b",
                "\"one\\ntwo\\nthree\"",
                "</b",
                "<c",
                "\"] ]> ]]\"",
                "</c",
                "</r"
            ]
        );
        let early = format!("<r><b>one\rtwo</b>{filler}<b>three\r\n</b></r>");
        let found = events(&early);
        assert_eq!(found.get(2).map(String::as_str), Some("\"one\\ntwo\""));
        assert_eq!(
            found.get(found.len() - 3).map(String::as_str),
            Some("\"three\\n\"")
        );
        for xml in [
            format!("<r>{filler}<b>a ]]> b</b></r>"),
            format!("<r><b>]</b>{filler}<b>a ]]> b</b></r>"),
            format!("<r><b>a ]]> b</b>{filler}</r>"),
        ] {
            assert!(error(&xml).ends_with("']]>' in text"));
        }
        let wide = utf16(&format!("<r>{filler}<b>one\r\ntwo</b></r>"), true);
        let found = record_bytes(&wide).unwrap();
        assert_eq!(
            found.get(found.len() - 3).map(String::as_str),
            Some("\"one\\ntwo\"")
        );
    }

    #[test]
    fn text_is_only_delivered_when_collected() {
        let quiet = record("<r>text<a>more</a></r>", Recorder::default()).unwrap();
        assert_eq!(quiet, ["<r", "<a", "</a", "</r"]);
    }

    #[test]
    fn comments_instructions_and_doctype_are_skipped() {
        assert_eq!(
            events(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?><!DOCTYPE r><!-- c --><r>a<!-- c -->b<?pi x?>c</r>"
            ),
            ["<r", "\"a\"", "\"b\"", "\"c\"", "</r"]
        );
    }

    #[test]
    fn a_paused_reader_carries_on_where_it_was() {
        let xml = "<r><a/><b>t</b></r>";
        for pause_at in ["<r", "<a", "</a", "<b", "</b", "</r"] {
            let mut recorder = Recorder {
                pause_at: Some(pause_at),
                ..Recorder::collecting()
            };
            let mut reader = XmlReader::new(xml.as_bytes(), None);
            assert_eq!(
                reader.parse(&mut recorder).unwrap(),
                Parsed::Paused,
                "{pause_at}"
            );
            let last = if pause_at == "<a" { "</a" } else { pause_at };
            assert_eq!(recorder.events.last().map(String::as_str), Some(last));
            assert_eq!(reader.parse(&mut recorder).unwrap(), Parsed::Done);
            assert_eq!(recorder.events, events(xml));
            assert_eq!(reader.parse(&mut recorder).unwrap(), Parsed::Done);
        }
    }

    #[test]
    fn an_empty_element_is_shown_its_end_whatever_its_start_says() {
        let xml = "<r><a/><b/></r>";
        let mut recorder = Recorder {
            pause_at: Some("<a"),
            stop_at: Some("</a"),
            ..Recorder::default()
        };
        let mut reader = XmlReader::new(xml.as_bytes(), None);
        assert_eq!(reader.parse(&mut recorder).unwrap(), Parsed::Done);
        assert_eq!(recorder.events, ["<r", "<a", "</a"]);

        let recorder = Recorder {
            stop_at: Some("<b"),
            ..Recorder::default()
        };
        assert_eq!(
            record(xml, recorder).unwrap(),
            ["<r", "<a", "</a", "<b", "</b"]
        );
    }

    #[test]
    fn a_stopped_reader_reads_no_further() {
        let tail = "<broken".repeat(50_000);
        let xml = format!("<r><a>1</a></r></oops>{tail}");
        let recorder = Recorder {
            stop_at: Some("</a"),
            ..Recorder::collecting()
        };
        assert_eq!(
            record(&xml, recorder).unwrap(),
            ["<r", "<a", "\"1\"", "</a"]
        );

        let mut recorder = Recorder {
            stop_at: Some("<a"),
            ..Recorder::default()
        };
        let mut reader = XmlReader::new(xml.as_bytes(), None);
        assert_eq!(reader.parse(&mut recorder).unwrap(), Parsed::Done);
        assert!(reader.bytes_read() <= SOURCE_BUFFER_LEN as u64);
        assert_eq!(reader.parse(&mut recorder).unwrap(), Parsed::Done);
        assert_eq!(recorder.events, ["<r", "<a"]);
    }

    #[test]
    fn an_empty_document_has_no_elements() {
        assert!(events("").is_empty());
    }

    #[test]
    fn documents_that_end_early_are_errors() {
        for xml in [
            "<r><a>",
            "<r>",
            "<r><a></a>",
            " \n",
            "<?xml version=\"1.0\"?>",
            "<r",
            "<r><!-- c",
        ] {
            error(xml);
        }
    }

    #[test]
    fn malformed_documents_are_errors() {
        for xml in [
            "<r></s>",
            "</r>",
            "<r><a></r></a>",
            "text<r/>",
            "<r/><s/>",
            "<r/>text",
            "<r a=\"1\" a=\"2\"/>",
            "<r a=1/>",
            "<r a=\"<\"/>",
            "<r a=\"&bogus;\"/>",
            "<r>&bogus;</r>",
            "<r>&#0;</r>",
            "<r>&#x1F;</r>",
            "<r>&#xD800;</r>",
            "<r>&#xFFFF;</r>",
            "<r>&#x110000;</r>",
            "<r>&#+65;</r>",
            "<r>&#X41;</r>",
            "<r>a & b</r>",
            "<r>a ]]> b</r>",
            "<r>a\u{1}b</r>",
            "<r><!-- \u{0} --></r>",
            "<r a=\"\u{8}\"/>",
            "<r><1a/></r>",
            "<r><></></r>",
            "<r><!-- a -- b --></r>",
            " <?xml version=\"1.0\"?><r/>",
            "<?xml encoding=\"UTF-8\"?><r/>",
        ] {
            error(xml);
        }
    }

    #[test]
    fn errors_after_the_stop_are_not_found() {
        let recorder = Recorder {
            stop_at: Some("</r"),
            ..Recorder::default()
        };
        assert_eq!(
            record("<r></r><s/>junk</x>", recorder).unwrap(),
            ["<r", "</r"]
        );
    }

    #[test]
    fn attributes_of_ignored_elements_are_checked() {
        struct Ignore;
        impl XmlHandler for Ignore {
            fn start_element(&mut self, _: &str, _: &Attributes<'_>) -> Result<Flow> {
                Ok(Flow::Continue)
            }
            fn end_element(&mut self, _: &str) -> Result<Flow> {
                Ok(Flow::Continue)
            }
        }
        let parse = |xml: &str| XmlReader::new(xml.as_bytes(), None).parse_all(&mut Ignore);
        assert!(parse("<r a=\"1\" b=\"2\"><s c='3'/></r>").is_ok());
        assert!(parse("<r a=\"1\" a=\"2\"/>").is_err());
        assert!(parse("<r><s a=x/></r>").is_err());
        let many: String = (0..40)
            .map(|index| format!(" a{index}=\"{index}\""))
            .collect();
        assert!(parse(&format!("<r{many}/>")).is_ok());
        assert!(parse(&format!("<r{many} a39=\"again\"/>")).is_err());
    }

    fn utf16(text: &str, is_big_endian: bool) -> Vec<u8> {
        text.encode_utf16()
            .flat_map(|unit| {
                if is_big_endian {
                    unit.to_be_bytes()
                } else {
                    unit.to_le_bytes()
                }
            })
            .collect()
    }

    fn latin1(text: &str) -> Vec<u8> {
        text.chars()
            .map(|character| u8::try_from(u32::from(character)).expect("a Latin-1 character"))
            .collect()
    }

    fn record_bytes(xml: &[u8]) -> Result<Vec<String>> {
        let mut recorder = Recorder::collecting();
        XmlReader::new(xml, None).parse_all(&mut recorder)?;
        Ok(recorder.events)
    }

    fn error_of_bytes(xml: &[u8]) -> String {
        let error = record_bytes(xml).unwrap_err();
        assert_eq!(error.code, DuckDBError::IoGeneral);
        assert!(
            error.message.starts_with("XML parse error"),
            "{}",
            error.message
        );
        error.message
    }

    #[test]
    fn utf8_documents_may_say_so() {
        for declaration in ["UTF-8", "utf-8", "US-ASCII"] {
            let xml = format!("<?xml version=\"1.0\" encoding=\"{declaration}\"?><r/>");
            assert_eq!(events(&xml), ["<r", "</r"]);
        }
        assert_eq!(
            events("\u{feff}<?xml version=\"1.0\"?><r>é</r>"),
            ["<r", "\"é\"", "</r"]
        );
    }

    #[test]
    fn utf16_documents_are_read() {
        let document = "<r a=\"ż\">é \u{1F600} 漢字<b r:id=\"ß\"/>&#228;&amp;</r>";
        let expected = events(document);
        assert_eq!(expected.len(), 7);
        for is_big_endian in [false, true] {
            for declaration in [
                "",
                "<?xml version=\"1.0\"?>",
                "<?xml version=\"1.0\" encoding=\"UTF-16\"?>",
                "<?xml version=\"1.0\" encoding=\"utf-16\" standalone=\"yes\"?>\r\n",
                if is_big_endian {
                    "<?xml version=\"1.0\" encoding=\"UTF-16BE\"?>"
                } else {
                    "<?xml version=\"1.0\" encoding=\"UTF-16LE\"?>"
                },
            ] {
                for mark in ["", "\u{feff}"] {
                    let bytes = utf16(&format!("{mark}{declaration}{document}"), is_big_endian);
                    assert_eq!(
                        record_bytes(&bytes).unwrap(),
                        expected,
                        "{is_big_endian} {declaration:?} {mark:?}"
                    );
                }
            }
        }
    }

    #[test]
    fn latin1_documents_are_read() {
        let document = "<r a=\"\u{e4}\">caf\u{e9} \u{a0}\u{ff}\u{80}<b/></r>";
        let expected = events(document);
        for declaration in [
            "<?xml version=\"1.0\" encoding=\"ISO-8859-1\"?>",
            "<?xml version='1.0' encoding='iso-8859-1' standalone='yes'?>\n",
        ] {
            let bytes = latin1(&format!("{declaration}{document}"));
            assert_eq!(record_bytes(&bytes).unwrap(), expected, "{declaration}");
            let with_mark = [b"\xEF\xBB\xBF".as_slice(), &bytes].concat();
            assert_eq!(record_bytes(&with_mark).unwrap(), expected, "{declaration}");
        }
        // Without the declaration the bytes are not UTF-8.
        let bytes = latin1(document);
        assert!(record_bytes(&bytes).is_err());
    }

    #[test]
    fn long_documents_in_other_encodings_are_read() {
        let item = "<i n=\"\u{e9}\">\u{e4}\u{f6}\u{fc} text</i>";
        let document = format!("<r>{}</r>", item.repeat(20_000));
        let expected = events(&document);
        assert_eq!(expected.len(), 2 + 3 * 20_000);
        let declared = format!("<?xml version=\"1.0\" encoding=\"ISO-8859-1\"?>{document}");
        assert_eq!(record_bytes(&latin1(&declared)).unwrap(), expected);
        for is_big_endian in [false, true] {
            let bytes = utf16(&format!("\u{feff}{document}"), is_big_endian);
            let mut recorder = Recorder::collecting();
            let mut reader = XmlReader::new(bytes.as_slice(), None);
            reader.parse_all(&mut recorder).unwrap();
            assert_eq!(recorder.events, expected);
            assert_eq!(reader.bytes_read(), bytes.len() as u64);
        }
    }

    #[test]
    fn a_document_must_be_in_the_encoding_it_declares() {
        for declared in ["ISO-8859-1", "UTF-8", "US-ASCII", "UTF-16LE"] {
            let xml = format!("<?xml version=\"1.0\" encoding=\"{declared}\"?><r/>");
            let message = error_of_bytes(&utf16(&xml, true));
            assert!(
                message.ends_with(&format!(
                    "the document is not in the encoding '{declared}' it declares"
                )),
                "{message}"
            );
        }
        let xml = "<?xml version=\"1.0\" encoding=\"UTF-16BE\"?><r/>";
        assert!(error_of_bytes(&utf16(xml, false)).contains("it declares"));
        for declared in ["UTF-16", "UTF-16LE", "utf-16be"] {
            let xml = format!("<?xml version=\"1.0\" encoding=\"{declared}\"?><r/>");
            assert!(error(&xml).contains("it declares"), "{declared}");
        }
    }

    #[test]
    fn other_encodings_are_refused() {
        for declared in ["windows-1252", "utf8", "latin1", "ISO-8859-15", "UCS-2", ""] {
            let xml = format!("<?xml version=\"1.0\" encoding=\"{declared}\"?><r/>");
            for bytes in [xml.as_bytes().to_vec(), utf16(&xml, false)] {
                let message = error_of_bytes(&bytes);
                assert!(
                    message.contains(&format!("unsupported encoding '{declared}'")),
                    "{message}"
                );
            }
        }
    }

    #[test]
    fn text_that_is_not_in_its_encoding_is_an_error() {
        let half_pair = [utf16("<r>", false), vec![0x3D, 0xD8], utf16("</r>", false)].concat();
        assert!(error_of_bytes(&half_pair).ends_with("half a surrogate pair in UTF-16 text"));
        let odd = [utf16("<r/>", true), vec![0x20]].concat();
        assert!(error_of_bytes(&odd).ends_with("the document ends in the middle of a character"));
        let control = utf16("<r>\u{1}</r>", true);
        assert!(error_of_bytes(&control).ends_with("invalid character"));
        let not_utf8 = b"<r>caf\xe9</r>";
        error_of_bytes(not_utf8);
        // A document that starts with a NUL byte is taken for UTF-16.
        error_of_bytes(b"\0\0\0<r/>");
    }

    #[test]
    fn all_attributes_come_in_the_order_they_are_written() {
        struct Collect(Vec<Vec<(String, String)>>);
        impl XmlHandler for Collect {
            fn start_element(&mut self, _: &str, attributes: &Attributes<'_>) -> Result<Flow> {
                let all = attributes.all()?;
                let owned =
                    |(name, value): &(&str, Cow<'_, str>)| ((*name).to_owned(), value.to_string());
                self.0.push(all.iter().map(owned).collect());
                Ok(Flow::Continue)
            }
            fn end_element(&mut self, _: &str) -> Result<Flow> {
                Ok(Flow::Continue)
            }
        }
        let xml = "<r xmlns:a=\"u\" b='x &amp; y' a:c=\"1\n2\"><s/><t z=\"\" y=\"\"/></r>";
        let mut collect = Collect(Vec::new());
        XmlReader::new(xml.as_bytes(), None)
            .parse_all(&mut collect)
            .unwrap();
        let pairs = |pairs: &[(&str, &str)]| -> Vec<(String, String)> {
            pairs
                .iter()
                .map(|(name, value)| ((*name).to_owned(), (*value).to_owned()))
                .collect()
        };
        assert_eq!(
            collect.0,
            [
                pairs(&[("xmlns:a", "u"), ("b", "x & y"), ("a:c", "1 2")]),
                pairs(&[]),
                pairs(&[("z", ""), ("y", "")])
            ]
        );
        let mut collect = Collect(Vec::new());
        let result =
            XmlReader::new("<r a=\"1\" a=\"2\"/>".as_bytes(), None).parse_all(&mut collect);
        assert!(result.is_err());
    }

    #[test]
    fn events_over_the_limit_are_refused() {
        let limit = 100_000;
        let long_text = format!("<r>{}</r>", "x".repeat(4 * limit));
        let long_comment = format!("<r><!--{}--></r>", "x".repeat(4 * limit));
        let long_tag = format!("<r a=\"{}\"/>", "x".repeat(4 * limit));
        for xml in [long_text, long_comment, long_tag] {
            let result =
                XmlReader::new(xml.as_bytes(), Some(limit)).parse_all(&mut Recorder::default());
            assert!(
                result
                    .unwrap_err()
                    .message
                    .contains("longer than 100000 bytes")
            );
            assert!(
                XmlReader::new(xml.as_bytes(), None)
                    .parse_all(&mut Recorder::default())
                    .is_ok()
            );
        }
        let many_short = format!("<r>{}</r>", "<a>x</a>".repeat(limit));
        let result =
            XmlReader::new(many_short.as_bytes(), Some(limit)).parse_all(&mut Recorder::default());
        assert!(result.is_ok());
    }

    #[test]
    fn errors_of_the_source_come_through_unchanged() {
        struct Failing;
        impl Read for Failing {
            fn read(&mut self, _: &mut [u8]) -> io::Result<usize> {
                Err(error::to_io(error::binder("the file is gone")))
            }
        }
        let failure = XmlReader::new(Failing, None)
            .parse_all(&mut Recorder::default())
            .unwrap_err();
        assert_eq!(failure.code, DuckDBError::QueryBinder);
        assert_eq!(failure.message, "the file is gone");
    }

    #[test]
    fn errors_of_the_handler_come_through_unchanged() {
        struct Refusing;
        impl XmlHandler for Refusing {
            fn start_element(&mut self, _: &str, _: &Attributes<'_>) -> Result<Flow> {
                Err(error::invalid_input("no"))
            }
            fn end_element(&mut self, _: &str) -> Result<Flow> {
                Ok(Flow::Continue)
            }
        }
        let failure = XmlReader::new("<r/>".as_bytes(), None)
            .parse_all(&mut Refusing)
            .unwrap_err();
        assert_eq!(failure.code, DuckDBError::InputInvalid);
        assert_eq!(failure.message, "no");
    }
}
