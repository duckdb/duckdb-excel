//! Reads the XML parts of a workbook and hands them to a handler one element at a time.
//!
//! Handlers see what a non-validating SAX parser without namespace support would show them: element names
//! without their prefix, attribute names as written, and attribute values and character data with entities and
//! line ends resolved. Only UTF-8 documents are read, and entities a document declares itself are not known.

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

use crate::error;

const SOURCE_BUFFER_LEN: usize = 32 * 1024;

/// What a handler wants the reader to do after an element.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Flow {
    Continue,
    /// Return from [`XmlReader::parse`]; the next call goes on after this element.
    Pause,
    /// Nothing after this element is of interest: the rest of the document is never read.
    Stop,
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
        let mut values = [const { None }; N];
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
            if let Some(index) = names.iter().position(|wanted| *wanted == name) {
                values[index] = Some(value);
            }
        }
        if count > Self::SMALL_TAG {
            for attribute in self.tag.attributes() {
                attribute.map_err(|error| self.error(error))?;
            }
        }
        self.checked.set(true);
        Ok(values)
    }

    fn error(&self, error: impl fmt::Display) -> Error {
        parse_error(self.position, error)
    }
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

/// Buffers the document and watches what quick-xml does not: it refuses to feed a single event more than
/// `limit` bytes, which keeps a document that never closes a tag or a text from being held in memory as a whole,
/// and it notes where the first control character that XML forbids is, which is cheaper to find in whole
/// buffers than in every event.
#[derive(Debug)]
struct Source<R> {
    inner: R,
    buffer: Box<[u8]>,
    start: usize,
    end: usize,
    bytes_read: u64,
    consumed: u64,
    first_control: Option<u64>,
    event_len: usize,
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

impl<R: Read> BufRead for Source<R> {
    fn fill_buf(&mut self) -> io::Result<&[u8]> {
        if self.start == self.end {
            if self.event_len > self.limit {
                return Err(io::Error::other(Oversized));
            }
            let count = loop {
                match self.inner.read(&mut self.buffer) {
                    Err(error) if error.kind() == io::ErrorKind::Interrupted => {}
                    result => break result?,
                }
            };
            self.start = 0;
            self.end = count.min(self.buffer.len());
            let fresh = self.buffer.get(..self.end).unwrap_or_default();
            let has_control = fresh
                .iter()
                .fold(false, |found, byte| found | is_control_character(*byte));
            if has_control && self.first_control.is_none() {
                self.first_control = fresh
                    .iter()
                    .position(|byte| is_control_character(*byte))
                    .map(|index| self.bytes_read + index as u64);
            }
            self.bytes_read += self.end as u64;
        }
        Ok(self.buffer.get(self.start..self.end).unwrap_or_default())
    }

    fn consume(&mut self, count: usize) {
        let count = count.min(self.end - self.start);
        self.start += count;
        self.consumed += count as u64;
        self.event_len = self.event_len.saturating_add(count);
    }
}

/// A pull parser: [`Self::parse`] runs until the handler pauses or stops it, and a paused reader carries on
/// where it was.
pub struct XmlReader<R: Read> {
    reader: quick_xml::Reader<Source<R>>,
    buffer: Vec<u8>,
    depth: usize,
    has_root: bool,
    started: bool,
    at_start: bool,
    pending_end: Option<String>,
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
            buffer: vec![0; SOURCE_BUFFER_LEN].into_boxed_slice(),
            start: 0,
            end: 0,
            bytes_read: 0,
            consumed: 0,
            first_control: None,
            event_len: 0,
            limit: limit.unwrap_or(usize::MAX),
        });
        reader.config_mut().check_comments = true;
        XmlReader {
            reader,
            buffer: Vec::new(),
            depth: 0,
            has_root: false,
            started: false,
            at_start: true,
            pending_end: None,
            done: false,
        }
    }

    /// The number of bytes taken from the source so far.
    pub fn bytes_read(&self) -> u64 {
        self.reader.get_ref().bytes_read
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
        if !self.started {
            self.started = true;
            self.reject_utf16()?;
        }
        let mut flow = match self.pending_end.take() {
            Some(name) => handler.end_element(&name)?,
            None => Flow::Continue,
        };
        loop {
            match flow {
                Flow::Continue => {}
                Flow::Pause => return Ok(Parsed::Paused),
                Flow::Stop => return Ok(Parsed::Done),
            }
            self.buffer.clear();
            self.reader.get_mut().event_len = 0;
            let event = match self.reader.read_event_into(&mut self.buffer) {
                Ok(event) => event,
                Err(error) => {
                    let position = self.reader.error_position();
                    let limit = self.reader.get_ref().limit;
                    return Err(read_error(error, position, limit, handler));
                }
            };
            let source = self.reader.get_ref();
            let position = source.consumed;
            if source
                .first_control
                .is_some_and(|control| control < position)
            {
                return Err(parse_error(position, "invalid character"));
            }
            let at_start = std::mem::take(&mut self.at_start);
            flow = match event {
                Event::Start(tag) => {
                    let flow = start_element(handler, &tag, position, self.depth, self.has_root)?;
                    self.depth += 1;
                    self.has_root = true;
                    flow
                }
                Event::Empty(tag) => {
                    let flow = start_element(handler, &tag, position, self.depth, self.has_root)?;
                    self.has_root = true;
                    let name = local_name(tag.name().0);
                    match flow {
                        Flow::Continue => handler.end_element(name)?,
                        Flow::Pause => {
                            // The end of the element is delivered when the reader carries on.
                            self.pending_end = Some(name.to_owned());
                            Flow::Pause
                        }
                        Flow::Stop => Flow::Stop,
                    }
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
                    } else if text.as_bytes().contains(&b'>') && text.contains("]]>") {
                        return Err(parse_error(position, "']]>' in text"));
                    } else if handler.collects_text() {
                        handler.text(&text.xml10_content())?;
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
                    if let Some(encoding) = declaration.encoding() {
                        let encoding = encoding.map_err(|error| parse_error(position, error))?;
                        if !["utf-8", "us-ascii"]
                            .iter()
                            .any(|supported| encoding.eq_ignore_ascii_case(supported))
                        {
                            return Err(unsupported_encoding(&encoding));
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

    fn reject_utf16(&mut self) -> Result<()> {
        let start = self
            .reader
            .get_mut()
            .fill_buf()
            .map_err(|error| error::from_io(&error))?;
        match start {
            [0xFF, 0xFE, ..] | [0xFE, 0xFF, ..] | [b'<', 0, ..] | [0, b'<', ..] => {
                Err(unsupported_encoding("UTF-16"))
            }
            _ => Ok(()),
        }
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

fn unsupported_encoding(encoding: &str) -> Error {
    error::io(format!(
        "XML parse error: unsupported encoding '{encoding}', only UTF-8 is supported"
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
            let oversized = io_error
                .get_ref()
                .is_some_and(|inner| inner.is::<Oversized>());
            if oversized {
                handler.oversized(limit)
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

fn is_control_character(byte: u8) -> bool {
    (byte < b' ') & (byte != b'\t') & (byte != b'\n') & (byte != b'\r')
}

fn is_name(name: &str) -> bool {
    let is_name_byte = |byte: u8| {
        byte.is_ascii_alphanumeric()
            || matches!(byte, b'_' | b':' | b'-' | b'.')
            || !byte.is_ascii()
    };
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
            assert_eq!(recorder.events.last().map(String::as_str), Some(pause_at));
            assert_eq!(reader.parse(&mut recorder).unwrap(), Parsed::Done);
            assert_eq!(recorder.events, events(xml));
            assert_eq!(reader.parse(&mut recorder).unwrap(), Parsed::Done);
        }
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

    #[test]
    fn only_utf8_is_read() {
        for declaration in ["UTF-8", "utf-8", "US-ASCII"] {
            let xml = format!("<?xml version=\"1.0\" encoding=\"{declaration}\"?><r/>");
            assert_eq!(events(&xml), ["<r", "</r"]);
        }
        assert_eq!(
            events("\u{feff}<?xml version=\"1.0\"?><r>é</r>"),
            ["<r", "\"é\"", "</r"]
        );
        for declaration in ["ISO-8859-1", "UTF-16", "windows-1252", "utf8"] {
            let xml = format!("<?xml version=\"1.0\" encoding=\"{declaration}\"?><r/>");
            assert!(
                error(&xml).contains("unsupported encoding"),
                "{declaration}"
            );
        }
        let utf16: Vec<u8> = "\u{feff}<r/>"
            .encode_utf16()
            .flat_map(u16::to_le_bytes)
            .collect();
        let utf16_be: Vec<u8> = "<r/>".encode_utf16().flat_map(u16::to_be_bytes).collect();
        for bytes in [utf16, utf16_be] {
            let result = XmlReader::new(bytes.as_slice(), None).parse_all(&mut Recorder::default());
            assert!(
                result
                    .unwrap_err()
                    .message
                    .contains("unsupported encoding 'UTF-16'")
            );
        }
        let latin1 = b"<r>caf\xe9</r>";
        let result = XmlReader::new(latin1.as_slice(), None).parse_all(&mut Recorder::collecting());
        assert!(result.unwrap_err().message.starts_with("XML parse error"));
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
