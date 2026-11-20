//! `xl/sharedStrings.xml`: the strings that cells of type `s` refer to by index.

use std::collections::HashMap;

use duckdb_neo::Result;

use crate::xlsx::{
    cell::parse_integer,
    escape::decode_escaped_units,
    strings::StringTable,
    xml::{Attributes, Flow, XmlHandler},
};

/// What a shared strings parser does with the strings it reads.
pub trait StringSink {
    /// Receives the `uniqueCount` the part states, which need not be the number of strings that follow.
    fn unique_count(&mut self, _count: usize) {}

    /// Receives the next string; strings come in index order.
    fn string(&mut self, string: &str) -> Flow;
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum State {
    Start,
    Sst,
    Si,
    T,
    Rph,
}

/// Puts the text runs of each string item together, each with its `_xHHHH_` escapes resolved, and leaves out
/// phonetic text.
#[derive(Debug)]
pub struct SharedStringsParser<S> {
    state: State,
    data: String,
    run_start: usize,
    pub sink: S,
}

impl<S: StringSink> SharedStringsParser<S> {
    pub fn new(sink: S) -> SharedStringsParser<S> {
        SharedStringsParser {
            state: State::Start,
            data: String::new(),
            run_start: 0,
            sink,
        }
    }
}

impl<S: StringSink> XmlHandler for SharedStringsParser<S> {
    fn start_element(&mut self, name: &str, attributes: &Attributes<'_>) -> Result<Flow> {
        match self.state {
            State::Start if name == "sst" => {
                self.state = State::Sst;
                if let [Some(count)] = attributes.get(["uniqueCount"])? {
                    self.sink
                        .unique_count(usize::try_from(parse_integer(&count)).unwrap_or(0));
                }
            }
            State::Sst if name == "si" => self.state = State::Si,
            State::Si if name == "t" => {
                self.state = State::T;
                self.run_start = self.data.len();
            }
            State::Si if name == "rPh" => self.state = State::Rph,
            _ => {}
        }
        Ok(Flow::Continue)
    }

    fn end_element(&mut self, name: &str) -> Result<Flow> {
        match self.state {
            State::T if name == "t" => {
                self.state = State::Si;
                decode_escaped_units(&mut self.data, self.run_start);
            }
            State::Si if name == "si" => {
                self.state = State::Sst;
                let flow = self.sink.string(&self.data);
                self.data.clear();
                return Ok(flow);
            }
            State::Sst if name == "sst" => return Ok(Flow::Stop),
            State::Rph if name == "rPh" => self.state = State::Si,
            _ => {}
        }
        Ok(Flow::Continue)
    }

    fn collects_text(&self) -> bool {
        self.state == State::T
    }

    fn text(&mut self, text: &str) -> Result<()> {
        self.data.push_str(text);
        Ok(())
    }
}

impl StringSink for &mut StringTable {
    fn unique_count(&mut self, count: usize) {
        self.reserve(count);
    }

    fn string(&mut self, string: &str) -> Flow {
        self.push(string);
        Flow::Continue
    }
}

/// Keeps only the strings with the indexes it was asked for and stops reading once it has them all.
#[derive(Debug)]
pub struct StringSearch {
    wanted: Vec<usize>,
    next_wanted: usize,
    next_string: usize,
    pub found: HashMap<usize, String>,
}

impl StringSearch {
    pub fn new(mut wanted: Vec<usize>) -> StringSearch {
        wanted.sort_unstable();
        wanted.dedup();
        StringSearch {
            wanted,
            next_wanted: 0,
            next_string: 0,
            found: HashMap::new(),
        }
    }
}

impl StringSink for StringSearch {
    fn string(&mut self, string: &str) -> Flow {
        let Some(wanted) = self.wanted.get(self.next_wanted) else {
            return Flow::Stop;
        };
        if *wanted == self.next_string {
            self.found.insert(*wanted, string.to_owned());
            self.next_wanted += 1;
        }
        self.next_string += 1;
        Flow::Continue
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xlsx::xml::XmlReader;

    const STRINGS: &str = r#"<sst xmlns="m" count="9" uniqueCount="5">
        <si><t>plain</t></si>
        <si><r><rPr><b/></rPr><t>rich </t></r><r><t xml:space="preserve">text</t></r></si>
        <si><t>漢字</t><rPh sb="0" eb="2"><t>カンジ</t></rPh><phoneticPr fontId="1"/></si>
        <si><t> a &amp; b &lt;c&gt; </t></si>
        <si><t/></si>
        <si/>
    </sst>"#;

    fn table(xml: &str) -> Vec<String> {
        let mut table = StringTable::new();
        XmlReader::new(xml.as_bytes(), None)
            .parse_all(&mut SharedStringsParser::new(&mut table))
            .unwrap();
        (0..table.len())
            .map(|index| table.get(index).unwrap().to_owned())
            .collect()
    }

    #[test]
    fn strings_are_put_together_without_phonetic_text() {
        assert_eq!(
            table(STRINGS),
            ["plain", "rich text", "漢字", " a & b <c> ", "", ""]
        );
    }

    #[test]
    fn escapes_are_resolved_in_every_run_on_its_own() {
        let xml = r#"<sst><si><t>line_x000D__x000A_break</t></si>
            <si><r><t>tab_x0009_</t></r><r><t>_x005F_x0041_ _xD83D__xDE00_</t></r></si>
            <si><r><t>_x00</t></r><r><t>41_</t></r></si>
            <si><t>a_x0041_<!-- c -->_x0042_&amp;_x0043</t><t>_</t></si>
            <si><t>x</t><rPh><t>_x0041_</t></rPh></si>
            <si><t>&#95;x0041_ _x0041&#95; _x00&#52;1_</t></si></sst>"#;
        assert_eq!(
            table(xml),
            [
                "line\r\nbreak",
                "tab\t_x0041_ \u{1F600}",
                "_x0041_",
                "aAB&_x0043_",
                "x",
                "A A A"
            ]
        );
    }

    #[test]
    fn other_documents_have_no_strings() {
        assert!(table("<other><si><t>x</t></si></other>").is_empty());
        assert!(table("").is_empty());
        assert_eq!(
            table("<sst uniqueCount=\"-4\"><si><t>x</t></si></sst>"),
            ["x"]
        );
    }

    #[test]
    fn a_search_keeps_the_wanted_strings_and_stops_after_them() {
        let broken_tail = format!("{}<si><t>unread", STRINGS.replace("</sst>", ""));
        let mut parser = SharedStringsParser::new(StringSearch::new(vec![3, 1, 3]));
        XmlReader::new(broken_tail.as_bytes(), None)
            .parse_all(&mut parser)
            .unwrap();
        let mut found: Vec<_> = parser.sink.found.into_iter().collect();
        found.sort();
        assert_eq!(
            found,
            [(1, "rich text".to_owned()), (3, " a & b <c> ".to_owned())]
        );
    }

    #[test]
    fn a_search_for_missing_strings_finds_the_others() {
        let mut parser = SharedStringsParser::new(StringSearch::new(vec![0, 99]));
        XmlReader::new(STRINGS.as_bytes(), None)
            .parse_all(&mut parser)
            .unwrap();
        assert_eq!(parser.sink.found.len(), 1);
        assert_eq!(parser.sink.found.get(&0).map(String::as_str), Some("plain"));
    }
}
