//! Worksheet parts such as `xl/worksheets/sheet1.xml`: the rows and cells of a sheet.

use duckdb_neo::{Error, Result};

use crate::{
    error::{self, invalid_input},
    xlsx::{
        cell::{CellPos, CellType, MAX_CELL_SIZE, parse_integer},
        xml::{Attributes, Flow, XmlHandler},
    },
};

/// The most bytes a tag, a comment or a text between tags of a worksheet may have. Nothing a spreadsheet
/// program writes comes near it; it keeps a broken or hostile sheet from being held in memory as a whole.
pub const MAX_EVENT_LEN: usize = 16 * 1024 * 1024;

/// What a worksheet parser does with the rows and cells it finds, in document order.
pub trait SheetVisitor {
    fn begin_row(&mut self, _row: u64) -> Flow {
        Flow::Continue
    }

    fn end_row(&mut self, row: u64) -> Flow;

    /// Receives a cell with its text: the value of a number, the index of a shared string, the text of an
    /// inline string.
    fn cell(&mut self, pos: CellPos, cell_type: CellType, data: &str, style: u64) -> Result<()>;
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum State {
    Start,
    SheetData,
    Row,
    Cell,
    V,
    Is,
    T,
}

/// Walks the `sheetData` of a worksheet and stops after it.
///
/// Rows and cells without a reference follow the one before them. The text of a cell is what its `v` element
/// and the `t` elements of its inline string hold, formatting runs put together.
#[derive(Debug)]
pub struct WorksheetParser<V> {
    state: State,
    cell_pos: CellPos,
    cell_type: CellType,
    cell_data: String,
    cell_style: u64,
    pub visitor: V,
}

impl<V: SheetVisitor> WorksheetParser<V> {
    pub fn new(visitor: V) -> WorksheetParser<V> {
        WorksheetParser {
            state: State::Start,
            cell_pos: CellPos::new(0, 0),
            cell_type: CellType::Number,
            cell_data: String::new(),
            cell_style: 0,
            visitor,
        }
    }
}

fn cell_too_large() -> Error {
    invalid_input("XLSX: Cell data too large (is the file corrupted?)")
}

impl<V: SheetVisitor> XmlHandler for WorksheetParser<V> {
    fn start_element(&mut self, name: &str, attributes: &Attributes<'_>) -> Result<Flow> {
        match self.state {
            State::Start if name == "sheetData" => self.state = State::SheetData,
            State::SheetData if name == "row" => {
                self.state = State::Row;
                self.cell_pos.col = 0;
                self.cell_pos.row = match attributes.get(["r"])? {
                    [Some(row)] => parse_integer(&row) as u64,
                    [None] => self.cell_pos.row.wrapping_add(1),
                };
                return Ok(self.visitor.begin_row(self.cell_pos.row));
            }
            State::Row if name == "c" => {
                self.state = State::Cell;
                self.cell_data.clear();
                let [cell_type, reference, style] = attributes.get(["t", "r", "s"])?;
                self.cell_style = style.map_or(0, |style| parse_integer(&style) as u64);
                self.cell_type = cell_type.map_or(CellType::Number, |cell_type| {
                    CellType::from_attribute(&cell_type)
                });
                match reference {
                    None => self.cell_pos.col = self.cell_pos.col.wrapping_add(1),
                    Some(reference) => {
                        let mut pos = CellPos::new(1, 1);
                        if pos.parse_prefix(reference.as_bytes()).is_none() {
                            return Err(invalid_input(format!(
                                "Invalid cell reference in sheet: {reference}"
                            )));
                        }
                        if pos.row != self.cell_pos.row {
                            return Err(invalid_input(
                                "Cell reference does not match row reference in sheet",
                            ));
                        }
                        self.cell_pos.col = pos.col;
                    }
                }
            }
            State::Cell if name == "v" => self.state = State::V,
            State::Cell if name == "is" => self.state = State::Is,
            State::Is if name == "t" => self.state = State::T,
            _ => {}
        }
        Ok(Flow::Continue)
    }

    fn end_element(&mut self, name: &str) -> Result<Flow> {
        match self.state {
            State::SheetData if name == "sheetData" => return Ok(Flow::Stop),
            State::Row if name == "row" => {
                self.state = State::SheetData;
                return Ok(self.visitor.end_row(self.cell_pos.row));
            }
            State::Cell if name == "c" => {
                self.state = State::Row;
                self.visitor.cell(
                    self.cell_pos,
                    self.cell_type,
                    &self.cell_data,
                    self.cell_style,
                )?;
            }
            State::V if name == "v" => self.state = State::Cell,
            State::Is if name == "is" => self.state = State::Cell,
            State::T if name == "t" => self.state = State::Is,
            _ => {}
        }
        Ok(Flow::Continue)
    }

    fn collects_text(&self) -> bool {
        matches!(self.state, State::V | State::T)
    }

    fn text(&mut self, text: &str) -> Result<()> {
        if self.cell_data.len() + text.len() > MAX_CELL_SIZE * 2 {
            return Err(cell_too_large());
        }
        self.cell_data.push_str(text);
        Ok(())
    }

    fn oversized(&self, limit: usize) -> Error {
        if self.collects_text() {
            cell_too_large()
        } else {
            error::io(format!(
                "XML parse error: a single tag, comment or text of the sheet is longer than {limit} bytes"
            ))
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xlsx::xml::XmlReader;
    use duckdb_neo::error::DuckDBError;

    /// Writes down the rows and cells it is shown.
    #[derive(Default)]
    struct Recorder(Vec<String>);

    impl SheetVisitor for Recorder {
        fn begin_row(&mut self, row: u64) -> Flow {
            self.0.push(format!("row {row}"));
            Flow::Continue
        }

        fn end_row(&mut self, row: u64) -> Flow {
            self.0.push(format!("end {row}"));
            Flow::Continue
        }

        fn cell(
            &mut self,
            pos: CellPos,
            cell_type: CellType,
            data: &str,
            style: u64,
        ) -> Result<()> {
            self.0
                .push(format!("{pos} {cell_type:?} {data:?} s{style}"));
            Ok(())
        }
    }

    fn parse(sheet_data: &str) -> Result<Vec<String>> {
        let xml = format!("<worksheet><sheetData>{sheet_data}</sheetData></worksheet>");
        let mut parser = WorksheetParser::new(Recorder::default());
        XmlReader::new(xml.as_bytes(), Some(MAX_EVENT_LEN)).parse_all(&mut parser)?;
        Ok(parser.visitor.0)
    }

    fn error(sheet_data: &str) -> String {
        let error = parse(sheet_data).unwrap_err();
        assert_eq!(error.code, DuckDBError::InputInvalid);
        error.message
    }

    #[test]
    fn cells_with_references() {
        let events = parse(
            r#"<row r="2" spans="1:3"><c r="A2"><v>1.5</v></c><c r="C2" s="3" t="s"><v>7</v></c></row>
               <row r="5"><c r="B5" t="inlineStr"><is><t>text</t></is></c><c r="D5" t="b"><v>1</v></c></row>"#,
        );
        assert_eq!(
            events.unwrap(),
            [
                "row 2",
                "A2 Number \"1.5\" s0",
                "C2 SharedString \"7\" s3",
                "end 2",
                "row 5",
                "B5 InlineString \"text\" s0",
                "D5 Boolean \"1\" s0",
                "end 5"
            ]
        );
    }

    #[test]
    fn rows_and_cells_without_references_follow_the_one_before() {
        let events = parse(
            r#"<row><c><v>1</v></c><c><v>2</v></c></row><row r="7"><c r="C7"><v>3</v></c><c><v>4</v></c></row>
               <row><c><v>5</v></c></row>"#,
        );
        assert_eq!(
            events.unwrap(),
            [
                "row 1",
                "A1 Number \"1\" s0",
                "B1 Number \"2\" s0",
                "end 1",
                "row 7",
                "C7 Number \"3\" s0",
                "D7 Number \"4\" s0",
                "end 7",
                "row 8",
                "A8 Number \"5\" s0",
                "end 8"
            ]
        );
    }

    #[test]
    fn the_text_of_a_cell_is_its_value_and_its_inline_string() {
        let events = parse(
            r#"<row r="1"><c r="A1" t="inlineStr"><is><r><rPr><b/></rPr><t>bold</t></r><r><t> plain</t></r><rPh><t>!</t></rPh></is></c>
               <c r="B1"><f>SUM(A1)</f><v>3</v></c><c r="C1" t="str"><v> a &amp; b </v></c><c r="D1"/><c r="E1"><v/></c>
               <c r="F1"><v>1<x>2</x>3</v><v>4</v></c><c r="G1" t="inlineStr"><is><t/></is></c></row>"#,
        );
        assert_eq!(
            events.unwrap(),
            [
                "row 1",
                "A1 InlineString \"bold plain!\" s0",
                "B1 Number \"3\" s0",
                "C1 FormulaString \" a & b \" s0",
                "D1 Number \"\" s0",
                "E1 Number \"\" s0",
                "F1 Number \"1234\" s0",
                "G1 InlineString \"\" s0",
                "end 1"
            ]
        );
    }

    #[test]
    fn only_the_sheet_data_is_read() {
        let xml = r#"<x:worksheet><x:dimension ref="A1"/><row r="9"><c r="A9"><v>0</v></c></row>
            <x:sheetData><x:row r="1"><x:c r="A1"><x:v>1</x:v></x:c></x:row></x:sheetData><broken"#;
        let mut parser = WorksheetParser::new(Recorder::default());
        XmlReader::new(xml.as_bytes(), Some(MAX_EVENT_LEN))
            .parse_all(&mut parser)
            .unwrap();
        assert_eq!(parser.visitor.0, ["row 1", "A1 Number \"1\" s0", "end 1"]);
    }

    #[test]
    fn odd_row_numbers_and_styles_are_taken_as_written() {
        let events = parse(
            r#"<row r="0"><c s="-1"><v>1</v></c></row><row r=" 12abc"><c s="x"><v>2</v></c></row>"#,
        );
        assert_eq!(
            events.unwrap(),
            [
                "row 0",
                format!("A0 Number \"1\" s{}", u64::MAX).as_str(),
                "end 0",
                "row 12",
                "A12 Number \"2\" s0",
                "end 12"
            ]
        );
    }

    #[test]
    fn cell_references_must_be_in_their_row() {
        assert_eq!(
            error(r#"<row r="2"><c r="a2"><v>1</v></c></row>"#),
            "Invalid cell reference in sheet: a2"
        );
        assert_eq!(
            error(r#"<row r="2"><c r="A3"><v>1</v></c></row>"#),
            "Cell reference does not match row reference in sheet"
        );
        assert_eq!(
            error(r#"<row r="2"><c r="B"><v>1</v></c></row>"#),
            "Cell reference does not match row reference in sheet"
        );
        assert!(parse(r#"<row r="1"><c r="B"><v>1</v></c><c r="3"><v>1</v></c></row>"#).is_err());
        assert_eq!(
            parse(r#"<row r="1"><c r="B"><v>1</v></c><c r="1"><v>2</v></c><c r="C1junk"><v>3</v></c></row>"#).unwrap(),
            ["row 1", "B1 Number \"1\" s0", "A1 Number \"2\" s0", "C1 Number \"3\" s0", "end 1"]
        );
    }

    #[test]
    fn cell_text_has_a_limit() {
        let limit = MAX_CELL_SIZE * 2;
        let cell = |text: &str| {
            format!(r#"<row r="1"><c r="A1" t="inlineStr"><is><t>{text}</t></is></c></row>"#)
        };
        let events = parse(&cell(&"x".repeat(limit))).unwrap();
        assert_eq!(events.len(), 3);
        let too_large = "XLSX: Cell data too large (is the file corrupted?)";
        assert_eq!(error(&cell(&"x".repeat(limit + 1))), too_large);
        assert_eq!(error(&cell(&"x&amp;".repeat(limit / 2 + 1))), too_large);
        assert_eq!(
            error(&cell(&"x".repeat(MAX_EVENT_LEN + 100_000))),
            too_large
        );
    }

    #[test]
    fn markup_over_the_event_limit_is_an_error() {
        let comment = format!("<!--{}-->", "x".repeat(MAX_EVENT_LEN + 100_000));
        let error = parse(&comment).unwrap_err();
        assert_eq!(error.code, DuckDBError::IoGeneral);
        assert!(error.message.starts_with("XML parse error"));
    }
}
