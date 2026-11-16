//! Finds the cells of a sheet to read when the caller gave no range.

use duckdb_neo::Result;

use crate::xlsx::{
    cell::{CellPos, CellRange, CellType},
    parsers::worksheet::SheetVisitor,
    xml::Flow,
};

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum State {
    Empty,
    Found,
    Ended,
}

/// Looks for the first row that has a cell with content. The range starts at that row and spans the columns
/// from the first cell with content to the last one before an empty cell; it is open at the bottom.
#[derive(Debug)]
pub struct RangeSniffer {
    state: State,
    beg_col: u64,
    end_col: u64,
    beg_row: u64,
}

impl RangeSniffer {
    pub fn new() -> RangeSniffer {
        RangeSniffer {
            state: State::Empty,
            beg_col: 0,
            end_col: 0,
            beg_row: 0,
        }
    }

    /// The range found, or the whole sheet when no row has content.
    pub fn range(&self) -> CellRange {
        if self.beg_row == 0 {
            return CellRange::WHOLE_SHEET;
        }
        CellRange {
            beg: CellPos::new(self.beg_row, self.beg_col),
            end: CellPos::new(u64::MAX, self.end_col.wrapping_add(1)),
        }
    }
}

impl Default for RangeSniffer {
    fn default() -> Self {
        RangeSniffer::new()
    }
}

impl SheetVisitor for RangeSniffer {
    fn end_row(&mut self, row: u64) -> Flow {
        match self.state {
            State::Found | State::Ended => {
                self.beg_row = row;
                Flow::Stop
            }
            State::Empty => {
                self.beg_col = 0;
                self.end_col = 0;
                Flow::Continue
            }
        }
    }

    fn cell(&mut self, pos: CellPos, _cell_type: CellType, data: &str, _style: u64) -> Result<()> {
        match self.state {
            State::Empty if !data.is_empty() => {
                self.state = State::Found;
                self.beg_col = pos.col;
                self.end_col = pos.col;
            }
            State::Found if data.is_empty() => self.state = State::Ended,
            State::Found => self.end_col = pos.col,
            State::Empty | State::Ended => {}
        }
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xlsx::{parsers::worksheet::WorksheetParser, xml::XmlReader};

    fn sniff(sheet_data: &str) -> CellRange {
        let xml = format!("<worksheet><sheetData>{sheet_data}</sheetData></worksheet>");
        let mut parser = WorksheetParser::new(RangeSniffer::new());
        XmlReader::new(xml.as_bytes(), None)
            .parse_all(&mut parser)
            .unwrap();
        parser.visitor.range()
    }

    fn range(beg_row: u64, beg_col: u64, end_col: u64) -> CellRange {
        CellRange {
            beg: CellPos::new(beg_row, beg_col),
            end: CellPos::new(u64::MAX, end_col),
        }
    }

    #[test]
    fn the_first_row_with_content_gives_the_range() {
        let sheet = r#"<row r="1"><c r="A1"/><c r="B1"><v></v></c></row><row r="2"/>
            <row r="4"><c r="B4"/><c r="C4"><v>1</v></c><c r="D4"><v>2</v></c><c r="F4"><v>3</v></c></row>
            <row r="5"><c r="A5"><v>9</v></c></row>"#;
        assert_eq!(sniff(sheet), range(4, 3, 7));
    }

    #[test]
    fn the_range_ends_before_the_first_empty_cell() {
        let sheet = r#"<row r="3"><c r="B3"><v>1</v></c><c r="C3"><v>2</v></c><c r="D3"/><c r="E3"><v>3</v></c></row>"#;
        assert_eq!(sniff(sheet), range(3, 2, 4));
    }

    #[test]
    fn a_sheet_without_content_is_read_whole() {
        assert_eq!(sniff(""), CellRange::WHOLE_SHEET);
        assert_eq!(
            sniff(r#"<row r="1"><c r="A1"/></row><row r="2"/>"#),
            CellRange::WHOLE_SHEET
        );
        assert_eq!(
            sniff(r#"<row r="0"><c><v>1</v></c></row>"#),
            CellRange::WHOLE_SHEET
        );
    }

    #[test]
    fn rows_after_the_first_with_content_are_not_read() {
        let sheet = r#"<row r="1"><c r="A1"><v>1</v></c></row><row r="2"><c r="broken"/></row>"#;
        assert_eq!(sniff(sheet), range(1, 1, 2));
    }
}
