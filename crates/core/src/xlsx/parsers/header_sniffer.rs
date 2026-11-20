//! Finds the header row of a range and the row that gives its columns their types.

use duckdb_neo::Result;

use crate::{
    error::invalid_input,
    xlsx::{
        cell::{CellPos, CellRange, CellType},
        options::HeaderMode,
        parsers::worksheet::SheetVisitor,
        xml::Flow,
    },
};

/// A cell of one of the first rows of a range.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SniffedCell {
    pub cell_type: CellType,
    pub pos: CellPos,
    pub data: String,
    pub style: u64,
}

/// Reads the first row of the range and, when that row is the header, the row after it.
///
/// Both rows come out with one cell per column of the range: cells a row leaves out are empty cells of the
/// default type. Without a header row the columns are named after the cells of the first row.
#[derive(Debug)]
pub struct HeaderSniffer {
    pub header_cells: Vec<SniffedCell>,
    pub column_cells: Vec<SniffedCell>,
    /// The range of the rows after the header.
    pub range: CellRange,
    header_mode: HeaderMode,
    last_col: u64,
    first_row: bool,
    absolute_range: bool,
    default_cell_type: CellType,
}

impl HeaderSniffer {
    /// `absolute_range` says the caller chose the range, which names columns `A`, `B`, ... instead of after
    /// their first cell.
    pub fn new(
        range: CellRange,
        header_mode: HeaderMode,
        absolute_range: bool,
        default_cell_type: CellType,
    ) -> HeaderSniffer {
        HeaderSniffer {
            header_cells: Vec::new(),
            column_cells: Vec::new(),
            range,
            header_mode,
            last_col: 0,
            first_row: true,
            absolute_range,
            default_cell_type,
        }
    }

    fn pad_until(&mut self, row: u64, end_col: u64) {
        for col in self.last_col.wrapping_add(1)..end_col {
            self.column_cells.push(SniffedCell {
                cell_type: self.default_cell_type,
                pos: CellPos::new(row, col),
                data: String::new(),
                style: 0,
            });
        }
    }

    fn is_header(&self) -> bool {
        match self.header_mode {
            HeaderMode::Never => false,
            HeaderMode::Force => true,
            HeaderMode::Maybe => self.column_cells.iter().all(|cell| {
                matches!(
                    cell.cell_type,
                    CellType::SharedString | CellType::InlineString
                ) && !cell.data.is_empty()
            }),
        }
    }
}

impl SheetVisitor for HeaderSniffer {
    fn begin_row(&mut self, row: u64) -> Flow {
        if self.range.contains_row(row) {
            self.column_cells.clear();
            self.last_col = self.range.beg.col.wrapping_sub(1);
        }
        Flow::Continue
    }

    fn end_row(&mut self, row: u64) -> Flow {
        if !self.range.contains_row(row) {
            self.column_cells.clear();
            self.last_col = self.range.beg.col.wrapping_sub(1);
            return Flow::Continue;
        }
        self.pad_until(row, self.range.end.col);
        if !self.first_row {
            return Flow::Stop;
        }
        if !self.is_header() {
            self.header_cells = self.column_cells.clone();
            for cell in &mut self.header_cells {
                cell.cell_type = CellType::InlineString;
                cell.style = 0;
                cell.data = if self.absolute_range {
                    cell.pos.column_name()
                } else {
                    cell.pos.to_string()
                };
            }
            return Flow::Stop;
        }
        self.header_cells = std::mem::take(&mut self.column_cells);
        self.last_col = self.range.beg.col.wrapping_sub(1);
        self.first_row = false;
        self.range.beg.row = row.wrapping_add(1);
        Flow::Continue
    }

    fn cell(&mut self, pos: CellPos, cell_type: CellType, data: &str, style: u64) -> Result<()> {
        if !self.range.contains(pos) {
            return Ok(());
        }
        // A cell that goes back in its row would give the range more columns than it is wide.
        if pos.col <= self.last_col {
            return Err(invalid_input(format!(
                "Cell '{pos}' does not follow the cell before it in its row (is the file corrupted?)"
            )));
        }
        self.pad_until(pos.row, pos.col);
        self.column_cells.push(SniffedCell {
            cell_type,
            pos,
            data: data.to_owned(),
            style,
        });
        self.last_col = pos.col;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xlsx::{parsers::worksheet::WorksheetParser, xml::XmlReader};

    const SHEET: &str = r#"<row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="inlineStr"><is><t>name</t></is></c><c r="C1" t="s"><v>1</v></c></row>
        <row r="2"><c r="A2"><v>1</v></c><c r="C2" s="2"><v>2</v></c></row>
        <row r="3"><c r="A3" t="inlineStr"><is><t>x</t></is></c></row>"#;

    fn sniff(
        sheet_data: &str,
        range: CellRange,
        mode: HeaderMode,
        absolute: bool,
    ) -> HeaderSniffer {
        let xml = format!("<worksheet><sheetData>{sheet_data}</sheetData></worksheet>");
        let mut parser =
            WorksheetParser::new(HeaderSniffer::new(range, mode, absolute, CellType::Number));
        XmlReader::new(xml.as_bytes(), None)
            .parse_all(&mut parser)
            .unwrap();
        parser.visitor
    }

    fn range(beg_row: u64, beg_col: u64, end_row: u64, end_col: u64) -> CellRange {
        CellRange {
            beg: CellPos::new(beg_row, beg_col),
            end: CellPos::new(end_row, end_col),
        }
    }

    fn texts(cells: &[SniffedCell]) -> Vec<String> {
        cells
            .iter()
            .map(|cell| {
                format!(
                    "{} {:?} {:?} s{}",
                    cell.pos, cell.cell_type, cell.data, cell.style
                )
            })
            .collect()
    }

    #[test]
    fn a_row_of_strings_is_the_header() {
        let sniffer = sniff(SHEET, range(1, 1, u64::MAX, 4), HeaderMode::Maybe, false);
        assert_eq!(
            texts(&sniffer.header_cells),
            [
                "A1 SharedString \"0\" s0",
                "B1 InlineString \"name\" s0",
                "C1 SharedString \"1\" s0"
            ]
        );
        assert_eq!(
            texts(&sniffer.column_cells),
            [
                "A2 Number \"1\" s0",
                "B2 Number \"\" s0",
                "C2 Number \"2\" s2"
            ]
        );
        assert_eq!(sniffer.range, range(2, 1, u64::MAX, 4));
    }

    #[test]
    fn a_row_with_other_cells_is_data() {
        let sniffer = sniff(SHEET, range(2, 1, u64::MAX, 4), HeaderMode::Maybe, false);
        assert_eq!(
            texts(&sniffer.header_cells),
            [
                "A2 InlineString \"A2\" s0",
                "B2 InlineString \"B2\" s0",
                "C2 InlineString \"C2\" s0"
            ]
        );
        assert_eq!(
            texts(&sniffer.column_cells),
            [
                "A2 Number \"1\" s0",
                "B2 Number \"\" s0",
                "C2 Number \"2\" s2"
            ]
        );
        assert_eq!(sniffer.range, range(2, 1, u64::MAX, 4));
    }

    #[test]
    fn the_header_can_be_forced_or_forbidden() {
        let forced = sniff(SHEET, range(2, 1, u64::MAX, 4), HeaderMode::Force, false);
        assert_eq!(
            texts(&forced.header_cells),
            [
                "A2 Number \"1\" s0",
                "B2 Number \"\" s0",
                "C2 Number \"2\" s2"
            ]
        );
        assert_eq!(
            texts(&forced.column_cells),
            [
                "A3 InlineString \"x\" s0",
                "B3 Number \"\" s0",
                "C3 Number \"\" s0"
            ]
        );
        assert_eq!(forced.range.beg.row, 3);

        let never = sniff(SHEET, range(1, 1, u64::MAX, 4), HeaderMode::Never, false);
        assert_eq!(
            texts(&never.header_cells),
            [
                "A1 InlineString \"A1\" s0",
                "B1 InlineString \"B1\" s0",
                "C1 InlineString \"C1\" s0"
            ]
        );
        assert_eq!(texts(&never.column_cells)[0], "A1 SharedString \"0\" s0");
        assert_eq!(never.range.beg.row, 1);
    }

    #[test]
    fn columns_of_a_chosen_range_are_named_by_their_letters() {
        let sniffer = sniff(SHEET, range(2, 2, 4, 5), HeaderMode::Maybe, true);
        let names: Vec<&str> = sniffer
            .header_cells
            .iter()
            .map(|cell| cell.data.as_str())
            .collect();
        assert_eq!(names, ["B", "C", "D"]);
        assert_eq!(
            texts(&sniffer.column_cells),
            [
                "B2 Number \"\" s0",
                "C2 Number \"2\" s2",
                "D2 Number \"\" s0"
            ]
        );
    }

    #[test]
    fn a_header_without_rows_below_it_has_no_column_cells() {
        let sniffer = sniff(SHEET, range(1, 1, 2, 4), HeaderMode::Maybe, true);
        assert_eq!(sniffer.header_cells.len(), 3);
        assert!(sniffer.column_cells.is_empty());
        let sniffer = sniff(SHEET, range(7, 1, 9, 3), HeaderMode::Maybe, true);
        assert!(sniffer.header_cells.is_empty() && sniffer.column_cells.is_empty());
    }

    #[test]
    fn the_row_below_the_header_may_come_after_a_gap() {
        let sheet = r#"<row r="1"><c r="A1" t="inlineStr"><is><t>h</t></is></c></row><row r="6"><c r="A6" t="b"><v>1</v></c></row>"#;
        let sniffer = sniff(sheet, range(1, 1, u64::MAX, 2), HeaderMode::Maybe, false);
        assert_eq!(texts(&sniffer.column_cells), ["A6 Boolean \"1\" s0"]);
        assert_eq!(sniffer.range.beg.row, 2);
    }

    #[test]
    fn cells_left_out_get_the_default_type() {
        let xml = r#"<worksheet><sheetData><row r="1"><c r="B1"><v>1</v></c></row></sheetData></worksheet>"#;
        let sniffer = HeaderSniffer::new(
            range(1, 1, 3, 4),
            HeaderMode::Never,
            true,
            CellType::InlineString,
        );
        let mut parser = WorksheetParser::new(sniffer);
        XmlReader::new(xml.as_bytes(), None)
            .parse_all(&mut parser)
            .unwrap();
        assert_eq!(
            texts(&parser.visitor.column_cells),
            [
                "A1 InlineString \"\" s0",
                "B1 Number \"1\" s0",
                "C1 InlineString \"\" s0"
            ]
        );
    }

    #[test]
    fn cells_that_go_back_in_a_sniffed_row_are_errors() {
        let sheet = r#"<row r="1"><c r="B1"><v>1</v></c><c r="A1"><v>2</v></c></row>"#;
        let xml = format!("<worksheet><sheetData>{sheet}</sheetData></worksheet>");
        for (range, fails) in [
            (range(1, 1, 3, 4), true),
            (range(2, 1, 3, 4), false),
            (range(1, 2, 3, 4), false),
        ] {
            let sniffer = HeaderSniffer::new(range, HeaderMode::Maybe, true, CellType::Number);
            let result =
                XmlReader::new(xml.as_bytes(), None).parse_all(&mut WorksheetParser::new(sniffer));
            assert_eq!(result.is_err(), fails, "{range:?}");
        }
    }
}
