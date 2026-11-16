//! Collects the rows of a range in batches for the scan.

use duckdb_neo::Result;

use crate::{
    error::invalid_input,
    xlsx::{
        cell::{CellPos, CellRange, CellType, parse_integer},
        parsers::worksheet::SheetVisitor,
        xml::Flow,
    },
};

/// The value of a cell of a [`RowBatch`].
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CellValue<'a> {
    Null,
    Text(&'a str),
    /// The index of a shared string.
    Shared(usize),
}

/// A cell of a row; `Open` is a cell the row being read has not come to yet.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum Slot {
    Open,
    Null,
    Text { start: usize, len: u32 },
    Shared(usize),
}

/// Rows of cells as text, one cell per column of the range.
#[derive(Debug)]
pub struct RowBatch {
    width: usize,
    slots: Vec<Slot>,
    text: String,
    sheet_rows: Vec<u64>,
}

impl RowBatch {
    fn new(width: usize) -> RowBatch {
        RowBatch {
            width,
            slots: Vec::new(),
            text: String::new(),
            sheet_rows: Vec::new(),
        }
    }

    /// The number of rows.
    pub fn len(&self) -> usize {
        self.sheet_rows.len()
    }

    pub fn is_empty(&self) -> bool {
        self.sheet_rows.is_empty()
    }

    /// The number of the sheet row that batch row `row` holds.
    pub fn sheet_row(&self, row: usize) -> Option<u64> {
        self.sheet_rows.get(row).copied()
    }

    pub fn cell(&self, row: usize, col: usize) -> CellValue<'_> {
        let slot = (col < self.width)
            .then(|| self.slots.get(row * self.width + col))
            .flatten();
        match slot {
            Some(Slot::Text { start, len }) => self
                .text
                .get(*start..*start + *len as usize)
                .map_or(CellValue::Null, CellValue::Text),
            Some(Slot::Shared(index)) => CellValue::Shared(*index),
            Some(Slot::Open | Slot::Null) | None => CellValue::Null,
        }
    }

    pub fn clear(&mut self) {
        self.slots.clear();
        self.text.clear();
        self.sheet_rows.clear();
    }

    fn push_row(&mut self, sheet_row: u64, slots: &[Slot]) {
        self.slots.extend_from_slice(slots);
        self.sheet_rows.push(sheet_row);
    }

    fn push_null_row(&mut self, sheet_row: u64) {
        self.slots
            .extend(std::iter::repeat_n(Slot::Null, self.width));
        self.sheet_rows.push(sheet_row);
    }
}

/// Gathers the cells of the rows of a range into a [`RowBatch`] and pauses the parser when the batch is full
/// or when the sheet leaves out rows, which [`Self::skip_rows`] then fills in.
///
/// A cell a row leaves out is NULL, and so is a cell that comes after a later cell of its row.
#[derive(Debug)]
pub struct RowCollector {
    range: CellRange,
    stop_at_empty: bool,
    string_count: usize,
    batch_rows: usize,
    pub batch: RowBatch,
    current: Vec<Slot>,
    last_col: u64,
    last_row: u64,
    current_row: u64,
    is_row_empty: bool,
}

impl RowCollector {
    /// Collects up to `batch_rows` rows of `range` at a time; cells may refer to `string_count` shared strings.
    pub fn new(
        range: CellRange,
        stop_at_empty: bool,
        string_count: usize,
        batch_rows: usize,
    ) -> RowCollector {
        let width = usize::try_from(range.width()).unwrap_or(usize::MAX);
        RowCollector {
            range,
            stop_at_empty,
            string_count,
            batch_rows,
            batch: RowBatch::new(width),
            current: vec![Slot::Open; width],
            last_col: range.beg.col.wrapping_sub(1),
            last_row: range.beg.row.wrapping_sub(1),
            current_row: range.beg.row,
            is_row_empty: false,
        }
    }

    pub fn is_batch_full(&self) -> bool {
        self.batch.len() >= self.batch_rows
    }

    /// Whether the sheet left out rows before the row the parser paused at.
    pub fn found_skipped_row(&self) -> bool {
        self.last_row.wrapping_add(1) < self.current_row
    }

    /// Adds the rows the sheet left out as rows of NULLs, as far as the batch has room.
    pub fn skip_rows(&mut self) {
        while self.found_skipped_row() && !self.is_batch_full() {
            self.last_row += 1;
            self.batch.push_null_row(self.last_row);
        }
    }

    /// Adds rows of NULLs up to the end of the range, as far as the batch has room.
    pub fn fill_rows(&mut self) {
        let remaining = self
            .range
            .end
            .row
            .saturating_sub(1)
            .saturating_sub(self.last_row);
        let room = self.batch_rows.saturating_sub(self.batch.len());
        for _ in 0..remaining.min(room as u64) {
            self.last_row += 1;
            self.batch.push_null_row(self.last_row);
        }
    }

    fn slot_mut(&mut self, col: u64) -> Option<&mut Slot> {
        let index = usize::try_from(col.checked_sub(self.range.beg.col)?).ok()?;
        self.current.get_mut(index)
    }

    fn set_null_between(&mut self, first_col: u64, end_col: u64) {
        for col in first_col..end_col {
            if let Some(slot) = self.slot_mut(col) {
                *slot = Slot::Null;
            }
        }
    }

    fn set_value(&mut self, col: u64, value: Slot) {
        // A cell that was passed over stays NULL even when the row comes back to it.
        if let Some(slot) = self.slot_mut(col).filter(|slot| **slot != Slot::Null) {
            *slot = value;
        }
    }
}

impl SheetVisitor for RowCollector {
    fn begin_row(&mut self, row: u64) -> Flow {
        if !self.range.contains_row(row) {
            return Flow::Continue;
        }
        self.last_col = self.range.beg.col.wrapping_sub(1);
        self.is_row_empty = true;
        self.current_row = row;
        if self.found_skipped_row() {
            Flow::Pause
        } else {
            Flow::Continue
        }
    }

    fn end_row(&mut self, row: u64) -> Flow {
        if !self.range.contains_row(row) {
            return Flow::Continue;
        }
        self.last_row = row;
        if self.stop_at_empty && self.is_row_empty {
            return Flow::Stop;
        }
        self.set_null_between(self.last_col.wrapping_add(1), self.range.end.col);
        self.batch.push_row(row, &self.current);
        self.current.fill(Slot::Open);
        if self.is_batch_full() {
            Flow::Pause
        } else {
            Flow::Continue
        }
    }

    fn cell(&mut self, pos: CellPos, cell_type: CellType, data: &str, _style: u64) -> Result<()> {
        if !self.range.contains(pos) {
            return Ok(());
        }
        self.set_null_between(self.last_col.wrapping_add(1), pos.col);
        if cell_type == CellType::SharedString {
            let index = usize::try_from(parse_integer(data))
                .ok()
                .filter(|index| *index < self.string_count)
                .ok_or_else(|| {
                    invalid_input(format!(
                        "Cell '{pos}' refers to shared string '{data}', which the xlsx file does not have"
                    ))
                })?;
            self.set_value(pos.col, Slot::Shared(index));
            self.is_row_empty = false;
        } else if data.is_empty() && cell_type != CellType::InlineString {
            self.set_null_between(pos.col, pos.col.wrapping_add(1));
        } else {
            let start = self.batch.text.len();
            let len = u32::try_from(data.len())
                .map_err(|_| invalid_input("XLSX: Cell data too large (is the file corrupted?)"))?;
            self.batch.text.push_str(data);
            self.set_value(pos.col, Slot::Text { start, len });
            self.is_row_empty &= data.is_empty();
        }
        self.last_col = pos.col;
        Ok(())
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::xlsx::{
        parsers::worksheet::WorksheetParser,
        xml::{Parsed, XmlReader},
    };
    use duckdb_neo::error::DuckDBError;

    const BATCH_ROWS: usize = 4;

    fn range(beg_row: u64, beg_col: u64, end_row: u64, end_col: u64) -> CellRange {
        CellRange {
            beg: CellPos::new(beg_row, beg_col),
            end: CellPos::new(end_row, end_col),
        }
    }

    fn batch_rows(batch: &RowBatch, width: usize) -> Vec<String> {
        (0..batch.len())
            .map(|row| {
                let cells: Vec<String> = (0..width)
                    .map(|col| match batch.cell(row, col) {
                        CellValue::Null => "NULL".to_owned(),
                        CellValue::Text(text) => format!("{text:?}"),
                        CellValue::Shared(index) => format!("#{index}"),
                    })
                    .collect();
                format!("{}: {}", batch.sheet_row(row).unwrap(), cells.join(" "))
            })
            .collect()
    }

    fn collect(
        sheet_data: &str,
        range: CellRange,
        stop_at_empty: bool,
        fill: bool,
    ) -> Result<Vec<Vec<String>>> {
        let xml = format!("<worksheet><sheetData>{sheet_data}</sheetData></worksheet>");
        let width = range.width() as usize;
        let mut reader = XmlReader::new(xml.as_bytes(), None);
        let mut parser =
            WorksheetParser::new(RowCollector::new(range, stop_at_empty, 3, BATCH_ROWS));
        let mut status = Parsed::Paused;
        let mut batches = Vec::new();
        loop {
            parser.visitor.batch.clear();
            while !parser.visitor.is_batch_full() && status == Parsed::Paused {
                if parser.visitor.found_skipped_row() {
                    if stop_at_empty {
                        status = Parsed::Done;
                    } else {
                        parser.visitor.skip_rows();
                    }
                } else {
                    status = reader.parse(&mut parser)?;
                }
            }
            if fill {
                parser.visitor.fill_rows();
            }
            if parser.visitor.batch.is_empty() {
                return Ok(batches);
            }
            batches.push(batch_rows(&parser.visitor.batch, width));
        }
    }

    fn number_rows(rows: impl IntoIterator<Item = u64>) -> String {
        rows.into_iter()
            .map(|row| format!(r#"<row r="{row}"><c r="A{row}"><v>{row}</v></c></row>"#))
            .collect()
    }

    #[test]
    fn rows_come_in_batches() {
        let batches = collect(&number_rows(1..=9), range(1, 1, u64::MAX, 2), true, false).unwrap();
        assert_eq!(
            batches,
            [
                vec!["1: \"1\"", "2: \"2\"", "3: \"3\"", "4: \"4\""],
                vec!["5: \"5\"", "6: \"6\"", "7: \"7\"", "8: \"8\""],
                vec!["9: \"9\""]
            ]
        );
    }

    #[test]
    fn cells_outside_the_range_are_left_out_and_missing_cells_are_null() {
        let sheet = r#"<row r="1"><c r="A1"><v>a1</v></c><c r="B1"><v>b1</v></c><c r="D1"><v>d1</v></c><c r="E1"><v>e1</v></c></row>
            <row r="2"><c r="C2" t="s"><v>2</v></c></row><row r="3"><c r="B3" t="inlineStr"><is><t></t></is></c><c r="C3"><v></v></c></row>
            <row r="4"><c r="B4"><v>b4</v></c></row>"#;
        let batches = collect(sheet, range(1, 2, 4, 5), false, false).unwrap();
        assert_eq!(
            batches,
            [vec![
                "1: \"b1\" NULL \"d1\"",
                "2: NULL #2 NULL",
                "3: \"\" NULL NULL"
            ]]
        );
    }

    #[test]
    fn an_empty_row_ends_the_scan_when_asked_to() {
        let sheet = format!(
            "{}<row r=\"3\"><c r=\"A3\"/></row>{}",
            number_rows(1..=2),
            number_rows(4..=5)
        );
        let stopped = collect(&sheet, range(1, 1, u64::MAX, 2), true, false).unwrap();
        assert_eq!(stopped, [vec!["1: \"1\"", "2: \"2\""]]);
        let all = collect(&sheet, range(1, 1, u64::MAX, 2), false, false).unwrap();
        assert_eq!(
            all,
            [
                vec!["1: \"1\"", "2: \"2\"", "3: NULL", "4: \"4\""],
                vec!["5: \"5\""]
            ]
        );
    }

    #[test]
    fn a_shared_string_makes_a_row_non_empty() {
        let sheet = r#"<row r="1"><c r="A1" t="s"><v></v></c></row><row r="2"><c r="A2" t="inlineStr"><is><t></t></is></c></row><row r="3"><c r="A3"><v>3</v></c></row>"#;
        let batches = collect(sheet, range(1, 1, u64::MAX, 2), true, false).unwrap();
        assert_eq!(batches, [vec!["1: #0"]]);
    }

    #[test]
    fn rows_the_sheet_leaves_out_are_rows_of_nulls() {
        let sheet = number_rows([2, 3, 9, 10]);
        let padded = collect(&sheet, range(1, 1, u64::MAX, 2), false, false).unwrap();
        assert_eq!(
            padded,
            [
                vec!["1: NULL", "2: \"2\"", "3: \"3\"", "4: NULL"],
                vec!["5: NULL", "6: NULL", "7: NULL", "8: NULL"],
                vec!["9: \"9\"", "10: \"10\""]
            ]
        );
        let stopped = collect(&sheet, range(2, 1, u64::MAX, 2), true, false).unwrap();
        assert_eq!(stopped, [vec!["2: \"2\"", "3: \"3\""]]);
        assert!(
            collect(&sheet, range(1, 1, u64::MAX, 2), true, false)
                .unwrap()
                .is_empty()
        );
    }

    #[test]
    fn a_range_is_filled_up_with_rows_of_nulls() {
        let batches = collect(&number_rows(1..=2), range(1, 1, 8, 2), false, true).unwrap();
        assert_eq!(batches.concat().len(), 7);
        assert_eq!(batches[0], ["1: \"1\"", "2: \"2\"", "3: NULL", "4: NULL"]);
        let empty = collect("", range(3, 1, 6, 3), false, true).unwrap();
        assert_eq!(
            empty,
            [vec!["3: NULL NULL", "4: NULL NULL", "5: NULL NULL"]]
        );
        let exact = collect(&number_rows(1..=4), range(1, 1, 5, 2), false, true).unwrap();
        assert_eq!(exact.len(), 1);
    }

    #[test]
    fn a_cell_that_was_passed_over_stays_null() {
        let sheet = r#"<row r="1"><c r="C1"><v>c</v></c><c r="A1"><v>a</v></c><c r="B1"><v>b</v></c></row>
            <row r="2"><c r="A2"><v>first</v></c><c r="A2"><v>second</v></c><c r="C2"><v>c</v></c></row>
            <row r="3"><c r="B3"><v>b</v></c><c r="A3"><v>a</v></c></row>"#;
        let batches = collect(sheet, range(1, 1, u64::MAX, 4), true, false).unwrap();
        assert_eq!(
            batches,
            [vec![
                "1: NULL NULL NULL",
                "2: \"second\" NULL \"c\"",
                "3: NULL NULL NULL"
            ]]
        );
    }

    #[test]
    fn rows_out_of_order_are_taken_as_they_come() {
        let sheet = number_rows([1, 4, 2, 3]);
        let batches = collect(&sheet, range(1, 1, u64::MAX, 2), false, false).unwrap();
        assert_eq!(
            batches.concat(),
            [
                "1: \"1\"", "2: NULL", "3: NULL", "4: \"4\"", "2: \"2\"", "3: \"3\""
            ]
        );
    }

    #[test]
    fn shared_strings_must_exist() {
        for index in ["3", "-1", "99999999999999999999"] {
            let sheet = format!(r#"<row r="1"><c r="A1" t="s"><v>{index}</v></c></row>"#);
            let error = collect(&sheet, range(1, 1, u64::MAX, 2), true, false).unwrap_err();
            assert_eq!(error.code, DuckDBError::InputInvalid);
            assert_eq!(
                error.message,
                format!(
                    "Cell 'A1' refers to shared string '{index}', which the xlsx file does not have"
                )
            );
        }
        let outside = r#"<row r="1"><c r="B1" t="s"><v>7</v></c></row>"#;
        assert!(collect(outside, range(1, 1, u64::MAX, 2), false, false).is_ok());
    }
}
