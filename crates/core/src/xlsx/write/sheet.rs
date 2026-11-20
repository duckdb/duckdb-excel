//! Writes a worksheet part row by row.

use std::io::Write;

use duckdb_neo::Result;

use crate::{
    error::{from_io, internal, invalid_input},
    xlsx::{
        cell::{CellPos, MAX_CELL_ROWS},
        escape::escape_text_into,
        write::package::XML_DECLARATION,
    },
};

const WORKSHEET_START: &str = concat!(
    "\n\t<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"",
    "\n\t           xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"",
    "\n\t           xmlns:mx=\"http://schemas.microsoft.com/office/mac/excel/2008/main\"",
    "\n\t           xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"",
    "\n\t           xmlns:mv=\"urn:schemas-microsoft-com:mac:vml\"",
    "\n\t           xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"",
    "\n\t           xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\"",
    "\n\t           xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\"",
    "\n\t           xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">",
    "\n\t<sheetData>",
    "\n\t",
);
const WORKSHEET_END: &str = "</sheetData></worksheet>";

const TEXT_START: &[u8] = b"\" t=\"inlineStr\"><is><t>";
const TEXT_END: &[u8] = b"</t></is></c>";
const VALUE_END: &[u8] = b"</v></c>";

/// Rows gather here until there is this much to hand to the output at once.
const BUFFER_LEN: usize = 64 * 1024;

/// What the cells of a column hold.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum CellKind {
    Number,
    /// Text that is written into the cell itself, escaped.
    Text,
    /// `1` or `0`.
    Boolean,
    /// Serial numbers shown as a date, a time of day, or a date and time with or without milliseconds. A date
    /// or date and time whose serial number is `inf` or `-inf` is written as the text `infinity` or `-infinity`.
    Date,
    Time,
    Timestamp,
    TimestampSeconds,
}

/// The cell styles, as indexes into the `cellXfs` of the workbook, that show a number as what it stands for.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct CellStyles {
    pub date: u64,
    pub timestamp_seconds: u64,
    pub time: u64,
    pub timestamp: u64,
    pub boolean: u64,
}

impl CellStyles {
    /// The styles of the style sheet a new workbook gets.
    pub const NEW_WORKBOOK: CellStyles = CellStyles {
        date: 1,
        timestamp_seconds: 2,
        time: 3,
        timestamp: 4,
        boolean: 5,
    };
}

#[derive(Debug)]
struct ColumnCells {
    reference_start: Vec<u8>,
    value_start: Vec<u8>,
    is_text: bool,
    is_date: bool,
}

fn infinity_text(serial: &[u8]) -> Option<&'static [u8]> {
    match serial {
        b"inf" => Some(b"infinity"),
        b"-inf" => Some(b"-infinity"),
        _ => None,
    }
}

/// Writes the rows of a sheet into `output`, which is the worksheet part.
///
/// A row holds the cells that have a value; the cells of a column are of the kind the column was given, but
/// for the cells of a header row, which are text. The sheet can take `row_limit` rows and fails on the row
/// after that, once the row is written.
#[derive(Debug)]
pub struct SheetWriter<W: Write> {
    output: W,
    buffer: Vec<u8>,
    columns: Vec<ColumnCells>,
    row_number: Vec<u8>,
    rows: u64,
    row_limit: u64,
}

impl<W: Write> SheetWriter<W> {
    /// Starts the part with a column of cells for each of `kinds`.
    pub fn begin(
        output: W,
        kinds: &[CellKind],
        styles: CellStyles,
        row_limit: u64,
    ) -> SheetWriter<W> {
        let columns = (1u64..)
            .zip(kinds)
            .map(|(number, kind)| {
                let name = CellPos::new(1, number).column_name();
                let style = |style: u64| format!("\" s=\"{style}\"><v>");
                let value_start = match kind {
                    CellKind::Number => "\"><v>".to_owned(),
                    CellKind::Text => String::new(),
                    CellKind::Boolean => format!("\" t=\"b\" s=\"{}\"><v>", styles.boolean),
                    CellKind::Date => style(styles.date),
                    CellKind::Time => style(styles.time),
                    CellKind::Timestamp => style(styles.timestamp),
                    CellKind::TimestampSeconds => style(styles.timestamp_seconds),
                };
                ColumnCells {
                    reference_start: format!("<c r=\"{name}").into_bytes(),
                    value_start: value_start.into_bytes(),
                    is_text: *kind == CellKind::Text,
                    is_date: matches!(
                        kind,
                        CellKind::Date | CellKind::Timestamp | CellKind::TimestampSeconds
                    ),
                }
            })
            .collect();
        let mut buffer = Vec::with_capacity(BUFFER_LEN + BUFFER_LEN / 4);
        buffer.extend_from_slice(XML_DECLARATION.as_bytes());
        buffer.extend_from_slice(WORKSHEET_START.as_bytes());
        SheetWriter {
            output,
            buffer,
            columns,
            row_number: b"1".to_vec(),
            rows: 0,
            row_limit,
        }
    }

    pub fn begin_row(&mut self) {
        self.buffer.extend_from_slice(b"<row r=\"");
        self.buffer.extend_from_slice(&self.row_number);
        self.buffer.extend_from_slice(b"\">");
    }

    /// Writes the cell of `column`, counted from 0, in the open row; `value` is what the kind of the column
    /// says it is. An empty cell is one that is not written.
    pub fn cell(&mut self, column: usize, value: &[u8]) -> Result<()> {
        let cells = self.columns.get(column).ok_or_else(no_such_column)?;
        self.buffer.extend_from_slice(&cells.reference_start);
        self.buffer.extend_from_slice(&self.row_number);
        if cells.is_text {
            self.buffer.extend_from_slice(TEXT_START);
            escape_text_into(value, &mut self.buffer);
            self.buffer.extend_from_slice(TEXT_END);
        } else if cells.is_date
            && let Some(infinity) = infinity_text(value)
        {
            // No number stands for an infinite date, so it is written as the text DuckDB reads as one.
            self.buffer.extend_from_slice(TEXT_START);
            self.buffer.extend_from_slice(infinity);
            self.buffer.extend_from_slice(TEXT_END);
        } else {
            self.buffer.extend_from_slice(&cells.value_start);
            self.buffer.extend_from_slice(value);
            self.buffer.extend_from_slice(VALUE_END);
        }
        Ok(())
    }

    /// Writes `text` as the cell of `column` in the open row, whatever the kind of the column is.
    pub fn text_cell(&mut self, column: usize, text: &str) -> Result<()> {
        let cells = self.columns.get(column).ok_or_else(no_such_column)?;
        self.buffer.extend_from_slice(&cells.reference_start);
        self.buffer.extend_from_slice(&self.row_number);
        self.buffer.extend_from_slice(TEXT_START);
        escape_text_into(text.as_bytes(), &mut self.buffer);
        self.buffer.extend_from_slice(TEXT_END);
        Ok(())
    }

    pub fn end_row(&mut self) -> Result<()> {
        self.buffer.extend_from_slice(b"</row>");
        if self.buffer.len() >= BUFFER_LEN {
            self.write_buffer()?;
        }
        self.rows += 1;
        self.row_number.clear();
        write_decimal(self.rows.wrapping_add(1), &mut self.row_number);
        if self.rows > self.row_limit {
            return Err(self.row_limit_exceeded());
        }
        Ok(())
    }

    /// Ends the part and returns the output.
    pub fn finish(mut self) -> Result<W> {
        self.buffer.extend_from_slice(WORKSHEET_END.as_bytes());
        self.write_buffer()?;
        Ok(self.output)
    }

    fn write_buffer(&mut self) -> Result<()> {
        let written = self.output.write_all(&self.buffer);
        self.buffer.clear();
        written.map_err(|error| from_io(&error))
    }

    fn row_limit_exceeded(&self) -> duckdb_neo::Error {
        let limit = self.row_limit;
        let mut message = format!("XLSX: Sheet row limit of '{limit}' rows exceeded!");
        if limit >= MAX_CELL_ROWS {
            message.push_str(&format!(
                "\n * XLSX files and compatible applications generally have a limit of '{MAX_CELL_ROWS}' rows\
                 \n * You can export larger sheets at your own risk by setting the 'sheet_row_limit' parameter \
                 to a higher value"
            ));
        }
        invalid_input(message)
    }
}

fn no_such_column() -> duckdb_neo::Error {
    internal("COPY TO xlsx: a row has more cells than the sheet has columns")
}

fn write_decimal(mut number: u64, text: &mut Vec<u8>) {
    let start = text.len();
    loop {
        text.push(b'0' + (number % 10) as u8);
        number /= 10;
        if number == 0 {
            break;
        }
    }
    if let Some(digits) = text.get_mut(start..) {
        digits.reverse();
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use duckdb_neo::error::DuckDBError;

    const ALL_KINDS: [CellKind; 7] = [
        CellKind::Number,
        CellKind::Text,
        CellKind::Boolean,
        CellKind::Date,
        CellKind::Time,
        CellKind::Timestamp,
        CellKind::TimestampSeconds,
    ];

    fn sheet(kinds: &[CellKind], row_limit: u64) -> SheetWriter<Vec<u8>> {
        SheetWriter::begin(Vec::new(), kinds, CellStyles::NEW_WORKBOOK, row_limit)
    }

    fn rows_of(sheet: SheetWriter<Vec<u8>>) -> String {
        let part = String::from_utf8(sheet.finish().unwrap()).unwrap();
        let start = format!("{XML_DECLARATION}{WORKSHEET_START}");
        part.strip_prefix(&start)
            .and_then(|rest| rest.strip_suffix(WORKSHEET_END))
            .unwrap_or_else(|| panic!("unexpected frame: {part}"))
            .to_owned()
    }

    #[test]
    fn every_kind_of_cell() {
        let mut sheet = sheet(&ALL_KINDS, 10);
        sheet.begin_row();
        for (column, value) in ["1.5", "a<b", "1", "45366.0", "0.5", "45366.25", "45366.75"]
            .iter()
            .enumerate()
        {
            sheet.cell(column, value.as_bytes()).unwrap();
        }
        sheet.end_row().unwrap();
        assert_eq!(
            rows_of(sheet),
            concat!(
                r#"<row r="1"><c r="A1"><v>1.5</v></c>"#,
                r#"<c r="B1" t="inlineStr"><is><t>a&lt;b</t></is></c>"#,
                r#"<c r="C1" t="b" s="5"><v>1</v></c>"#,
                r#"<c r="D1" s="1"><v>45366.0</v></c>"#,
                r#"<c r="E1" s="3"><v>0.5</v></c>"#,
                r#"<c r="F1" s="4"><v>45366.25</v></c>"#,
                r#"<c r="G1" s="2"><v>45366.75</v></c></row>"#,
            )
        );
    }

    #[test]
    fn infinite_dates_are_text_and_other_infinite_numbers_are_not() {
        let mut sheet = sheet(&ALL_KINDS, 10);
        for value in ["inf", "-inf"] {
            sheet.begin_row();
            for column in 0..ALL_KINDS.len() {
                sheet.cell(column, value.as_bytes()).unwrap();
            }
            sheet.end_row().unwrap();
        }
        sheet.begin_row();
        sheet.cell(3, b"infinity").unwrap();
        sheet.cell(5, b"nan").unwrap();
        sheet.end_row().unwrap();
        assert_eq!(
            rows_of(sheet),
            concat!(
                r#"<row r="1"><c r="A1"><v>inf</v></c>"#,
                r#"<c r="B1" t="inlineStr"><is><t>inf</t></is></c>"#,
                r#"<c r="C1" t="b" s="5"><v>inf</v></c>"#,
                r#"<c r="D1" t="inlineStr"><is><t>infinity</t></is></c>"#,
                r#"<c r="E1" s="3"><v>inf</v></c>"#,
                r#"<c r="F1" t="inlineStr"><is><t>infinity</t></is></c>"#,
                r#"<c r="G1" t="inlineStr"><is><t>infinity</t></is></c></row>"#,
                r#"<row r="2"><c r="A2"><v>-inf</v></c>"#,
                r#"<c r="B2" t="inlineStr"><is><t>-inf</t></is></c>"#,
                r#"<c r="C2" t="b" s="5"><v>-inf</v></c>"#,
                r#"<c r="D2" t="inlineStr"><is><t>-infinity</t></is></c>"#,
                r#"<c r="E2" s="3"><v>-inf</v></c>"#,
                r#"<c r="F2" t="inlineStr"><is><t>-infinity</t></is></c>"#,
                r#"<c r="G2" t="inlineStr"><is><t>-infinity</t></is></c></row>"#,
                r#"<row r="3"><c r="D3" s="1"><v>infinity</v></c><c r="F3" s="4"><v>nan</v></c></row>"#,
            )
        );
    }

    #[test]
    fn styles_are_those_the_workbook_has() {
        let styles = CellStyles {
            date: 7,
            timestamp_seconds: 12,
            time: 0,
            timestamp: 3,
            boolean: 18446744073709551615,
        };
        let mut sheet = SheetWriter::begin(Vec::new(), &ALL_KINDS, styles, 10);
        sheet.begin_row();
        for column in 0..ALL_KINDS.len() {
            sheet.cell(column, b"0").unwrap();
        }
        sheet.end_row().unwrap();
        assert_eq!(
            rows_of(sheet),
            concat!(
                r#"<row r="1"><c r="A1"><v>0</v></c><c r="B1" t="inlineStr"><is><t>0</t></is></c>"#,
                r#"<c r="C1" t="b" s="18446744073709551615"><v>0</v></c><c r="D1" s="7"><v>0</v></c>"#,
                r#"<c r="E1" s="0"><v>0</v></c><c r="F1" s="3"><v>0</v></c><c r="G1" s="12"><v>0</v></c></row>"#,
            )
        );
    }

    #[test]
    fn rows_are_numbered_and_empty_cells_left_out() {
        let mut sheet = sheet(&[CellKind::Number; 28], 100);
        for row in 0..12 {
            sheet.begin_row();
            if row == 0 {
                sheet.text_cell(0, "a\"b'c&\0d").unwrap();
                sheet.text_cell(27, "last").unwrap();
            } else if row == 10 {
                sheet.cell(1, b"7").unwrap();
                sheet.cell(26, b"8").unwrap();
            }
            sheet.end_row().unwrap();
        }
        let rows = rows_of(sheet);
        assert!(rows.starts_with(concat!(
            r#"<row r="1"><c r="A1" t="inlineStr"><is><t>a&quot;b&apos;c&amp;_x0000_d</t></is></c>"#,
            r#"<c r="AB1" t="inlineStr"><is><t>last</t></is></c></row><row r="2"></row>"#
        )));
        assert!(rows.ends_with(concat!(
            r#"<row r="10"></row><row r="11"><c r="B11"><v>7</v></c><c r="AA11"><v>8</v></c></row>"#,
            r#"<row r="12"></row>"#
        )));
    }

    #[test]
    fn a_sheet_without_columns_has_empty_rows() {
        let mut sheet = sheet(&[], 5);
        sheet.begin_row();
        sheet.end_row().unwrap();
        assert_eq!(
            sheet.cell(0, b"1").unwrap_err().code,
            DuckDBError::RuntimeInternal
        );
        assert_eq!(rows_of(sheet), r#"<row r="1"></row>"#);
    }

    #[test]
    fn the_row_after_the_limit_is_an_error() {
        let mut sheet = sheet(&[CellKind::Number], 2);
        for _ in 0..2 {
            sheet.begin_row();
            sheet.end_row().unwrap();
        }
        sheet.begin_row();
        let error = sheet.end_row().unwrap_err();
        assert_eq!(error.code, DuckDBError::InputInvalid);
        assert_eq!(error.message, "XLSX: Sheet row limit of '2' rows exceeded!");

        let mut sheet = self::sheet(&[CellKind::Number], 0);
        sheet.begin_row();
        assert_eq!(
            sheet.end_row().unwrap_err().message,
            "XLSX: Sheet row limit of '0' rows exceeded!"
        );
    }

    #[test]
    fn a_limit_no_smaller_than_that_of_a_sheet_explains_itself() {
        let mut sheet = sheet(&[], MAX_CELL_ROWS);
        sheet.rows = MAX_CELL_ROWS;
        sheet.begin_row();
        assert_eq!(
            sheet.end_row().unwrap_err().message,
            "XLSX: Sheet row limit of '1048576' rows exceeded!\n \
             * XLSX files and compatible applications generally have a limit of '1048576' rows\n \
             * You can export larger sheets at your own risk by setting the 'sheet_row_limit' parameter to a \
             higher value"
        );
    }

    #[test]
    fn large_sheets_reach_the_output_in_pieces() {
        struct Pieces(Vec<usize>);
        impl Write for Pieces {
            fn write(&mut self, buffer: &[u8]) -> std::io::Result<usize> {
                self.0.push(buffer.len());
                Ok(buffer.len())
            }
            fn flush(&mut self) -> std::io::Result<()> {
                Ok(())
            }
        }
        let mut sheet = SheetWriter::begin(
            Pieces(Vec::new()),
            &[CellKind::Text],
            CellStyles::NEW_WORKBOOK,
            1000,
        );
        let text = "x".repeat(1000);
        for _ in 0..200 {
            sheet.begin_row();
            sheet.cell(0, text.as_bytes()).unwrap();
            sheet.end_row().unwrap();
        }
        let pieces = sheet.finish().unwrap().0;
        assert!(pieces.len() >= 3, "{pieces:?}");
        assert!(pieces.iter().all(|len| *len < 2 * BUFFER_LEN), "{pieces:?}");
        assert!(pieces.iter().sum::<usize>() > 200 * 1000);
    }

    #[test]
    fn decimal_numbers() {
        for number in [0, 9, 10, 1_048_576, u64::MAX] {
            let mut text = b"r".to_vec();
            write_decimal(number, &mut text);
            assert_eq!(text, format!("r{number}").into_bytes());
        }
    }
}
