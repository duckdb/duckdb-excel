//! Finds the sheet to read in a workbook, its range, its column names and its column types.

use std::collections::HashMap;

use duckdb_neo::Result;

use crate::{
    duckdb_compat::candidates_error_message,
    error::{binder, invalid_input},
    xlsx::{
        archive::Archive,
        cell::{CellPos, CellRange, CellType, MAX_CELL_COLS, parse_integer},
        escape::unescape_xml_entities,
        options::ReadOptions,
        parsers::{
            content_types::ContentTypesParser,
            header_sniffer::{HeaderSniffer, SniffedCell},
            range_sniffer::RangeSniffer,
            relationships::RelationshipsParser,
            shared_strings::{SharedStringsParser, StringSearch},
            styles::{NumberKind, StylesParser},
            workbook::WorkbookParser,
            worksheet::{MAX_EVENT_LEN, WorksheetParser},
        },
    },
};

pub const SHARED_STRINGS_PATH: &str = "xl/sharedStrings.xml";

/// The type `read_xlsx` gives a column.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ColumnType {
    Varchar,
    Boolean,
    Double,
    Date,
    Time,
    Timestamp,
}

impl ColumnType {
    /// The name DuckDB shows for the type.
    pub fn name(self) -> &'static str {
        match self {
            ColumnType::Varchar => "VARCHAR",
            ColumnType::Boolean => "BOOLEAN",
            ColumnType::Double => "DOUBLE",
            ColumnType::Date => "DATE",
            ColumnType::Time => "TIME",
            ColumnType::Timestamp => "TIMESTAMP",
        }
    }
}

/// A column of the range that is read.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SheetColumn {
    /// The text of the header cell, or the generated name.
    pub name: String,
    pub column_type: ColumnType,
    /// The type of the cell the column type was taken from.
    pub source_type: CellType,
}

/// Where the rows of a sheet are and what their columns look like.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct SheetLayout {
    /// The part of the workbook that holds the sheet.
    pub sheet_path: String,
    /// The cells of the rows after the header.
    pub range: CellRange,
    pub columns: Vec<SheetColumn>,
}

/// Finds the sheet `options` ask for in the workbook at `file_path` and looks at its first rows.
pub fn resolve_sheet(
    archive: &mut Archive,
    file_path: &str,
    options: &ReadOptions,
) -> Result<SheetLayout> {
    let sheet_path = find_sheet(archive, file_path, &options.sheet)?;
    let cell_styles = parse_cell_styles(archive)?;
    let range = if options.has_explicit_range {
        options.range
    } else {
        sniff_range(archive, &sheet_path)?
    };
    sniff_header(archive, sheet_path, range, options, &cell_styles)
}

fn sheet_not_found(sheet_path: &str) -> duckdb_neo::Error {
    binder(format!("Sheet '{sheet_path}' not found in xlsx file"))
}

fn find_sheet(archive: &mut Archive, file_path: &str, sheet: &str) -> Result<String> {
    if !archive.parse_entry("[Content_Types].xml", &mut ContentTypesParser::new())? {
        return Err(binder("No [Content_Types].xml found in xlsx file"));
    }
    let mut workbook = WorkbookParser::new();
    if !archive.parse_entry("xl/workbook.xml", &mut workbook)? {
        return Err(binder("No xl/workbook.xml found in xlsx file"));
    }
    let mut relationships = RelationshipsParser::new();
    if !archive.parse_entry("xl/_rels/workbook.xml.rels", &mut relationships)? {
        return Err(binder("No xl/_rels/workbook.xml.rels found in xlsx file"));
    }

    // Of several relationships with one id the last counts, and so does the last of several sheets with one
    // name. Sheets are looked up by their name with entities resolved once more than XML asks for.
    let sheet_targets: HashMap<&str, &str> = relationships
        .relations
        .iter()
        .filter(|relation| relation.kind.ends_with("/worksheet"))
        .map(|relation| (relation.id.as_str(), relation.target.as_str()))
        .collect();
    let mut names: Vec<String> = Vec::new();
    let mut paths: HashMap<String, String> = HashMap::new();
    let mut primary_sheet = "";
    for entry in &workbook.sheets {
        let Some(target) = sheet_targets.get(entry.relation_id.as_str()) else {
            continue;
        };
        let path = match target.strip_prefix('/') {
            Some(absolute) if absolute.starts_with("xl/") => absolute.to_owned(),
            _ => format!("xl/{target}"),
        };
        let name = unescape_xml_entities(&entry.name);
        if paths.insert(name.clone(), path).is_none() {
            names.push(name);
        }
        if primary_sheet.is_empty() {
            primary_sheet = &entry.name;
        }
    }
    if names.is_empty() {
        return Err(binder(
            "No sheets found in xlsx file (is the file corrupt?)",
        ));
    }

    let sheet = if sheet.is_empty() {
        primary_sheet
    } else {
        sheet
    };
    match paths.remove(&unescape_xml_entities(sheet)) {
        Some(path) => Ok(path),
        None => {
            let suggestions = candidates_error_message(&names, sheet, "Did you mean");
            Err(binder(format!(
                "Sheet \"{sheet}\" not found in xlsx file \"{file_path}\"{suggestions}"
            )))
        }
    }
}

fn parse_cell_styles(archive: &mut Archive) -> Result<Vec<NumberKind>> {
    let mut styles = StylesParser::new();
    archive.parse_entry("xl/styles.xml", &mut styles)?;
    Ok(styles.cell_styles)
}

fn sniff_range(archive: &mut Archive, sheet_path: &str) -> Result<CellRange> {
    let mut parser = WorksheetParser::new(RangeSniffer::new());
    if !archive.parse_entry_with_limit(sheet_path, Some(MAX_EVENT_LEN), &mut parser)? {
        return Err(sheet_not_found(sheet_path));
    }
    let range = parser.visitor.range();
    if range.width() > MAX_CELL_COLS {
        return Err(invalid_input(format!(
            "The first row of the sheet spans {} columns, more than the {MAX_CELL_COLS} a sheet can have \
             (is the file corrupted?)",
            range.width()
        )));
    }
    Ok(range)
}

fn sniff_header(
    archive: &mut Archive,
    sheet_path: String,
    range: CellRange,
    options: &ReadOptions,
    cell_styles: &[NumberKind],
) -> Result<SheetLayout> {
    let default_cell_type = options.default_cell_type;
    let mut parser = WorksheetParser::new(HeaderSniffer::new(
        range,
        options.header_mode,
        options.has_explicit_range,
        default_cell_type,
    ));
    if !archive.parse_entry_with_limit(&sheet_path, Some(MAX_EVENT_LEN), &mut parser)? {
        return Err(sheet_not_found(&sheet_path));
    }
    let HeaderSniffer {
        mut header_cells,
        mut column_cells,
        range,
        ..
    } = parser.visitor;

    if column_cells.is_empty() {
        if header_cells.is_empty() {
            if !options.has_explicit_range {
                return Err(binder("No rows found in xlsx file"));
            }
            // A range without any row still has columns, named after the cells of its first row.
            header_cells = (range.beg.col..range.end.col)
                .map(|col| {
                    let pos = CellPos::new(range.beg.row, col);
                    SniffedCell {
                        cell_type: default_cell_type,
                        pos,
                        data: pos.to_string(),
                        style: 0,
                    }
                })
                .collect();
        }
        // A header without rows below it still makes a table; its columns get the type of an empty cell.
        column_cells = header_cells
            .iter()
            .map(|cell| SniffedCell {
                cell_type: default_cell_type,
                pos: cell.pos,
                data: String::new(),
                style: 0,
            })
            .collect();
    }
    if header_cells.len() != column_cells.len() {
        return Err(invalid_input(
            "The first rows of the sheet do not have the same number of cells (is the file corrupted?)",
        ));
    }

    resolve_shared_strings(&mut header_cells, archive)?;
    let columns = header_cells
        .into_iter()
        .zip(&column_cells)
        .map(|(header, cell)| {
            Ok(SheetColumn {
                name: header.data,
                column_type: column_type(cell, options.all_varchar, cell_styles)?,
                source_type: cell.cell_type,
            })
        })
        .collect::<Result<Vec<_>>>()?;
    Ok(SheetLayout {
        sheet_path,
        range,
        columns,
    })
}

fn resolve_shared_strings(header_cells: &mut [SniffedCell], archive: &mut Archive) -> Result<()> {
    let shared: Vec<(usize, i64)> = header_cells
        .iter()
        .enumerate()
        .filter(|(_, cell)| cell.cell_type == CellType::SharedString)
        .map(|(position, cell)| (position, parse_integer(&cell.data)))
        .collect();
    if shared.is_empty() {
        return Ok(());
    }
    let wanted = shared
        .iter()
        .filter_map(|(_, index)| usize::try_from(*index).ok())
        .collect();
    let mut parser = SharedStringsParser::new(StringSearch::new(wanted));
    if !archive.parse_entry(SHARED_STRINGS_PATH, &mut parser)? {
        return Err(binder("No shared strings found in xlsx file"));
    }
    let found = parser.sink.found;
    for (position, index) in shared {
        let string = usize::try_from(index)
            .ok()
            .and_then(|index| found.get(&index));
        match (string, header_cells.get_mut(position)) {
            (Some(string), Some(cell)) => cell.data.clone_from(string),
            _ => {
                return Err(binder(format!(
                    "Shared string {index} of the header not found in xlsx file"
                )));
            }
        }
    }
    Ok(())
}

fn column_type(
    cell: &SniffedCell,
    all_varchar: bool,
    cell_styles: &[NumberKind],
) -> Result<ColumnType> {
    if all_varchar {
        return Ok(ColumnType::Varchar);
    }
    if cell.data.eq_ignore_ascii_case("true") || cell.data.eq_ignore_ascii_case("false") {
        return Ok(ColumnType::Boolean);
    }
    match cell.cell_type {
        CellType::Number => {
            let style = usize::try_from(cell.style)
                .ok()
                .and_then(|style| cell_styles.get(style));
            Ok(match style {
                Some(NumberKind::Date) => ColumnType::Date,
                Some(NumberKind::Time) => ColumnType::Time,
                Some(NumberKind::Timestamp) => ColumnType::Timestamp,
                Some(NumberKind::Double) | None => ColumnType::Double,
            })
        }
        CellType::Boolean => Ok(ColumnType::Boolean),
        CellType::SharedString
        | CellType::InlineString
        | CellType::FormulaString
        | CellType::Error => Ok(ColumnType::Varchar),
        CellType::Date => Ok(ColumnType::Date),
        CellType::Unknown => Err(binder("Unknown cell type in xlsx file")),
    }
}
