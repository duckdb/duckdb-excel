//! Adds a sheet to an existing workbook, or replaces one of its sheets.
//!
//! The workbook is rebuilt in a new archive: every entry that stays as it is gets copied in its compressed
//! form, and the few parts that must name the new sheet or its cell styles are patched as text, so that
//! whatever else a producer put into them survives.

use std::{
    collections::{HashMap, HashSet},
    io::Write,
};

use duckdb_neo::Result;
use excel_zip::ZipWriter;

use crate::{
    error::{from_io, invalid_input, io},
    xlsx::{
        archive::Archive,
        cell::parse_unsigned,
        escape::escape_xml,
        parsers::{
            relationships::{Relation, RelationshipsParser},
            styles_append::{NumberFormat, StylesAppendParser},
            workbook_append::{AppendSheetEntry, WorkbookAppendParser},
        },
        write::{
            package::{
                CONTENT_TYPES_PATH, STYLES_PATH, WORKBOOK_PATH, WORKBOOK_RELATIONS_PATH,
                WORKSHEETS_DIRECTORY, sheet_element, worksheet_content_type, worksheet_relation,
            },
            sheet::CellStyles,
        },
        xml::{XmlHandler, XmlReader},
    },
};

/// The number formats below this id are built into spreadsheets; a workbook defines its own from here on.
/// The first id is left to the format a new workbook defines for general numbers.
const FIRST_FREE_FORMAT: u64 = 165;

/// The number formats of the cell styles a sheet needs: date, timestamp, time, timestamp with milliseconds
/// and boolean.
const FORMAT_CODES: [&str; 5] = [
    "dd/mm/yy",
    "dd/mm/yyyy\\ hh:mm:ss",
    "hh:mm:ss",
    "dd/mm/yyyy\\ hh:mm:ss.000",
    "\"TRUE\";\"TRUE\";\"FALSE\"",
];

/// How a sheet gets into an existing workbook.
///
/// A new sheet is numbered after the largest sheet id and relationship id in use, where ids that are not of
/// the form `rId<number>` count as 0, and is stored as the first `sheet<number>.xml` no worksheet has. A sheet
/// that replaces another takes over its part, which must be in `xl/worksheets/`, and with it its ids.
#[derive(Debug)]
pub struct AppendPlan {
    /// The file name of the part of the sheet in `xl/worksheets/`.
    pub sheet_file: String,
    pub styles: CellStyles,
    /// The parts that replace those of the workbook, in the order they are written after the sheet.
    pub patched_parts: Vec<(&'static str, Vec<u8>)>,
    replaced_sheet: Option<String>,
}

impl AppendPlan {
    /// Plans to add the sheet `sheet_name` to `workbook`, or to put it in the place of the sheet of that name
    /// when `replace` is set. `workbook_path` is what errors call the workbook.
    pub fn new(
        workbook: &mut Archive,
        workbook_path: &str,
        sheet_name: &str,
        replace: bool,
    ) -> Result<AppendPlan> {
        let mut required = |name: &str| {
            workbook
                .read_entry(name)?
                .ok_or_else(|| io(format!("Required entry '{name}' not found in xlsx file")))
        };
        let content_types = required(CONTENT_TYPES_PATH)?;
        let workbook_xml = required(WORKBOOK_PATH)?;
        let relations_xml = required(WORKBOOK_RELATIONS_PATH)?;
        let styles_xml = required(STYLES_PATH)?;

        let sheets = parsed(&workbook_xml, WorkbookAppendParser::new())?.sheets;
        let relations = parsed(&relations_xml, RelationshipsParser::new())?.relations;
        let styles_parser = parsed(&styles_xml, StylesAppendParser::new())?;
        let needs = StyleNeeds::resolve(&styles_parser.number_formats, &styles_parser.cell_styles);

        let lower_name = sheet_name.to_ascii_lowercase();
        let existing = sheets
            .iter()
            .find(|sheet| sheet.name.to_ascii_lowercase() == lower_name);

        let mut patched_parts = Vec::new();
        let (sheet_file, replaced_sheet) = if replace {
            let existing = existing.ok_or_else(|| {
                invalid_input(format!(
                    "Sheet '{sheet_name}' not found in '{workbook_path}' (REPLACE)"
                ))
            })?;
            let (file, part) = replaced_part(existing, &relations, sheet_name)?;
            (file, Some(part))
        } else {
            if existing.is_some() {
                return Err(invalid_input(format!(
                    "Sheet '{sheet_name}' already exists in '{workbook_path}' (use REPLACE to overwrite)"
                )));
            }
            let sheet_id = next_sheet_id(&sheets);
            let relation = next_relation(&relations);
            let file = free_sheet_file(&relations);
            patched_parts.push((
                CONTENT_TYPES_PATH,
                insert_before_last(
                    &content_types,
                    "</Types>",
                    &worksheet_content_type(&file),
                    CONTENT_TYPES_PATH,
                )?,
            ));
            patched_parts.push((
                WORKBOOK_PATH,
                insert_before_last(
                    &workbook_xml,
                    "</sheets>",
                    &sheet_element(sheet_name, sheet_id, relation),
                    WORKBOOK_PATH,
                )?,
            ));
            patched_parts.push((
                WORKBOOK_RELATIONS_PATH,
                insert_before_last(
                    &relations_xml,
                    "</Relationships>",
                    &worksheet_relation(relation, &file),
                    WORKBOOK_RELATIONS_PATH,
                )?,
            ));
            (file, None)
        };
        if needs.changes_styles() {
            patched_parts.push((STYLES_PATH, needs.patch(&styles_xml)?));
        }
        Ok(AppendPlan {
            sheet_file,
            styles: needs.styles,
            patched_parts,
            replaced_sheet,
        })
    }

    /// Copies the entries of `workbook` that stay as they are into `target`. Of several entries with one name
    /// the last is the one that counts, and directories are left out.
    pub fn copy_kept_entries<W: Write>(
        &self,
        workbook: &mut Archive,
        target: &mut ZipWriter<W>,
    ) -> Result<()> {
        let mut dropped: HashSet<&str> = self.patched_parts.iter().map(|(name, _)| *name).collect();
        dropped.extend(self.replaced_sheet.as_deref());
        let entries = workbook.entries();
        let last_with_name: HashMap<&str, usize> = entries
            .iter()
            .enumerate()
            .map(|(index, entry)| (entry.name.as_str(), index))
            .collect();
        let kept: Vec<usize> = entries
            .iter()
            .enumerate()
            .filter(|(index, entry)| {
                last_with_name.get(entry.name.as_str()) == Some(index)
                    && !dropped.contains(entry.name.as_str())
                    && !entry.is_directory()
            })
            .map(|(index, _)| index)
            .collect();
        for index in kept {
            workbook.copy_entry(index, target)?;
        }
        Ok(())
    }

    /// Writes the patched parts into `target`.
    pub fn write_patched_parts<W: Write>(&self, target: &mut ZipWriter<W>) -> Result<()> {
        let written = self.patched_parts.iter().try_for_each(|(name, content)| {
            target.begin_file(name)?;
            target.write_all(content)?;
            target.end_file()
        });
        written.map_err(|error| from_io(&error))
    }
}

fn parsed<H: XmlHandler>(xml: &[u8], mut parser: H) -> Result<H> {
    XmlReader::new(xml, None).parse_all(&mut parser)?;
    Ok(parser)
}

fn replaced_part(
    existing: &AppendSheetEntry,
    relations: &[Relation],
    sheet_name: &str,
) -> Result<(String, String)> {
    let target = relations
        .iter()
        .find(|relation| relation.id == existing.relation_id)
        .map(|relation| relation.target.as_str())
        .filter(|target| !target.is_empty())
        .ok_or_else(|| {
            io(format!(
                "Could not resolve relationship '{}' for sheet '{sheet_name}'",
                existing.relation_id
            ))
        })?;
    let part = match target.strip_prefix('/') {
        Some(absolute) => absolute.to_owned(),
        None => format!("xl/{target}"),
    };
    let file = part.rsplit('/').next().unwrap_or_default();
    if part != format!("{WORKSHEETS_DIRECTORY}{file}") {
        return Err(io(format!(
            "REPLACE not supported: sheet '{sheet_name}' is stored at '{part}', expected xl/worksheets/<name>"
        )));
    }
    Ok((file.to_owned(), part))
}

fn next_sheet_id(sheets: &[AppendSheetEntry]) -> u64 {
    sheets
        .iter()
        .filter(|sheet| !sheet.sheet_id.is_empty())
        .map(|sheet| parse_unsigned(&sheet.sheet_id))
        .max()
        .unwrap_or(0)
        .wrapping_add(1)
}

fn relation_number(id: &str) -> u64 {
    match id.strip_prefix("rId") {
        Some(digits) if !digits.is_empty() && digits.bytes().all(|byte| byte.is_ascii_digit()) => {
            parse_unsigned(digits)
        }
        _ => 0,
    }
}

fn next_relation(relations: &[Relation]) -> u64 {
    relations
        .iter()
        .map(|relation| relation_number(&relation.id))
        .max()
        .unwrap_or(0)
        .wrapping_add(1)
}

fn free_sheet_file(relations: &[Relation]) -> String {
    let used: HashSet<u64> = relations
        .iter()
        .filter(|relation| relation.kind.ends_with("/worksheet"))
        .filter_map(|relation| {
            let file = relation.target.rsplit('/').next().unwrap_or_default();
            let digits = file.strip_prefix("sheet")?.strip_suffix(".xml")?;
            let is_number = !digits.is_empty() && digits.bytes().all(|byte| byte.is_ascii_digit());
            is_number.then(|| parse_unsigned(digits))
        })
        .collect();
    let number = (1u64..).find(|number| !used.contains(number)).unwrap_or(0);
    format!("sheet{number}.xml")
}

fn find(haystack: &[u8], needle: &[u8], from: usize) -> Option<usize> {
    haystack
        .get(from..)?
        .windows(needle.len())
        .position(|window| window == needle)
        .map(|position| position + from)
}

fn find_last(haystack: &[u8], needle: &[u8]) -> Option<usize> {
    haystack
        .windows(needle.len())
        .rposition(|window| window == needle)
}

fn insert_at(xml: &mut Vec<u8>, position: usize, insertion: &[u8]) {
    let position = position.min(xml.len());
    xml.splice(position..position, insertion.iter().copied());
}

fn insert_before_last(
    xml: &[u8],
    closing_tag: &str,
    insertion: &str,
    part: &str,
) -> Result<Vec<u8>> {
    let position = find_last(xml, closing_tag.as_bytes())
        .ok_or_else(|| io(format!("Cannot patch {part}: missing {closing_tag}")))?;
    let mut patched = xml.to_vec();
    insert_at(&mut patched, position, insertion.as_bytes());
    Ok(patched)
}

fn set_count(xml: &mut Vec<u8>, tag_start: &str, count: u64) {
    const ATTRIBUTE: &[u8] = b"count=\"";
    let value = find(xml, tag_start.as_bytes(), 0).and_then(|tag| {
        let tag_end = find(xml, b">", tag)?;
        let attribute = find(xml, ATTRIBUTE, tag).filter(|attribute| *attribute < tag_end)?;
        let value_start = attribute + ATTRIBUTE.len();
        let value_end = find(xml, b"\"", value_start)?;
        Some(value_start..value_end)
    });
    if let Some(value) = value {
        xml.splice(value, count.to_string().into_bytes());
    }
}

/// The cell styles a sheet uses in a workbook that has its own style sheet, and what the style sheet lacks
/// for them.
///
/// For each of [`FORMAT_CODES`] the first cell style with that number format is the one to use. A format the
/// workbook lacks gets the next free id, and a style it lacks is added after those it has; both are put in
/// front of the last closing tag of their lists, whose `count` is brought up to date if they have one.
#[derive(Debug)]
struct StyleNeeds {
    styles: CellStyles,
    new_formats: Vec<String>,
    new_cell_styles: Vec<String>,
    format_count: u64,
    cell_style_count: u64,
}

impl StyleNeeds {
    fn resolve(formats: &[NumberFormat], cell_styles: &[u64]) -> StyleNeeds {
        let mut formats: Vec<(u64, &str)> = formats
            .iter()
            .map(|format| (format.id, format.format_code.as_str()))
            .collect();
        let mut cell_styles = cell_styles.to_vec();
        let mut next_format = formats.iter().fold(FIRST_FREE_FORMAT, |next, (id, _)| {
            if *id >= next {
                id.wrapping_add(1)
            } else {
                next
            }
        });
        let mut new_formats = Vec::new();
        let mut new_cell_styles = Vec::new();

        let indexes = FORMAT_CODES.map(|code| {
            let known = formats.iter().find(|(_, known)| *known == code);
            let format = match known {
                Some((id, _)) => *id,
                None => {
                    let id = next_format;
                    next_format = next_format.wrapping_add(1);
                    formats.push((id, code));
                    new_formats.push(format!(
                        "<numFmt formatCode=\"{}\" numFmtId=\"{id}\"/>",
                        escape_xml(code)
                    ));
                    id
                }
            };
            match cell_styles.iter().position(|known| *known == format) {
                Some(index) => index as u64,
                None => {
                    cell_styles.push(format);
                    new_cell_styles.push(format!("<xf numFmtId=\"{format}\" xfId=\"0\"/>"));
                    cell_styles.len() as u64 - 1
                }
            }
        });
        let [date, timestamp_seconds, time, timestamp, boolean] = indexes;
        StyleNeeds {
            styles: CellStyles {
                date,
                timestamp_seconds,
                time,
                timestamp,
                boolean,
            },
            new_formats,
            new_cell_styles,
            format_count: formats.len() as u64,
            cell_style_count: cell_styles.len() as u64,
        }
    }

    fn changes_styles(&self) -> bool {
        !self.new_formats.is_empty() || !self.new_cell_styles.is_empty()
    }

    fn patch(&self, xml: &[u8]) -> Result<Vec<u8>> {
        let cannot_patch = |problem: &str| io(format!("Cannot patch styles.xml: {problem}"));
        let mut patched = xml.to_vec();
        if !self.new_cell_styles.is_empty() {
            let position = find_last(&patched, b"</cellXfs>")
                .ok_or_else(|| cannot_patch("missing </cellXfs>"))?;
            insert_at(
                &mut patched,
                position,
                self.new_cell_styles.concat().as_bytes(),
            );
            set_count(&mut patched, "<cellXfs", self.cell_style_count);
        }
        if !self.new_formats.is_empty() {
            let formats = self.new_formats.concat();
            if let Some(position) = find_last(&patched, b"</numFmts>") {
                insert_at(&mut patched, position, formats.as_bytes());
                set_count(&mut patched, "<numFmts", self.format_count);
            } else {
                let style_sheet = find(&patched, b"<styleSheet", 0)
                    .ok_or_else(|| cannot_patch("missing <styleSheet>"))?;
                let tag_end = find(&patched, b">", style_sheet)
                    .ok_or_else(|| cannot_patch("malformed <styleSheet>"))?;
                let block = format!(
                    "<numFmts count=\"{}\">{formats}</numFmts>",
                    self.new_formats.len()
                );
                insert_at(&mut patched, tag_end + 1, block.as_bytes());
            }
        }
        Ok(patched)
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use duckdb_neo::error::DuckDBError;

    fn format(id: u64, code: &str) -> NumberFormat {
        NumberFormat {
            id,
            format_code: code.to_owned(),
        }
    }

    fn relation(id: &str, kind: &str, target: &str) -> Relation {
        Relation {
            id: id.to_owned(),
            kind: kind.to_owned(),
            target: target.to_owned(),
        }
    }

    fn sheet(name: &str, relation_id: &str, sheet_id: &str) -> AppendSheetEntry {
        AppendSheetEntry {
            name: name.to_owned(),
            relation_id: relation_id.to_owned(),
            sheet_id: sheet_id.to_owned(),
        }
    }

    fn indexes(needs: &StyleNeeds) -> [u64; 5] {
        let styles = needs.styles;
        [
            styles.date,
            styles.timestamp_seconds,
            styles.time,
            styles.timestamp,
            styles.boolean,
        ]
    }

    #[test]
    fn the_styles_of_a_workbook_of_this_writer_are_reused() {
        let formats = [
            format(164, "General"),
            format(165, FORMAT_CODES[0]),
            format(166, FORMAT_CODES[1]),
            format(167, FORMAT_CODES[2]),
            format(168, FORMAT_CODES[3]),
            format(169, FORMAT_CODES[4]),
        ];
        let needs = StyleNeeds::resolve(&formats, &[164, 165, 166, 167, 168, 169]);
        assert_eq!(indexes(&needs), [1, 2, 3, 4, 5]);
        assert_eq!(needs.styles, CellStyles::NEW_WORKBOOK);
        assert!(!needs.changes_styles());
    }

    #[test]
    fn a_workbook_without_the_formats_gets_them_all() {
        let needs = StyleNeeds::resolve(&[], &[0, 14]);
        assert_eq!(indexes(&needs), [2, 3, 4, 5, 6]);
        assert_eq!(
            needs.new_formats,
            [
                r#"<numFmt formatCode="dd/mm/yy" numFmtId="165"/>"#,
                r#"<numFmt formatCode="dd/mm/yyyy\ hh:mm:ss" numFmtId="166"/>"#,
                r#"<numFmt formatCode="hh:mm:ss" numFmtId="167"/>"#,
                r#"<numFmt formatCode="dd/mm/yyyy\ hh:mm:ss.000" numFmtId="168"/>"#,
                r#"<numFmt formatCode="&quot;TRUE&quot;;&quot;TRUE&quot;;&quot;FALSE&quot;" numFmtId="169"/>"#,
            ]
        );
        assert_eq!(
            needs.new_cell_styles,
            [165u64, 166, 167, 168, 169].map(|id| format!(r#"<xf numFmtId="{id}" xfId="0"/>"#))
        );
        assert_eq!((needs.format_count, needs.cell_style_count), (5, 7));
    }

    #[test]
    fn known_formats_and_styles_are_found_and_new_ids_follow_the_largest() {
        let formats = [
            format(200, "hh:mm:ss"),
            format(170, "0.00"),
            format(171, "dd/mm/yy"),
        ];
        let needs = StyleNeeds::resolve(&formats, &[0, 200, 171, 200]);
        assert_eq!(indexes(&needs), [2, 4, 1, 5, 6]);
        assert_eq!(
            needs.new_formats,
            [
                r#"<numFmt formatCode="dd/mm/yyyy\ hh:mm:ss" numFmtId="201"/>"#,
                r#"<numFmt formatCode="dd/mm/yyyy\ hh:mm:ss.000" numFmtId="202"/>"#,
                r#"<numFmt formatCode="&quot;TRUE&quot;;&quot;TRUE&quot;;&quot;FALSE&quot;" numFmtId="203"/>"#,
            ]
        );
        assert_eq!(needs.new_cell_styles.len(), 3);
        assert_eq!((needs.format_count, needs.cell_style_count), (6, 7));

        let needs = StyleNeeds::resolve(&[format(165, "dd/mm/yy")], &[]);
        assert_eq!(indexes(&needs), [0, 1, 2, 3, 4]);
        assert_eq!(needs.new_formats.len(), 4);
        assert_eq!(needs.new_cell_styles.len(), 5);
    }

    #[test]
    fn format_ids_wrap_around() {
        let needs = StyleNeeds::resolve(&[format(u64::MAX, "0")], &[]);
        assert!(needs.new_formats[0].ends_with(r#"numFmtId="0"/>"#));
        assert!(needs.new_formats[4].ends_with(r#"numFmtId="4"/>"#));
    }

    fn patched(needs: &StyleNeeds, xml: &str) -> String {
        String::from_utf8(needs.patch(xml.as_bytes()).unwrap()).unwrap()
    }

    #[test]
    fn style_sheets_are_patched_in_place() {
        let needs = StyleNeeds::resolve(
            &[format(165, "dd/mm/yy"), format(166, "hh:mm:ss")],
            &[0, 165],
        );
        let xml = concat!(
            r#"<styleSheet xmlns="m"><numFmts count="2"><numFmt numFmtId="165" formatCode="dd/mm/yy"/>"#,
            r#"<numFmt numFmtId="166" formatCode="hh:mm:ss"/></numFmts><fonts count="9"/>"#,
            r#"<cellXfs count="2"><xf numFmtId="0"/><xf numFmtId="165"/></cellXfs><dxfs count="0"/></styleSheet>"#
        );
        assert_eq!(
            patched(&needs, xml),
            concat!(
                r#"<styleSheet xmlns="m"><numFmts count="5"><numFmt numFmtId="165" formatCode="dd/mm/yy"/>"#,
                r#"<numFmt numFmtId="166" formatCode="hh:mm:ss"/>"#,
                r#"<numFmt formatCode="dd/mm/yyyy\ hh:mm:ss" numFmtId="167"/>"#,
                r#"<numFmt formatCode="dd/mm/yyyy\ hh:mm:ss.000" numFmtId="168"/>"#,
                r#"<numFmt formatCode="&quot;TRUE&quot;;&quot;TRUE&quot;;&quot;FALSE&quot;" numFmtId="169"/>"#,
                r#"</numFmts><fonts count="9"/>"#,
                r#"<cellXfs count="6"><xf numFmtId="0"/><xf numFmtId="165"/>"#,
                r#"<xf numFmtId="167" xfId="0"/><xf numFmtId="166" xfId="0"/>"#,
                r#"<xf numFmtId="168" xfId="0"/><xf numFmtId="169" xfId="0"/>"#,
                r#"</cellXfs><dxfs count="0"/></styleSheet>"#
            )
        );
    }

    #[test]
    fn a_style_sheet_without_formats_gets_a_block_of_them() {
        let needs = StyleNeeds::resolve(&[], &[0]);
        let xml = "<?xml version=\"1.0\"?>\n<styleSheet a=\"b\">\n<fonts/><cellXfs count='1'><xf/></cellXfs></styleSheet>";
        let patched = patched(&needs, xml);
        assert!(patched.starts_with(concat!(
            "<?xml version=\"1.0\"?>\n<styleSheet a=\"b\">",
            r#"<numFmts count="5"><numFmt formatCode="dd/mm/yy" numFmtId="165"/>"#
        )));
        assert!(patched.contains(concat!(
            r#"numFmtId="169"/></numFmts>"#,
            "\n<fonts/><cellXfs count='1'><xf/>",
            r#"<xf numFmtId="165" xfId="0"/>"#
        )));
        assert!(patched.ends_with(r#"<xf numFmtId="169" xfId="0"/></cellXfs></styleSheet>"#));
    }

    #[test]
    fn only_a_count_inside_the_opening_tag_is_set() {
        let mut xml = b"<cellXfs><xf count=\"3\"/></cellXfs>".to_vec();
        set_count(&mut xml, "<cellXfs", 9);
        assert_eq!(xml, b"<cellXfs><xf count=\"3\"/></cellXfs>");
        let mut xml = b"<a/><cellXfs x=\"1\" count=\"12345\" y=\"2\">".to_vec();
        set_count(&mut xml, "<cellXfs", 9);
        assert_eq!(xml, b"<a/><cellXfs x=\"1\" count=\"9\" y=\"2\">");
        let mut xml = b"<cellXfs count=\"1".to_vec();
        set_count(&mut xml, "<cellXfs", 9);
        assert_eq!(xml, b"<cellXfs count=\"1");
    }

    #[test]
    fn style_sheets_that_cannot_be_patched() {
        let needs = StyleNeeds::resolve(&[], &[]);
        for (xml, problem) in [
            ("<styleSheet><cellXfs/></styleSheet>", "missing </cellXfs>"),
            ("<other><cellXfs></cellXfs></other>", "missing <styleSheet>"),
            (
                "<x:styleSheet><cellXfs></cellXfs></x:styleSheet>",
                "missing <styleSheet>",
            ),
            ("<cellXfs></cellXfs><styleSheet", "malformed <styleSheet>"),
        ] {
            let error = needs.patch(xml.as_bytes()).unwrap_err();
            assert_eq!(error.code, DuckDBError::IoGeneral);
            assert_eq!(error.message, format!("Cannot patch styles.xml: {problem}"));
        }
    }

    #[test]
    fn insertions_go_before_the_last_closing_tag() {
        let patched = insert_before_last(b"<a></a><a>x</a>\n", "</a>", "<b/>", "part.xml").unwrap();
        assert_eq!(patched, b"<a></a><a>x<b/></a>\n");
        let error = insert_before_last(b"<a/>", "</a>", "<b/>", "part.xml").unwrap_err();
        assert_eq!(error.code, DuckDBError::IoGeneral);
        assert_eq!(error.message, "Cannot patch part.xml: missing </a>");
    }

    #[test]
    fn ids_follow_the_largest_in_use() {
        let sheets = [
            sheet("a", "rId1", "3"),
            sheet("b", "rId2", ""),
            sheet("c", "rId9", " 12x"),
        ];
        assert_eq!(next_sheet_id(&sheets), 13);
        assert_eq!(next_sheet_id(&[]), 1);
        assert_eq!(next_sheet_id(&[sheet("a", "rId1", "-1")]), 0);

        let relations = [
            relation("rId3", "t/styles", "styles.xml"),
            relation("rId12", "t/worksheet", "worksheets/sheet1.xml"),
            relation("rIdx", "t/worksheet", "worksheets/sheet2.xml"),
            relation("R99", "t/theme", "theme/theme1.xml"),
            relation("rId", "t/other", "other.xml"),
        ];
        assert_eq!(next_relation(&relations), 13);
        assert_eq!(next_relation(&[]), 1);
        assert_eq!(
            next_relation(&[relation("rId99999999999999999999", "t", "x")]),
            0
        );
    }

    #[test]
    fn the_sheet_file_is_the_first_free_one() {
        let worksheet = |target: &str| relation("rId1", "http://x/relationships/worksheet", target);
        assert_eq!(free_sheet_file(&[]), "sheet1.xml");
        let relations = [
            worksheet("worksheets/sheet1.xml"),
            worksheet("/xl/worksheets/sheet2.xml"),
            worksheet("sheet04.xml"),
            worksheet("worksheets/sheet3x.xml"),
            worksheet("worksheets/Sheet3.xml"),
            worksheet("worksheets/sheet.xml"),
            relation(
                "rId7",
                "http://x/relationships/chartsheet",
                "chartsheets/sheet3.xml",
            ),
        ];
        assert_eq!(free_sheet_file(&relations), "sheet3.xml");
        let relations = [worksheet("worksheets/sheet2.xml")];
        assert_eq!(free_sheet_file(&relations), "sheet1.xml");
    }

    #[test]
    fn a_replaced_sheet_keeps_its_part() {
        let relations = [
            relation("rId1", "t/worksheet", "worksheets/sheet7.xml"),
            relation("rId2", "t/worksheet", "/xl/worksheets/data.xml"),
            relation("rId3", "t/worksheet", "sheets/sheet1.xml"),
            relation("rId4", "t/worksheet", ""),
        ];
        let part = |id: &str| replaced_part(&sheet("S", id, "1"), &relations, "S");
        assert_eq!(
            part("rId1").unwrap(),
            (
                "sheet7.xml".to_owned(),
                "xl/worksheets/sheet7.xml".to_owned()
            )
        );
        assert_eq!(
            part("rId2").unwrap(),
            ("data.xml".to_owned(), "xl/worksheets/data.xml".to_owned())
        );
        let error = part("rId3").unwrap_err();
        assert_eq!(error.code, DuckDBError::IoGeneral);
        assert_eq!(
            error.message,
            "REPLACE not supported: sheet 'S' is stored at 'xl/sheets/sheet1.xml', expected xl/worksheets/<name>"
        );
        for missing in ["rId4", "rId5"] {
            assert_eq!(
                part(missing).unwrap_err().message,
                format!("Could not resolve relationship '{missing}' for sheet 'S'")
            );
        }
    }
}
