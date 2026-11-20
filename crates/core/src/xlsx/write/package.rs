//! The parts of a new workbook besides its sheet, and the elements that tie a sheet into a workbook.
//!
//! The parts are fixed texts, whitespace included, so that the same rows always give the same workbook.

use std::io::{self, Write};

use excel_zip::ZipWriter;

use crate::xlsx::escape::escape_xml;

pub const XML_DECLARATION: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";

pub const CONTENT_TYPES_PATH: &str = "[Content_Types].xml";
pub const WORKBOOK_PATH: &str = "xl/workbook.xml";
pub const WORKBOOK_RELATIONS_PATH: &str = "xl/_rels/workbook.xml.rels";
pub const STYLES_PATH: &str = "xl/styles.xml";
pub const WORKSHEETS_DIRECTORY: &str = "xl/worksheets/";

/// The file name of the sheet of a new workbook, and what its relationship and the sheet itself are numbered.
pub const FIRST_SHEET_FILE: &str = "sheet1.xml";
const FIRST_SHEET_RELATION: u64 = 3;
const FIRST_SHEET_ID: u64 = 1;

const CONTENT_TYPES_START: &str = concat!(
    "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">",
    "<Default Extension=\"xml\" ContentType=\"application/xml\"/>",
    "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>",
    "<Override PartName=\"/docProps/core.xml\" ",
    "ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>",
    "<Override PartName=\"/docProps/app.xml\" ",
    "ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>",
    "<Override PartName=\"/xl/workbook.xml\" ",
    "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>",
    "<Override PartName=\"/xl/styles.xml\" ",
    "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>",
    "<Override PartName=\"/xl/sharedStrings.xml\" ",
    "ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>",
);
const CONTENT_TYPES_END: &str = "</Types>";

const PACKAGE_RELATIONS: &str = concat!(
    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">",
    "<Relationship Id=\"rId1\" ",
    "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" ",
    "Target=\"xl/workbook.xml\"/>",
    "<Relationship Id=\"rId2\" ",
    "Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\" ",
    "Target=\"docProps/core.xml\"/>",
    "<Relationship Id=\"rId3\" ",
    "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\" ",
    "Target=\"docProps/app.xml\"/>",
    "</Relationships>",
);

const WORKBOOK_START: &str = concat!(
    "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"",
    " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"",
    " xmlns:mx=\"http://schemas.microsoft.com/office/mac/excel/2008/main\"",
    " xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"",
    " xmlns:mv=\"urn:schemas-microsoft-com:mac:vml\"",
    " xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"",
    " xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\"",
    " xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\"",
    " xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">",
    "<workbookPr/><sheets>",
);
const WORKBOOK_END: &str = "</sheets><definedNames/><calcPr/></workbook>";

const WORKBOOK_RELATIONS_START: &str = concat!(
    "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">",
    "<Relationship Id=\"rId1\" ",
    "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\" ",
    "Target=\"styles.xml\"/>",
    "<Relationship Id=\"rId2\" ",
    "Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\" ",
    "Target=\"sharedStrings.xml\"/>",
);
const WORKBOOK_RELATIONS_END: &str = "</Relationships>";

/// The cell styles are, by index: general, date, timestamp, time, timestamp with milliseconds, boolean.
/// Spreadsheets show at most milliseconds, whatever the precision of the number.
const STYLES: &str = concat!(
    "\n\t<styleSheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\">",
    "\n\t\t<numFmts count=\"6\">",
    "\n\t\t    <numFmt formatCode=\"General\" numFmtId=\"164\"/>",
    "\n\t\t    <numFmt formatCode=\"dd/mm/yy\" numFmtId=\"165\"/>",
    "\n\t\t    <numFmt formatCode=\"dd/mm/yyyy\\ hh:mm:ss\" numFmtId=\"166\"/>",
    "\n\t\t    <numFmt formatCode=\"hh:mm:ss\" numFmtId=\"167\"/>",
    "\n\t\t\t<numFmt formatCode=\"dd/mm/yyyy\\ hh:mm:ss.000\" numFmtId=\"168\"/>",
    "\n\t\t\t<numFmt formatCode=\"&quot;TRUE&quot;;&quot;TRUE&quot;;&quot;FALSE&quot;\" numFmtId=\"169\"/>",
    "\n\t\t</numFmts>",
    "\n\t\t<fonts count=\"1\">",
    "\n\t\t\t<font>",
    "\n\t\t\t\t<name val=\"Arial\"/>",
    "\n\t\t\t\t<family val=\"2\"/>",
    "\n\t\t\t\t<sz val=\"12\"/>",
    "\n\t\t\t</font>",
    "\n\t\t</fonts>",
    "\n\t\t<fills count=\"1\">",
    "\n\t\t\t<fill>",
    "\n\t\t\t\t<patternFill patternType=\"none\"/>",
    "\n\t\t\t</fill>",
    "\n\t\t</fills>",
    "\n\t\t<borders count=\"1\">",
    "\n\t\t\t<border diagonalDown=\"false\" diagonalUp=\"false\">",
    "\n\t\t\t\t<left/>",
    "\n\t\t\t\t<right/>",
    "\n\t\t\t\t<top/>",
    "\n\t\t\t\t<bottom/>",
    "\n\t\t\t\t<diagonal/>",
    "\n\t\t\t</border>",
    "\n\t\t</borders>",
    "\n\t\t<cellStyleXfs count=\"1\">",
    "\n\t\t\t<xf numFmtId=\"164\"></xf>",
    "\n\t\t</cellStyleXfs>",
    "\n\t\t<cellXfs count=\"6\">",
    "\n\t\t\t<xf numFmtId=\"164\" xfId=\"0\"/>",
    "\n\t\t\t<xf numFmtId=\"165\" xfId=\"0\"/>",
    "\n\t\t\t<xf numFmtId=\"166\" xfId=\"0\"/>",
    "\n\t\t\t<xf numFmtId=\"167\" xfId=\"0\"/>",
    "\n\t\t\t<xf numFmtId=\"168\" xfId=\"0\"/>",
    "\n\t\t\t<xf numFmtId=\"169\" xfId=\"0\"/>",
    "\n\t\t</cellXfs>",
    "\n\t\t<cellStyles count=\"1\">",
    "\n\t\t\t<cellStyle builtinId=\"0\" customBuiltin=\"false\" name=\"Normal\" xfId=\"0\"/>",
    "\n\t\t</cellStyles>",
    "\n\t</styleSheet>",
    "\n\t",
);

/// Strings are written into their cells, but readers expect the part to exist.
const SHARED_STRINGS: &str = concat!(
    "<sst xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"",
    " count=\"0\" uniqueCount=\"0\"/>",
);

const APP_PROPERTIES: &str = concat!(
    "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"",
    " xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">",
    "<Application>DuckDB</Application>",
    "<TotalTime>0</TotalTime>",
    "</Properties>",
);

const CORE_PROPERTIES: &str = concat!(
    "<cp:coreProperties",
    " xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\"",
    " xmlns:dc=\"http://purl.org/dc/elements/1.1/\"",
    " xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\"",
    " xmlns:dcterms=\"http://purl.org/dc/terms/\"",
    " xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">",
    "<dcterms:created xsi:type=\"dcterms:W3CDTF\">2024-11-15T13:37:00.00Z</dcterms:created>",
    "<dc:creator>DuckDB</dc:creator>",
    "<cp:lastModifiedBy>DuckDB</cp:lastModifiedBy>",
    "<dcterms:modified xsi:type=\"dcterms:W3CDTF\">2024-11-15T13:37:00.00Z</dcterms:modified>",
    "<cp:revision>1</cp:revision>",
    "</cp:coreProperties>",
);

/// The `Override` element that declares the worksheet in `file_name` in `[Content_Types].xml`.
pub fn worksheet_content_type(file_name: &str) -> String {
    format!(
        "<Override PartName=\"/xl/worksheets/{file_name}\" \
         ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
    )
}

/// The `Relationship` element that leads from the workbook to the worksheet in `file_name`.
pub fn worksheet_relation(relation: u64, file_name: &str) -> String {
    format!(
        "<Relationship Id=\"rId{relation}\" \
         Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" \
         Target=\"worksheets/{file_name}\"/>"
    )
}

/// The `sheet` element of `xl/workbook.xml` for the sheet called `name`.
pub fn sheet_element(name: &str, sheet_id: u64, relation: u64) -> String {
    format!(
        "<sheet name=\"{}\" state=\"visible\" sheetId=\"{sheet_id}\" r:id=\"rId{relation}\"/>",
        escape_xml(name)
    )
}

/// Writes every part of a new workbook but its sheet, which is the part [`FIRST_SHEET_FILE`] and already
/// written; `sheet_name` is what the workbook calls the sheet.
pub fn write_workbook_parts<W: Write>(zip: &mut ZipWriter<W>, sheet_name: &str) -> io::Result<()> {
    let content_types = [
        CONTENT_TYPES_START,
        &worksheet_content_type(FIRST_SHEET_FILE),
        CONTENT_TYPES_END,
    ];
    let workbook = [
        WORKBOOK_START,
        &sheet_element(sheet_name, FIRST_SHEET_ID, FIRST_SHEET_RELATION),
        WORKBOOK_END,
    ];
    let workbook_relations = [
        WORKBOOK_RELATIONS_START,
        &worksheet_relation(FIRST_SHEET_RELATION, FIRST_SHEET_FILE),
        WORKBOOK_RELATIONS_END,
    ];
    let parts: [(&str, &[&str]); 8] = [
        (CONTENT_TYPES_PATH, &content_types),
        ("_rels/.rels", &[PACKAGE_RELATIONS]),
        (WORKBOOK_PATH, &workbook),
        (WORKBOOK_RELATIONS_PATH, &workbook_relations),
        (STYLES_PATH, &[STYLES]),
        ("xl/sharedStrings.xml", &[SHARED_STRINGS]),
        ("docProps/app.xml", &[APP_PROPERTIES]),
        ("docProps/core.xml", &[CORE_PROPERTIES]),
    ];
    for (path, pieces) in parts {
        zip.begin_file(path)?;
        zip.write_all(XML_DECLARATION.as_bytes())?;
        for piece in pieces {
            zip.write_all(piece.as_bytes())?;
        }
        zip.end_file()?;
    }
    Ok(())
}
