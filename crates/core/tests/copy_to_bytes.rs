//! The exact bytes of the workbooks `COPY ... TO` writes. Whoever reads these workbooks, be it a spreadsheet
//! or a test that compares files, relies on them: the same rows must always give the same parts.

mod support;

use support::{
    STYLES, TempDir, Workbook, connect, execute, inline, row, workbook_part, workbook_parts,
};

const XML_DECLARATION: &str = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";

/// A row with a value of every kind of cell and a row of empty cells but for a boolean and a number.
const ROWS: &str = "SELECT * EXCLUDE (o) FROM (
        SELECT 1 AS o, 42 AS \"int\", 1.5::DOUBLE AS dbl, 'a<b>&\"c''' AS txt, true AS flag,
               DATE '2024-03-15' AS day, TIME '08:15:00' AS tod, TIMESTAMP '2024-03-15 12:30:45.123' AS ts,
               TIMESTAMP_S '2024-03-15 12:30:45' AS ts_s, NULL::INTEGER AS nothing
        UNION ALL
        SELECT 2, NULL, NULL, NULL, false, NULL, NULL, NULL, NULL, 7
    ) ORDER BY o";

/// The parts of the workbook of [`ROWS`] with a header in the sheet `Pin & <Sheet>`, without the XML
/// declaration each of them starts with.
const NEW_WORKBOOK: [(&str, &str); 9] = [
    (
        "xl/worksheets/sheet1.xml",
        concat!(
            "\n\t<worksheet xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"",
            "\n\t          ",
            " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"",
            "\n\t           xmlns:mx=\"http://schemas.microsoft.com/office/mac/excel/2008/main\"",
            "\n\t           xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"",
            "\n\t           xmlns:mv=\"urn:schemas-microsoft-com:mac:vml\"",
            "\n\t           xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"",
            "\n\t          ",
            " xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\"",
            "\n\t           xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\"",
            "\n\t           xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\">",
            "\n\t<sheetData>",
            "\n\t<row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>int</t></is></c>",
            "<c r=\"B1\" t=\"inlineStr\"><is><t>dbl</t></is></c><c r=\"C1\" t=\"inlineStr\"><is><t>",
            "txt</t></is></c><c r=\"D1\" t=\"inlineStr\"><is><t>flag</t></is></c>",
            "<c r=\"E1\" t=\"inlineStr\"><is><t>day</t></is></c><c r=\"F1\" t=\"inlineStr\"><is><t>",
            "tod</t></is></c><c r=\"G1\" t=\"inlineStr\"><is><t>ts</t></is></c>",
            "<c r=\"H1\" t=\"inlineStr\"><is><t>ts_s</t></is></c><c r=\"I1\" t=\"inlineStr\"><is><t>",
            "nothing</t></is></c></row><row r=\"2\"><c r=\"A2\"><v>42</v></c><c r=\"B2\"><v>1.5</v></c>",
            "<c r=\"C2\" t=\"inlineStr\"><is><t>a&lt;b&gt;&amp;&quot;c&apos;</t></is></c>",
            "<c r=\"D2\" t=\"b\" s=\"5\"><v>1</v></c><c r=\"E2\" s=\"1\"><v>45366.0</v></c>",
            "<c r=\"F2\" s=\"3\"><v>0.34375</v></c><c r=\"G2\" s=\"4\"><v>45366.521355590274</v></c>",
            "<c r=\"H2\" s=\"2\"><v>45366.52135416667</v></c></row><row r=\"3\">",
            "<c r=\"D3\" t=\"b\" s=\"5\"><v>0</v></c><c r=\"I3\"><v>7</v></c></row></sheetData>",
            "</worksheet>",
        ),
    ),
    (
        "[Content_Types].xml",
        concat!(
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">",
            "<Default Extension=\"xml\" ContentType=\"application/xml\"/><Default Extension=\"rels\"",
            " ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>",
            "<Override PartName=\"/docProps/core.xml\"",
            " ContentType=\"application/vnd.openxmlformats-package.core-properties+xml\"/>",
            "<Override PartName=\"/docProps/app.xml\"",
            " ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>",
            "<Override PartName=\"/xl/workbook.xml\"",
            " ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml\"/>",
            "<Override PartName=\"/xl/styles.xml\"",
            " ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml\"/>",
            "<Override PartName=\"/xl/sharedStrings.xml\"",
            " ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml\"/>",
            "<Override PartName=\"/xl/worksheets/sheet1.xml\"",
            " ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>",
            "</Types>",
        ),
    ),
    (
        "_rels/.rels",
        concat!(
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">",
            "<Relationship Id=\"rId1\"",
            " Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\"",
            " Target=\"xl/workbook.xml\"/><Relationship Id=\"rId2\"",
            " Type=\"http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties\"",
            " Target=\"docProps/core.xml\"/><Relationship Id=\"rId3\"",
            " Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties\"",
            " Target=\"docProps/app.xml\"/></Relationships>",
        ),
    ),
    (
        "xl/workbook.xml",
        concat!(
            "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\"",
            " xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"",
            " xmlns:mx=\"http://schemas.microsoft.com/office/mac/excel/2008/main\"",
            " xmlns:mc=\"http://schemas.openxmlformats.org/markup-compatibility/2006\"",
            " xmlns:mv=\"urn:schemas-microsoft-com:mac:vml\"",
            " xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\"",
            " xmlns:x15=\"http://schemas.microsoft.com/office/spreadsheetml/2010/11/main\"",
            " xmlns:x14ac=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac\"",
            " xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\"><workbookPr/><sheets>",
            "<sheet name=\"Pin &amp; &lt;Sheet&gt;\" state=\"visible\" sheetId=\"1\" r:id=\"rId3\"/>",
            "</sheets><definedNames/><calcPr/></workbook>",
        ),
    ),
    (
        "xl/_rels/workbook.xml.rels",
        concat!(
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">",
            "<Relationship Id=\"rId1\"",
            " Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles\"",
            " Target=\"styles.xml\"/><Relationship Id=\"rId2\"",
            " Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings\"",
            " Target=\"sharedStrings.xml\"/><Relationship Id=\"rId3\"",
            " Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\"",
            " Target=\"worksheets/sheet1.xml\"/></Relationships>",
        ),
    ),
    (
        "xl/styles.xml",
        concat!(
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
        ),
    ),
    (
        "xl/sharedStrings.xml",
        concat!(
            "<sst",
            " xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" count=\"0\" uniqueCount=\"0\"/>",
        ),
    ),
    (
        "docProps/app.xml",
        concat!(
            "<Properties",
            " xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"",
            " xmlns:vt=\"http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes\">",
            "<Application>DuckDB</Application><TotalTime>0</TotalTime></Properties>",
        ),
    ),
    (
        "docProps/core.xml",
        concat!(
            "<cp:coreProperties",
            " xmlns:cp=\"http://schemas.openxmlformats.org/package/2006/metadata/core-properties\"",
            " xmlns:dc=\"http://purl.org/dc/elements/1.1/\"",
            " xmlns:dcmitype=\"http://purl.org/dc/dcmitype/\"",
            " xmlns:dcterms=\"http://purl.org/dc/terms/\"",
            " xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\">",
            "<dcterms:created xsi:type=\"dcterms:W3CDTF\">2024-11-15T13:37:00.00Z</dcterms:created>",
            "<dc:creator>DuckDB</dc:creator><cp:lastModifiedBy>DuckDB</cp:lastModifiedBy>",
            "<dcterms:modified xsi:type=\"dcterms:W3CDTF\">2024-11-15T13:37:00.00Z</dcterms:modified>",
            "<cp:revision>1</cp:revision></cp:coreProperties>",
        ),
    ),
];

fn new_workbook_part(name: &str) -> String {
    let (_, content) = NEW_WORKBOOK
        .iter()
        .find(|(part, _)| *part == name)
        .unwrap_or_else(|| panic!("no part {name}"));
    format!("{XML_DECLARATION}{content}")
}

fn part_names(path: &str) -> Vec<String> {
    workbook_parts(path)
        .into_iter()
        .map(|(name, _)| name)
        .collect()
}

#[test]
fn a_new_workbook() {
    let connection = connect();
    let directory = TempDir::new("bytes_new");
    let path = directory.path("pin.xlsx");
    execute(
        &connection,
        &format!("COPY ({ROWS}) TO '{path}' (FORMAT xlsx, HEADER true, SHEET 'Pin & <Sheet>')"),
    );
    let parts = workbook_parts(&path);
    let names: Vec<&str> = parts.iter().map(|(name, _)| name.as_str()).collect();
    assert_eq!(names, NEW_WORKBOOK.map(|(name, _)| name));
    for (name, content) in &parts {
        assert_eq!(*content, new_workbook_part(name), "{name}");
    }
    assert_eq!(directory.file_names(), ["pin.xlsx"]);
}

#[test]
fn a_sheet_appended_to_a_workbook_of_this_writer() {
    let connection = connect();
    let directory = TempDir::new("bytes_append");
    let path = directory.path("pin.xlsx");
    execute(
        &connection,
        &format!("COPY ({ROWS}) TO '{path}' (FORMAT xlsx, HEADER true, SHEET 'Pin & <Sheet>')"),
    );
    execute(
        &connection,
        &format!(
            "COPY (SELECT DATE '2024-03-16' AS d, true AS b, 'x' AS s) TO '{path}' \
             (FORMAT xlsx, SHEET 'More \"& more\"', MODE 'append')"
        ),
    );

    // What stays comes first and in its old order, then the new sheet, then what had to change.
    assert_eq!(
        part_names(&path),
        [
            "xl/worksheets/sheet1.xml",
            "_rels/.rels",
            "xl/styles.xml",
            "xl/sharedStrings.xml",
            "docProps/app.xml",
            "docProps/core.xml",
            "xl/worksheets/sheet2.xml",
            "[Content_Types].xml",
            "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels",
        ]
    );
    for kept in [
        "xl/worksheets/sheet1.xml",
        "_rels/.rels",
        "xl/styles.xml",
        "xl/sharedStrings.xml",
        "docProps/app.xml",
        "docProps/core.xml",
    ] {
        assert_eq!(
            workbook_part(&path, kept),
            new_workbook_part(kept),
            "{kept}"
        );
    }
    let new_sheet = workbook_part(&path, "xl/worksheets/sheet2.xml");
    let first_sheet = new_workbook_part("xl/worksheets/sheet1.xml");
    let (frame, _) = first_sheet.split_once("<row").unwrap();
    assert_eq!(
        new_sheet,
        format!(
            "{frame}{}",
            concat!(
                r#"<row r="1"><c r="A1" s="1"><v>45367.0</v></c><c r="B1" t="b" s="5"><v>1</v></c>"#,
                r#"<c r="C1" t="inlineStr"><is><t>x</t></is></c></row></sheetData></worksheet>"#
            )
        )
    );
    assert_eq!(
        workbook_part(&path, "[Content_Types].xml"),
        new_workbook_part("[Content_Types].xml").replace(
            "</Types>",
            concat!(
                r#"<Override PartName="/xl/worksheets/sheet2.xml" ContentType=""#,
                r#"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/></Types>"#
            )
        )
    );
    assert_eq!(
        workbook_part(&path, "xl/workbook.xml"),
        new_workbook_part("xl/workbook.xml").replace(
            "</sheets>",
            concat!(
                r#"<sheet name="More &quot;&amp; more&quot;" state="visible" sheetId="2" r:id="rId4"/>"#,
                "</sheets>"
            )
        )
    );
    assert_eq!(
        workbook_part(&path, "xl/_rels/workbook.xml.rels"),
        new_workbook_part("xl/_rels/workbook.xml.rels").replace(
            "</Relationships>",
            concat!(
                r#"<Relationship Id="rId4" Type=""#,
                r#"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet""#,
                r#" Target="worksheets/sheet2.xml"/></Relationships>"#
            )
        )
    );
}

#[test]
fn a_sheet_replaced_in_a_workbook_of_this_writer() {
    let connection = connect();
    let directory = TempDir::new("bytes_replace");
    let path = directory.path("pin.xlsx");
    execute(
        &connection,
        &format!("COPY ({ROWS}) TO '{path}' (FORMAT xlsx, HEADER true, SHEET 'Pin & <Sheet>')"),
    );
    execute(
        &connection,
        &format!(
            "COPY (SELECT 1 AS a) TO '{path}' (FORMAT xlsx, SHEET 'pin & <sheet>', MODE 'replace')"
        ),
    );
    assert_eq!(
        part_names(&path),
        [
            "[Content_Types].xml",
            "_rels/.rels",
            "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels",
            "xl/styles.xml",
            "xl/sharedStrings.xml",
            "docProps/app.xml",
            "docProps/core.xml",
            "xl/worksheets/sheet1.xml",
        ]
    );
    for (name, content) in workbook_parts(&path) {
        if name == "xl/worksheets/sheet1.xml" {
            assert!(
                content.ends_with(
                    "<sheetData>\n\t<row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData></worksheet>"
                ),
                "{content}"
            );
        } else {
            assert_eq!(content, new_workbook_part(&name), "{name}");
        }
    }
}

#[test]
fn a_sheet_appended_to_a_workbook_with_other_styles() {
    let connection = connect();
    let workbook = Workbook::new()
        .sheet("Data", &row(1, &[inline("A1", "kept")]))
        .write("bytes_styles.xlsx");
    let path = workbook.path();
    execute(
        &connection,
        &format!(
            "COPY (SELECT DATE '2024-03-16' AS d, TIME '01:00:00' AS t, TIMESTAMP '2024-03-16 01:00:00' AS ts, \
             TIMESTAMP_S '2024-03-16 01:00:00' AS tss, false AS b) TO '{path}' \
             (FORMAT xlsx, SHEET 'Typed', MODE 'append')"
        ),
    );
    let styles = workbook_part(path, "xl/styles.xml");
    let expected = STYLES
        .replace(
            r#"<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd hh:mm:ss"/>"#,
            concat!(
                r#"<numFmts count="6"><numFmt numFmtId="164" formatCode="yyyy-mm-dd hh:mm:ss"/>"#,
                r#"<numFmt formatCode="dd/mm/yy" numFmtId="165"/>"#,
                r#"<numFmt formatCode="dd/mm/yyyy\ hh:mm:ss" numFmtId="166"/>"#,
                r#"<numFmt formatCode="hh:mm:ss" numFmtId="167"/>"#,
                r#"<numFmt formatCode="dd/mm/yyyy\ hh:mm:ss.000" numFmtId="168"/>"#,
                r#"<numFmt formatCode="&quot;TRUE&quot;;&quot;TRUE&quot;;&quot;FALSE&quot;" numFmtId="169"/>"#,
            ),
        )
        .replace(
            r#"<cellXfs count="5">"#,
            r#"<cellXfs count="10">"#,
        )
        .replace(
            "</cellXfs>",
            concat!(
                r#"<xf numFmtId="165" xfId="0"/><xf numFmtId="166" xfId="0"/><xf numFmtId="167" xfId="0"/>"#,
                r#"<xf numFmtId="168" xfId="0"/><xf numFmtId="169" xfId="0"/></cellXfs>"#,
            ),
        );
    assert_eq!(styles, expected);
    let sheet = workbook_part(path, "xl/worksheets/sheet2.xml");
    assert!(
        sheet.ends_with(concat!(
            r#"<row r="1"><c r="A1" s="5"><v>45367.0</v></c><c r="B1" s="7"><v>0.041666666666666664</v></c>"#,
            r#"<c r="C1" s="8"><v>45367.04166666667</v></c><c r="D1" s="6"><v>45367.04166666667</v></c>"#,
            r#"<c r="E1" t="b" s="9"><v>0</v></c></row></sheetData></worksheet>"#
        )),
        "{sheet}"
    );
    assert_eq!(
        part_names(path),
        [
            "xl/worksheets/sheet1.xml",
            "xl/worksheets/sheet2.xml",
            "[Content_Types].xml",
            "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels",
            "xl/styles.xml",
        ]
    );
}
