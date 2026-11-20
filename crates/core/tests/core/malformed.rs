//! Broken workbooks give errors, never panics.

use crate::support::{
    CONTENT_TYPES, TempFile, Workbook, connect, inline, number, numbered_rows, query, query_error,
    row, shared, try_query,
};
use duckdb_neo::error::DuckDBError;

fn simple() -> Workbook {
    Workbook::new()
        .sheet(
            "S",
            &[
                row(1, &[inline("A1", "id"), shared("B1", 0)]),
                row(2, &[number("A2", 1), shared("B2", 1)]),
                row(3, &[number("A3", 2), inline("B3", "inline")]),
            ]
            .concat(),
        )
        .shared_strings(&["name", "shared"])
}

fn read_error(book: &Workbook, name: &str) -> String {
    let connection = connect();
    let file = book.write(name);
    query_error(&connection, &format!("FROM read_xlsx('{}')", file.path()))
}

#[test]
fn the_simple_workbook_reads() {
    let connection = connect();
    let file = simple().write("simple.xlsx");
    assert_eq!(
        query(&connection, &format!("FROM read_xlsx('{}')", file.path())),
        ["1.0|shared", "2.0|inline"]
    );
}

#[test]
fn files_that_are_not_workbooks() {
    let connection = connect();
    let file = TempFile::new("not_a_workbook.xlsx");
    for content in [&b""[..], b"id,name\n1,a\n", &[0xFFu8; 3000]] {
        std::fs::write(file.path(), content).unwrap();
        let error = query_error(&connection, &format!("FROM read_xlsx('{}')", file.path()));
        assert!(error.starts_with("IO Error: "), "{error}");
    }
    let directory = std::env::temp_dir();
    let error = query_error(
        &connection,
        &format!("FROM read_xlsx('{}')", directory.display()),
    );
    assert!(error.starts_with("IO Error: "), "{error}");
}

#[test]
fn missing_parts() {
    for (part, error) in [
        (
            "[Content_Types].xml",
            "Binder Error: No [Content_Types].xml found in xlsx file",
        ),
        (
            "xl/workbook.xml",
            "Binder Error: No xl/workbook.xml found in xlsx file",
        ),
        (
            "xl/_rels/workbook.xml.rels",
            "Binder Error: No xl/_rels/workbook.xml.rels found in xlsx file",
        ),
        (
            "xl/worksheets/sheet1.xml",
            "Binder Error: Sheet 'xl/worksheets/sheet1.xml' not found in xlsx file",
        ),
        (
            "xl/sharedStrings.xml",
            "Binder Error: No shared strings found in xlsx file",
        ),
    ] {
        assert_eq!(
            read_error(&simple().without(part), "missing_part.xlsx"),
            error,
            "{part}"
        );
    }
    let connection = connect();
    let file = simple().without("xl/styles.xml").write("no_styles.xlsx");
    assert_eq!(
        query(
            &connection,
            &format!("SELECT count(*) FROM read_xlsx('{}')", file.path())
        ),
        ["2"]
    );
}

#[test]
fn parts_without_what_they_must_have() {
    let relationship = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    for (part, content, error) in [
        (
            "[Content_Types].xml",
            "<Types><Override PartName=\"/xl/workbook.xml\"/></Types>".to_owned(),
            "Invalid Input Error: Invalid content type entry in [Content_Types].xml",
        ),
        (
            "xl/workbook.xml",
            "<workbook><sheets><sheet name=\"S\"/></sheets></workbook>".to_owned(),
            "Invalid Input Error: Invalid sheet entry in workbook.xml",
        ),
        (
            "xl/_rels/workbook.xml.rels",
            "<Relationships><Relationship Id=\"rId1\" Target=\"x\"/></Relationships>".to_owned(),
            "Invalid Input Error: Invalid relationship entry in _rels/.rels",
        ),
        (
            "xl/workbook.xml",
            "<workbook><sheets/></workbook>".to_owned(),
            "Binder Error: No sheets found in xlsx file (is the file corrupt?)",
        ),
        (
            "xl/workbook.xml",
            String::new(),
            "Binder Error: No sheets found in xlsx file (is the file corrupt?)",
        ),
        (
            "xl/_rels/workbook.xml.rels",
            format!("<Relationships><Relationship Id=\"rId1\" Type=\"{relationship}/chartsheet\" Target=\"worksheets/sheet1.xml\"/></Relationships>"),
            "Binder Error: No sheets found in xlsx file (is the file corrupt?)",
        ),
        (
            "xl/styles.xml",
            "<styleSheet><numFmts><numFmt formatCode=\"0\"/></numFmts></styleSheet>".to_owned(),
            "Invalid Input Error: Invalid numFmt entry in styles.xml",
        ),
        (
            "xl/worksheets/sheet1.xml",
            "<worksheet><sheetData><row r=\"1\"><c r=\"a1\"><v>1</v></c></row></sheetData></worksheet>".to_owned(),
            "Invalid Input Error: Invalid cell reference in sheet: a1",
        ),
        (
            "xl/worksheets/sheet1.xml",
            "<worksheet><sheetData><row r=\"1\"><c r=\"A2\"><v>1</v></c></row></sheetData></worksheet>".to_owned(),
            "Invalid Input Error: Cell reference does not match row reference in sheet",
        ),
        (
            "xl/worksheets/sheet1.xml",
            "<worksheet><sheetData/></worksheet>".to_owned(),
            "Binder Error: No rows found in xlsx file",
        ),
        (
            "xl/worksheets/sheet1.xml",
            String::new(),
            "Binder Error: No rows found in xlsx file",
        ),
    ] {
        assert_eq!(read_error(&simple().part(part, content), "bad_part.xlsx"), error, "{part}");
    }
}

#[test]
fn malformed_xml_is_an_io_error() {
    let sheet = |body: &str| format!("<worksheet><sheetData>{body}</sheetData></worksheet>");
    let cases = [
        (
            "xl/workbook.xml",
            "<workbook><sheets><sheet name=\"S\" r:id=\"rId1\"></sheets></workbook>".to_owned(),
        ),
        ("xl/workbook.xml", "<workbook><sheets>".to_owned()),
        ("xl/workbook.xml", "  ".to_owned()),
        ("[Content_Types].xml", CONTENT_TYPES.replace("</Types>", "")),
        (
            "xl/_rels/workbook.xml.rels",
            "<Relationships><Relationship Id=\"a\" Id=\"b\"/></Relationships>".to_owned(),
        ),
        (
            "xl/styles.xml",
            "<styleSheet><cellXfs><xf numFmtId=\"0\"></cellXfs></styleSheet>".to_owned(),
        ),
        (
            "xl/sharedStrings.xml",
            "<sst><si><t>name</t></si><si><t>a &nbsp; b</t></si></sst>".to_owned(),
        ),
        (
            "xl/sharedStrings.xml",
            "<sst><si><t>name</t></si><si><t>truncated".to_owned(),
        ),
        (
            "xl/worksheets/sheet1.xml",
            sheet("<row r=\"1\"><c r=\"A1\"><v>1</v></x></row>"),
        ),
        (
            "xl/worksheets/sheet1.xml",
            sheet("<row r=\"1\"><c r=\"A1\"><v>1</v></c>"),
        ),
        (
            "xl/worksheets/sheet1.xml",
            sheet("<row r=\"1\"><c r=\"A1\"><v>&unknown;</v></c></row>"),
        ),
        (
            "xl/worksheets/sheet1.xml",
            sheet("<row r=\"1\"><c r=\"A1\" r=\"A1\"><v>1</v></c></row>"),
        ),
        (
            "xl/worksheets/sheet1.xml",
            sheet("<row r=\"1\"><c r=\"A1\"><v>\u{1}</v></c></row>"),
        ),
        (
            "xl/worksheets/sheet1.xml",
            "<?xml version=\"1.0\" encoding=\"UTF-16\"?><worksheet/>".to_owned(),
        ),
    ];
    for (part, content) in cases {
        let error = read_error(&simple().part(part, &content), "malformed.xml.xlsx");
        assert!(
            error.starts_with("IO Error: XML parse error"),
            "{part} {content}: {error}"
        );
    }
    let error = read_error(
        &simple().part(
            "xl/worksheets/sheet1.xml",
            "<?xml version=\"1.0\" encoding=\"windows-1252\"?><worksheet/>",
        ),
        "other_encoding.xlsx",
    );
    assert_eq!(
        error,
        "IO Error: XML parse error: unsupported encoding 'windows-1252', only UTF-8, UTF-16 and ISO-8859-1 are \
         supported"
    );
}

#[test]
fn malformed_xml_after_the_rows_that_are_read_is_not_found() {
    let connection = connect();
    let sheet = format!(
        "<worksheet><sheetData>{}{}</sheetData><broken></worksheet>",
        row(1, &[inline("A1", "id"), inline("B1", "label")]),
        numbered_rows(2, 3000)
    );
    let file = simple()
        .part("xl/worksheets/sheet1.xml", &sheet)
        .part("xl/workbook.xml", "<workbook xmlns:r=\"r\"><sheets><sheet name=\"S\" r:id=\"rId1\"/></sheets></workbook><broken")
        .write("broken_tail.xlsx");
    assert_eq!(
        query(
            &connection,
            &format!("SELECT count(*) FROM read_xlsx('{}')", file.path())
        ),
        ["2999"]
    );

    let late = format!(
        "<worksheet><sheetData>{}{}<row r=\"3001\"><c></x></row></sheetData></worksheet>",
        row(1, &[inline("A1", "id"), inline("B1", "label")]),
        numbered_rows(2, 3000)
    );
    let file = simple()
        .part("xl/worksheets/sheet1.xml", &late)
        .write("broken_late.xlsx");
    let path = file.path();
    assert_eq!(
        query(
            &connection,
            &format!("SELECT column_name FROM (DESCRIBE FROM read_xlsx('{path}'))")
        ),
        ["id", "label"]
    );
    assert_eq!(
        query(
            &connection,
            &format!("SELECT id FROM read_xlsx('{path}') LIMIT 1")
        ),
        ["2.0"]
    );
    let error = query_error(
        &connection,
        &format!("SELECT count(*) FROM read_xlsx('{path}')"),
    );
    assert!(error.starts_with("IO Error: XML parse error"), "{error}");
}

#[test]
fn cell_text_has_a_limit() {
    let connection = connect();
    let cell = |length: usize| {
        simple().sheet(
            "Long",
            &[
                row(1, &[inline("A1", "text")]),
                row(2, &[inline("A2", &"x".repeat(length))]),
            ]
            .concat(),
        )
    };
    let file = cell(65534).write("cell_limit.xlsx");
    assert_eq!(
        query(
            &connection,
            &format!(
                "SELECT length(text) FROM read_xlsx('{}', sheet = 'Long')",
                file.path()
            )
        ),
        ["65534"]
    );
    for length in [65535, 40_000_000] {
        let file = cell(length).write("cell_over_limit.xlsx");
        let error = query_error(
            &connection,
            &format!("FROM read_xlsx('{}', sheet = 'Long')", file.path()),
        );
        assert_eq!(
            error,
            "Invalid Input Error: XLSX: Cell data too large (is the file corrupted?)"
        );
    }
}

#[test]
fn shared_strings_that_do_not_exist() {
    let connection = connect();
    let book = Workbook::new()
        .sheet(
            "S",
            &[row(1, &[shared("A1", 0)]), row(2, &[shared("A2", 5)])].concat(),
        )
        .shared_strings(&["name"]);
    let file = book.write("shared_out_of_range.xlsx");
    assert_eq!(
        query_error(&connection, &format!("FROM read_xlsx('{}')", file.path())),
        "Invalid Input Error: Cell 'A2' refers to shared string '5', which the xlsx file does not have"
    );
    let header = Workbook::new()
        .sheet(
            "S",
            &[row(1, &[shared("A1", 3)]), row(2, &[number("A2", 5)])].concat(),
        )
        .shared_strings(&["name"]);
    assert_eq!(
        read_error(&header, "shared_header_out_of_range.xlsx"),
        "Binder Error: Shared string 3 of the header not found in xlsx file"
    );
}

#[test]
fn a_part_that_does_not_match_its_checksum() {
    let connection = connect();
    let mut bytes = simple().bytes();
    let name = b"xl/worksheets/sheet1.xml";
    let in_directory = bytes
        .windows(name.len())
        .rposition(|window| window == name)
        .expect("the central directory names the sheet");
    // The checksum sits 30 bytes before the name of a central directory entry.
    bytes[in_directory - 30] ^= 0xFF;
    let file = TempFile::new("checksum.xlsx");
    std::fs::write(file.path(), bytes).unwrap();
    let error = query_error(&connection, &format!("FROM read_xlsx('{}')", file.path()));
    assert!(error.starts_with("IO Error: "), "{error}");
    assert!(error.contains("checksum"), "{error}");
}

/// A small deterministic generator, so that failures can be reproduced.
struct Random(u64);

impl Random {
    fn next(&mut self) -> u64 {
        self.0 ^= self.0 << 13;
        self.0 ^= self.0 >> 7;
        self.0 ^= self.0 << 17;
        self.0
    }

    fn below(&mut self, bound: usize) -> usize {
        (self.next() % bound as u64) as usize
    }
}

fn never_internal(
    connection: &duckdb_neo::connection::Connection,
    path: &str,
    what: &str,
) -> usize {
    let mut readable_ways = 0;
    for arguments in [
        "",
        ", header = false, all_varchar = true",
        ", range = 'A1:C5', stop_at_empty = false",
    ] {
        match try_query(connection, &format!("FROM read_xlsx('{path}'{arguments})")) {
            Ok(_) => readable_ways += 1,
            Err(error) => {
                assert_ne!(
                    error.code,
                    DuckDBError::RuntimeInternal,
                    "{what}: {}",
                    error.message
                );
            }
        }
    }
    readable_ways
}

#[test]
fn corrupted_parts_never_panic() {
    let connection = connect();
    let mut random = Random(0x9E37_79B9_7F4A_7C15);
    let file = TempFile::new("mutated_parts.xlsx");
    let sheet_data = [
        row(
            1,
            &[inline("A1", "id"), shared("B1", 0), inline("C1", "when")],
        ),
        row(
            2,
            &[
                number("A2", 1),
                shared("B2", 1),
                crate::support::styled("C2", 44197.5, 3),
            ],
        ),
        row(
            4,
            &[
                number("A4", 2),
                inline("B4", "a &amp; b"),
                crate::support::styled("C4", 1, 1),
            ],
        ),
    ]
    .concat();
    let book = || {
        Workbook::new()
            .sheet("S", &sheet_data)
            .shared_strings(&["name", "shared"])
    };
    let mut pristine = excel_zip::ZipReader::new(std::io::Cursor::new(book().bytes())).unwrap();
    let mut parts: Vec<(String, Vec<u8>)> = Vec::new();
    for index in 0..pristine.entries().len() {
        let name = pristine.entries()[index].name.clone();
        let mut content = Vec::new();
        pristine.open_entry(index).unwrap();
        std::io::Read::read_to_end(&mut pristine, &mut content).unwrap();
        pristine.close_entry().unwrap();
        parts.push((name, content));
    }
    let fragments: [&[u8]; 14] = [
        b"<",
        b">",
        b"\"",
        b"&",
        b"/",
        b"<c r=\"ZZ9\">",
        b"</row>",
        b"\0",
        b"\xff",
        b"t=\"s\"",
        b"r=\"0\"",
        b"-1",
        b"9",
        b" ",
    ];
    let (mut readable, mut unreadable) = (0, 0);
    for round in 0..1500 {
        let (name, pristine) = &parts[random.below(parts.len())];
        let mut bytes = pristine.clone();
        for _ in 0..=random.below(2) {
            let position = random.below(bytes.len());
            match random.below(4) {
                0 => bytes[position] = b' ' + (random.next() % 95) as u8,
                1 => {
                    let fragment = fragments[random.below(fragments.len())];
                    let end = (position + fragment.len()).min(bytes.len());
                    bytes[position..end].copy_from_slice(&fragment[..end - position]);
                }
                2 => bytes[position] ^= 1 << random.below(8),
                _ => {
                    let other = random.below(bytes.len());
                    bytes.swap(position, other);
                }
            }
        }
        std::fs::write(file.path(), book().part(name, &bytes).bytes()).unwrap();
        match never_internal(
            &connection,
            file.path(),
            &format!("round {round} in {name}"),
        ) {
            0 => unreadable += 1,
            _ => readable += 1,
        }
    }
    // The corruptions are small enough to leave many workbooks readable and many not.
    assert!(
        readable > 150 && unreadable > 150,
        "{readable} readable, {unreadable} unreadable"
    );
}

#[test]
fn truncated_and_garbled_archives_never_panic() {
    let connection = connect();
    let mut random = Random(0xD1B5_4A32_D192_ED03);
    let file = TempFile::new("mutated_archive.xlsx");
    let pristine = simple().bytes();
    for round in 0..300 {
        let mut bytes = pristine.clone();
        match round % 3 {
            0 => bytes.truncate(random.below(bytes.len())),
            1 => {
                let start = random.below(bytes.len());
                let end = (start + random.below(64)).min(bytes.len());
                for byte in &mut bytes[start..end] {
                    *byte = random.next() as u8;
                }
            }
            _ => {
                let position = random.below(bytes.len());
                let garbage: Vec<u8> = (0..random.below(40)).map(|_| random.next() as u8).collect();
                bytes.splice(position..position, garbage);
            }
        }
        std::fs::write(file.path(), &bytes).unwrap();
        never_internal(&connection, file.path(), &format!("round {round}"));
    }
}
