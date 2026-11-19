//! `COPY ... TO 'file.xlsx' (FORMAT xlsx, MODE 'append')` and `MODE 'replace'`: a sheet goes into a workbook
//! that exists, and everything else in the workbook stays as it is.

use std::{io::Write, path::Path};

use crate::support::{
    CONTENT_TYPES, StoredEntry, TempDir, Workbook, archive_entries, connect, execute,
    execute_error, inline, number, produced_fixtures, query, row, sheet_rows, styled, try_query,
    utf16, workbook_part,
};
use duckdb_neo::{Result, connection::Connection};
use excel_core::xlsx::{
    parsers::workbook::WorkbookParser,
    xml::{Attributes, Flow, XmlHandler, XmlReader},
};
use excel_zip::ZipWriter;

/// A column of every kind of cell, with empty cells in the last one.
const TYPED_ROWS: &str = "SELECT i AS n, 'text ' || i AS s, i % 2 = 0 AS flag, DATE '2024-03-15' + i::INTEGER AS day, \
     TIME '08:15:00' + INTERVAL (i) MINUTE AS tod, TIMESTAMP '2024-03-15 12:30:45' + INTERVAL (i) HOUR AS ts, \
     TIMESTAMP_S '2024-03-15 12:30:45' AS ts_s, CASE WHEN i % 2 = 1 THEN i END AS odd FROM range(4) t(i)";

const TYPED_ROWS_READ: [&str; 4] = [
    "0.0|text 0|true|2024-03-15|08:15:00|2024-03-15 12:30:45|2024-03-15 12:30:45|NULL",
    "1.0|text 1|false|2024-03-16|08:16:00|2024-03-15 13:30:45|2024-03-15 12:30:45|1.0",
    "2.0|text 2|true|2024-03-17|08:17:00|2024-03-15 14:30:45|2024-03-15 12:30:45|NULL",
    "3.0|text 3|false|2024-03-18|08:18:00|2024-03-15 15:30:45|2024-03-15 12:30:45|3.0",
];

/// The parts a new sheet has to be named in, and the one that may need more cell styles.
const PATCHED_PARTS: [&str; 4] = [
    "[Content_Types].xml",
    "xl/workbook.xml",
    "xl/_rels/workbook.xml.rels",
    "xl/styles.xml",
];

fn copy(connection: &Connection, rows: &str, path: &str, options: &str) {
    execute(
        connection,
        &format!("COPY ({rows}) TO '{path}' (FORMAT xlsx{options})"),
    );
}

fn copy_error(connection: &Connection, rows: &str, path: &str, options: &str) -> String {
    execute_error(
        connection,
        &format!("COPY ({rows}) TO '{path}' (FORMAT xlsx{options})"),
    )
}

fn read_sheet(connection: &Connection, path: &str, sheet: &str) -> Vec<String> {
    let sheet = sheet.replace('\'', "''");
    query(
        connection,
        &format!("FROM read_xlsx('{path}', sheet = '{sheet}', header = true)"),
    )
}

/// The rows of a sheet as text, or the error reading it gives; what does not depend on the rest of the
/// workbook.
fn sheet_as_text(connection: &Connection, path: &str, sheet: &str) -> Vec<String> {
    let sheet = sheet.replace('\'', "''");
    let sql =
        format!("FROM read_xlsx('{path}', sheet = '{sheet}', header = false, all_varchar = true)");
    try_query(connection, &sql).unwrap_or_else(|error| vec![error.message.replace(path, "")])
}

fn sheet_names(path: &str) -> Vec<String> {
    let workbook = archive_entries(path)
        .into_iter()
        .rfind(|entry| entry.name == "xl/workbook.xml")
        .expect("a workbook part");
    let mut parser = WorkbookParser::new();
    XmlReader::new(workbook.content.as_slice(), None)
        .parse_all(&mut parser)
        .expect("a well-formed workbook part");
    parser.sheets.into_iter().map(|sheet| sheet.name).collect()
}

struct AnyDocument;

impl XmlHandler for AnyDocument {
    fn start_element(&mut self, _name: &str, _attributes: &Attributes<'_>) -> Result<Flow> {
        Ok(Flow::Continue)
    }

    fn end_element(&mut self, _name: &str) -> Result<Flow> {
        Ok(Flow::Continue)
    }
}

fn assert_well_formed(entry: &StoredEntry) {
    let parsed = XmlReader::new(entry.content.as_slice(), None).parse_all(&mut AnyDocument);
    if let Err(error) = parsed {
        panic!("{} is not well-formed: {}", entry.name, error.message);
    }
}

/// The entries that count: of several with one name the last, and no directories.
fn effective_entries(entries: &[StoredEntry]) -> Vec<&StoredEntry> {
    entries
        .iter()
        .enumerate()
        .filter(|(index, entry)| {
            !entry.name.ends_with('/')
                && entries.iter().rposition(|other| other.name == entry.name) == Some(*index)
        })
        .map(|(_, entry)| entry)
        .collect()
}

/// Checks that `after` has every entry of `before` but those in `changed` exactly as it was stored, each name
/// once, and that its XML parts are well-formed.
fn assert_kept(before: &[StoredEntry], after: &[StoredEntry], changed: &[&str], context: &str) {
    for entry in effective_entries(before) {
        if changed.contains(&entry.name.as_str()) {
            continue;
        }
        let kept: Vec<&StoredEntry> = after
            .iter()
            .filter(|kept| kept.name == entry.name)
            .collect();
        assert_eq!(kept, [entry], "{context}: {}", entry.name);
    }
    for entry in after {
        let same_name = after
            .iter()
            .filter(|other| other.name == entry.name)
            .count();
        assert_eq!(same_name, 1, "{context}: {}", entry.name);
        assert!(!entry.name.ends_with('/'), "{context}: {}", entry.name);
        if entry.name.ends_with(".xml") || entry.name.ends_with(".rels") {
            assert_well_formed(entry);
        }
    }
}

#[test]
fn sheets_are_added_to_a_workbook_of_this_writer() {
    let connection = connect();
    let directory = TempDir::new("append_own");
    let path = directory.path("book.xlsx");
    copy(
        &connection,
        "SELECT 1 AS x, 'hello' AS y",
        &path,
        ", HEADER true, SHEET 'A'",
    );
    let first = archive_entries(&path);
    copy(
        &connection,
        TYPED_ROWS,
        &path,
        ", HEADER true, SHEET 'Typed', MODE 'append'",
    );
    let second = archive_entries(&path);
    // The workbook has the cell styles the new sheet needs.
    assert_kept(&first, &second, &PATCHED_PARTS[..3], "typed sheet");
    copy(
        &connection,
        "SELECT 2 AS x, 'world' AS y",
        &path,
        ", HEADER true, SHEET 'B & <C>', MODE 'APPEND'",
    );
    assert_kept(
        &second,
        &archive_entries(&path),
        &PATCHED_PARTS[..3],
        "third sheet",
    );

    assert_eq!(sheet_names(&path), ["A", "Typed", "B & <C>"]);
    assert_eq!(read_sheet(&connection, &path, "A"), ["1.0|hello"]);
    assert_eq!(read_sheet(&connection, &path, "Typed"), TYPED_ROWS_READ);
    assert_eq!(read_sheet(&connection, &path, "B & <C>"), ["2.0|world"]);
    assert_eq!(query(&connection, &format!("FROM '{path}'")), ["1.0|hello"]);
    assert!(sheet_rows(&path, "sheet2.xml").contains(concat!(
        r#"<row r="2"><c r="A2"><v>0</v></c><c r="B2" t="inlineStr"><is><t>text 0</t></is></c>"#,
        r#"<c r="C2" t="b" s="5"><v>1</v></c><c r="D2" s="1"><v>45366.0</v></c>"#,
        r#"<c r="E2" s="3"><v>0.34375</v></c><c r="F2" s="4"><v>45366.52135416667</v></c>"#,
        r#"<c r="G2" s="2"><v>45366.52135416667</v></c></row>"#
    )));
    assert_eq!(directory.file_names(), ["book.xlsx"]);
}

#[test]
fn without_a_workbook_append_and_replace_create_one() {
    let connection = connect();
    let directory = TempDir::new("append_new");
    for mode in ["append", "replace"] {
        for temporary in ["", ", USE_TMP_FILE false", ", USE_TMP_FILE true"] {
            let path = directory.path("new.xlsx");
            copy(
                &connection,
                "SELECT 42 AS x",
                &path,
                &format!(", HEADER true, SHEET 'Only', MODE '{mode}'{temporary}"),
            );
            assert_eq!(read_sheet(&connection, &path, "Only"), ["42.0"]);
            assert_eq!(archive_entries(&path).len(), 9);
            assert_eq!(directory.file_names(), ["new.xlsx"]);
            std::fs::remove_file(&path).unwrap();
        }
    }
}

#[test]
fn a_sheet_takes_the_place_of_the_one_with_its_name() {
    let connection = connect();
    let directory = TempDir::new("replace_own");
    let path = directory.path("book.xlsx");
    copy(
        &connection,
        "SELECT 1 AS x",
        &path,
        ", HEADER true, SHEET 'A'",
    );
    copy(
        &connection,
        "SELECT 2 AS x",
        &path,
        ", HEADER true, SHEET 'Second Sheet', MODE 'append'",
    );
    for temporary in ["", ", USE_TMP_FILE false"] {
        let before = archive_entries(&path);
        copy(
            &connection,
            TYPED_ROWS,
            &path,
            &format!(", HEADER true, SHEET 'a', MODE 'replace'{temporary}"),
        );
        let after = archive_entries(&path);
        assert_kept(&before, &after, &["xl/worksheets/sheet1.xml"], temporary);
        assert_eq!(after.len(), before.len());
        assert_eq!(sheet_names(&path), ["A", "Second Sheet"]);
        assert_eq!(read_sheet(&connection, &path, "A"), TYPED_ROWS_READ);
        assert_eq!(read_sheet(&connection, &path, "Second Sheet"), ["2.0"]);

        copy(
            &connection,
            "SELECT 100 AS x",
            &path,
            &format!(", HEADER true, SHEET 'SECOND SHEET', MODE 'replace'{temporary}"),
        );
        assert_eq!(read_sheet(&connection, &path, "Second Sheet"), ["100.0"]);
        assert_eq!(read_sheet(&connection, &path, "A"), TYPED_ROWS_READ);
        for (sheet, value) in [("A", 1), ("Second Sheet", 2)] {
            copy(
                &connection,
                &format!("SELECT {value} AS x"),
                &path,
                &format!(", HEADER true, SHEET '{sheet}', MODE 'replace'{temporary}"),
            );
        }
        assert_eq!(directory.file_names(), ["book.xlsx"]);
    }
}

#[test]
fn a_workbook_changed_in_place_is_there_all_the_time() {
    let connection = connect();
    let directory = TempDir::new("in_place");
    let path = directory.path("book.xlsx");
    copy(&connection, "SELECT 0 AS round", &path, ", SHEET 'Start'");

    let stop = std::sync::atomic::AtomicBool::new(false);
    let (looks, misses) = std::thread::scope(|scope| {
        let watcher = scope.spawn(|| {
            let (mut looks, mut misses) = (0u64, 0u64);
            while !stop.load(std::sync::atomic::Ordering::Relaxed) {
                looks += 1;
                misses += u64::from(!Path::new(&path).exists());
            }
            (looks, misses)
        });
        for round in 1..=40 {
            let (sheet, mode) = if round % 4 == 0 {
                ("Start".to_owned(), "replace")
            } else {
                (format!("Round {round}"), "append")
            };
            copy(
                &connection,
                &format!("SELECT {round} AS round"),
                &path,
                &format!(", SHEET '{sheet}', MODE '{mode}', USE_TMP_FILE false"),
            );
        }
        stop.store(true, std::sync::atomic::Ordering::Relaxed);
        watcher.join().unwrap()
    });
    assert!(looks > 1000, "{looks} looks");
    assert_eq!(misses, 0, "of {looks} looks");
    assert_eq!(sheet_names(&path).len(), 31);
    assert_eq!(sheet_as_text(&connection, &path, "Start"), ["40"]);
    assert_eq!(directory.file_names(), ["book.xlsx"]);
}

#[test]
fn sheet_names_must_fit_the_mode() {
    let connection = connect();
    let directory = TempDir::new("collisions");
    let path = directory.path("book.xlsx");
    copy(&connection, "SELECT 1 AS x", &path, ", SHEET 'Existing'");
    let original = std::fs::read(&path).unwrap();
    for temporary in ["", ", USE_TMP_FILE false", ", USE_TMP_FILE true"] {
        for name in ["Existing", "existing", "EXISTING"] {
            assert_eq!(
                copy_error(
                    &connection,
                    "SELECT 2",
                    &path,
                    &format!(", SHEET '{name}', MODE 'append'{temporary}")
                ),
                format!(
                    "Invalid Input Error: Sheet '{name}' already exists in '{path}' (use REPLACE to overwrite)"
                )
            );
        }
        for name in ["NotThere", "Existing ", "Éxisting"] {
            assert_eq!(
                copy_error(
                    &connection,
                    "SELECT 2",
                    &path,
                    &format!(", SHEET '{name}', MODE 'replace'{temporary}")
                ),
                format!("Invalid Input Error: Sheet '{name}' not found in '{path}' (REPLACE)")
            );
        }
        assert_eq!(std::fs::read(&path).unwrap(), original, "{temporary}");
        assert_eq!(
            directory.file_names_after_failure(&["book.xlsx"]),
            ["book.xlsx"],
            "{temporary}"
        );
    }
}

#[test]
fn a_failed_append_leaves_the_workbook_as_it_was() {
    let connection = connect();
    let directory = TempDir::new("failed_append");
    let path = directory.path("book.xlsx");
    copy(
        &connection,
        TYPED_ROWS,
        &path,
        ", HEADER true, SHEET 'Data'",
    );
    let original = std::fs::read(&path).unwrap();
    for temporary in ["", ", USE_TMP_FILE false", ", USE_TMP_FILE true"] {
        for (sheet, mode) in [("More", "append"), ("Data", "replace")] {
            let options = format!(", SHEET '{sheet}', MODE '{mode}'{temporary}");
            let context = format!("{mode}{temporary}");
            assert_eq!(
                copy_error(
                    &connection,
                    "SELECT * FROM range(5000)",
                    &path,
                    &format!("{options}, SHEET_ROW_LIMIT 3000")
                ),
                "Invalid Input Error: XLSX: Sheet row limit of '3000' rows exceeded!",
                "{context}"
            );
            assert_eq!(
                directory.file_names_after_failure(&["book.xlsx"]),
                ["book.xlsx"],
                "{context}"
            );
            let error = copy_error(
                &connection,
                "SELECT i::UTINYINT AS small FROM range(100000) t(i)",
                &path,
                &options,
            );
            assert!(error.starts_with("Conversion Error"), "{context}: {error}");
            assert_eq!(
                directory.file_names_after_failure(&["book.xlsx"]),
                ["book.xlsx"],
                "{context}"
            );
            let error = copy_error(
                &connection,
                "SELECT 'x' AS a",
                &path,
                &format!("{options}, HEADER true, SHEET_ROW_LIMIT 0"),
            );
            assert_eq!(
                error, "Invalid Input Error: XLSX: Sheet row limit of '0' rows exceeded!",
                "{context}"
            );
            assert_eq!(
                directory.file_names_after_failure(&["book.xlsx"]),
                ["book.xlsx"],
                "{context}"
            );
            assert_eq!(std::fs::read(&path).unwrap(), original, "{context}");
        }
    }
    assert_eq!(read_sheet(&connection, &path, "Data"), TYPED_ROWS_READ);
}

#[test]
fn workbooks_of_other_producers_keep_everything_else() {
    let connection = connect();
    let directory = TempDir::new("third_party");
    let paths = produced_fixtures();
    assert!(paths.len() >= 25, "{paths:?}");
    let mut without_styles = 0;
    for (index, fixture_path) in paths.iter().enumerate() {
        let temporary = if index % 2 == 0 {
            ""
        } else {
            ", USE_TMP_FILE false"
        };
        let context = format!("{fixture_path}{temporary}");
        let path = directory.copy_fixture(fixture_path, "book.xlsx");
        let original = archive_entries(&path);
        let sheets = sheet_names(&path);
        let texts: Vec<Vec<String>> = sheets
            .iter()
            .map(|sheet| sheet_as_text(&connection, &path, sheet))
            .collect();

        if !original.iter().any(|entry| entry.name == "xl/styles.xml") {
            let error = copy_error(
                &connection,
                TYPED_ROWS,
                &path,
                &format!(", SHEET 'DuckDB & Co', MODE 'append'{temporary}"),
            );
            assert_eq!(
                error, "IO Error: Required entry 'xl/styles.xml' not found in xlsx file",
                "{context}"
            );
            assert_eq!(archive_entries(&path), original, "{context}");
            without_styles += 1;
            continue;
        }
        copy(
            &connection,
            TYPED_ROWS,
            &path,
            &format!(", HEADER true, SHEET 'DuckDB & Co', MODE 'append'{temporary}"),
        );
        let appended = archive_entries(&path);
        assert_kept(&original, &appended, &PATCHED_PARTS, &context);
        let mut all_sheets = sheets.clone();
        all_sheets.push("DuckDB & Co".to_owned());
        assert_eq!(sheet_names(&path), all_sheets, "{context}");
        assert_eq!(
            read_sheet(&connection, &path, "DuckDB & Co"),
            TYPED_ROWS_READ,
            "{context}"
        );
        for (sheet, text) in sheets.iter().zip(&texts) {
            assert_eq!(
                &sheet_as_text(&connection, &path, sheet),
                text,
                "{context}: {sheet}"
            );
        }

        // The new sheet can be replaced, and so can a sheet of the producer.
        copy(
            &connection,
            "SELECT 'again' AS word",
            &path,
            &format!(", SHEET 'duckdb & co', MODE 'replace'{temporary}"),
        );
        let replaced = archive_entries(&path);
        let new_sheet: Vec<&str> = appended
            .iter()
            .filter(|old| !replaced.contains(old))
            .map(|old| old.name.as_str())
            .collect();
        assert_eq!(new_sheet.len(), 1, "{context}: {new_sheet:?}");
        assert_kept(&appended, &replaced, &new_sheet, &context);
        assert_eq!(
            sheet_as_text(&connection, &path, "DuckDB & Co"),
            ["again"],
            "{context}"
        );

        let last = sheets.last().expect("a sheet");
        let last_sql = last.replace('\'', "''");
        copy(
            &connection,
            TYPED_ROWS,
            &path,
            &format!(", HEADER true, SHEET '{last_sql}', MODE 'replace'{temporary}"),
        );
        assert_eq!(sheet_names(&path), all_sheets, "{context}");
        assert_eq!(
            read_sheet(&connection, &path, last),
            TYPED_ROWS_READ,
            "{context}"
        );
        for (sheet, text) in sheets.iter().zip(&texts).rev().skip(1) {
            assert_eq!(
                &sheet_as_text(&connection, &path, sheet),
                text,
                "{context}: {sheet}"
            );
        }
        assert_eq!(directory.file_names(), ["book.xlsx"], "{context}");
    }
    assert!(without_styles <= 3, "{without_styles}");
}

#[test]
fn new_ids_follow_the_largest_and_the_file_name_is_the_first_free_one() {
    let connection = connect();
    let workbook = Workbook::new()
        .sheet("One", &row(1, &[number("A1", 1)]))
        .part(
            "xl/workbook.xml",
            concat!(
                r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" "#,
                r#"xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>"#,
                r#"<sheet name="One" sheetId="7" r:id="rId12"/><sheet name="Chart" r:id="rIdChart"/>"#,
                r#"</sheets><definedNames/></workbook>"#
            ),
        )
        .part(
            "xl/_rels/workbook.xml.rels",
            concat!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
                r#"<Relationship Id="rId12" Type="http://x/relationships/worksheet" Target="/xl/worksheets/sheet1.xml"/>"#,
                r#"<Relationship Id="rIdChart" Type="http://x/relationships/worksheet" Target="worksheets/sheet2.xml"/>"#,
                r#"<Relationship Id="rId40" Type="http://x/relationships/theme" Target="theme/sheet3.xml"/>"#,
                r#"<Relationship Id="rId5" Type="http://x/relationships/worksheet" Target="worksheets/sheet4.xml"/>"#,
                r#"</Relationships>"#
            ),
        )
        .write("append_ids.xlsx");
    let path = workbook.path();
    copy(
        &connection,
        "SELECT 5 AS five",
        path,
        ", SHEET 'New', MODE 'append'",
    );
    assert!(workbook_part(path, "xl/workbook.xml").contains(concat!(
        r#"<sheet name="Chart" r:id="rIdChart"/>"#,
        r#"<sheet name="New" state="visible" sheetId="8" r:id="rId41"/></sheets><definedNames/>"#
    )));
    assert!(
        workbook_part(path, "xl/_rels/workbook.xml.rels").ends_with(concat!(
            r#"<Relationship Id="rId41" Type="http://schemas.openxmlformats.org/officeDocument/2006/"#,
            r#"relationships/worksheet" Target="worksheets/sheet3.xml"/></Relationships>"#
        ))
    );
    assert!(
        workbook_part(path, "[Content_Types].xml").ends_with(concat!(
            r#"<Override PartName="/xl/worksheets/sheet3.xml" ContentType="application/vnd."#,
            "openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/></Types>"
        ))
    );
    assert_eq!(
        query(
            &connection,
            &format!("FROM read_xlsx('{path}', sheet = 'New')")
        ),
        ["5.0"]
    );
    assert_eq!(
        query(
            &connection,
            &format!("FROM read_xlsx('{path}', sheet = 'One')")
        ),
        ["1.0"]
    );
}

#[test]
fn the_last_of_repeated_entries_counts_and_directories_are_dropped() {
    let connection = connect();
    let directory = TempDir::new("repeated");
    let path = directory.path("book.xlsx");
    let base = Workbook::new().sheet("Data", &row(1, &[inline("A1", "first")]));
    let base_path = base.write("repeated_base.xlsx");
    let mut zip = ZipWriter::new(std::fs::File::create(&path).unwrap());
    let mut add = |name: &str, content: &[u8]| {
        zip.begin_file(name).unwrap();
        zip.write_all(content).unwrap();
        zip.end_file().unwrap();
    };
    add("xl/", b"");
    add("xl/worksheets/", b"");
    add("xl/worksheets/sheet1.xml", b"<stale/>");
    add("docProps/custom.xml", b"<stale/>");
    add("[Content_Types].xml", b"<stale/>");
    for entry in archive_entries(base_path.path()) {
        add(&entry.name, &entry.content);
    }
    add("docProps/custom.xml", b"<current/>");
    zip.finish().unwrap();

    copy(
        &connection,
        "SELECT 'second' AS word",
        &path,
        ", SHEET 'More', MODE 'append'",
    );
    let entries = archive_entries(&path);
    let names: Vec<&str> = entries.iter().map(|entry| entry.name.as_str()).collect();
    assert_eq!(
        names,
        [
            "xl/worksheets/sheet1.xml",
            "docProps/custom.xml",
            "xl/worksheets/sheet2.xml",
            "[Content_Types].xml",
            "xl/workbook.xml",
            "xl/_rels/workbook.xml.rels",
            "xl/styles.xml",
        ]
    );
    assert_eq!(entries[1].content, b"<current/>");
    assert!(
        workbook_part(&path, "[Content_Types].xml")
            .starts_with(CONTENT_TYPES.trim_end_matches("</Types>"))
    );
    assert_eq!(sheet_as_text(&connection, &path, "Data"), ["first"]);
    assert_eq!(sheet_as_text(&connection, &path, "More"), ["second"]);
}

#[test]
fn workbooks_that_lack_what_a_new_sheet_needs() {
    let connection = connect();
    let complete = || Workbook::new().sheet("Data", &row(1, &[number("A1", 1)]));
    let rows = "SELECT 1 AS a";
    for part in PATCHED_PARTS {
        let workbook = complete().without(part).write("append_missing.xlsx");
        let path = workbook.path();
        for mode in ["append", "replace"] {
            assert_eq!(
                copy_error(
                    &connection,
                    rows,
                    path,
                    &format!(", SHEET 'Data', MODE '{mode}'")
                ),
                format!("IO Error: Required entry '{part}' not found in xlsx file"),
            );
        }
    }
    for (part, content, message) in [
        (
            "[Content_Types].xml",
            "<Types/>",
            "IO Error: Cannot patch [Content_Types].xml: missing </Types>",
        ),
        (
            "xl/workbook.xml",
            "<workbook><sheets/></workbook>",
            "IO Error: Cannot patch xl/workbook.xml: missing </sheets>",
        ),
        (
            "xl/_rels/workbook.xml.rels",
            "<Relationships/>",
            "IO Error: Cannot patch xl/_rels/workbook.xml.rels: missing </Relationships>",
        ),
        (
            "xl/styles.xml",
            "<styleSheet><cellXfs/></styleSheet>",
            "IO Error: Cannot patch styles.xml: missing </cellXfs>",
        ),
        (
            "xl/styles.xml",
            "<other><cellXfs></cellXfs></other>",
            "IO Error: Cannot patch styles.xml: missing <styleSheet>",
        ),
        (
            "xl/styles.xml",
            "<styleSheet><numFmts><numFmt formatCode=\"0\"/></numFmts></styleSheet>",
            "Invalid Input Error: Invalid numFmt entry in styles.xml",
        ),
        (
            "xl/workbook.xml",
            "<workbook><sheets><sheet name=\"\" r:id=\"rId1\"/></sheets></workbook>",
            "Invalid Input Error: Invalid sheet entry in workbook.xml",
        ),
        (
            "xl/_rels/workbook.xml.rels",
            "<Relationships><Relationship Id=\"rId1\"/></Relationships>",
            "Invalid Input Error: Invalid relationship entry in _rels/.rels",
        ),
        // A new sheet writes its relationship id as `r:id`, so the sheets that are there must do so as well.
        (
            "xl/workbook.xml",
            "<workbook xmlns:rel=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"><sheets>\
             <sheet name=\"Data\" sheetId=\"1\" rel:id=\"rId1\"/></sheets></workbook>",
            "Invalid Input Error: Invalid sheet entry in workbook.xml",
        ),
    ] {
        let directory = TempDir::new("append_broken");
        let path = directory.path("book.xlsx");
        std::fs::write(&path, complete().part(part, content).bytes()).unwrap();
        let original = std::fs::read(&path).unwrap();
        for temporary in ["", ", USE_TMP_FILE false"] {
            assert_eq!(
                copy_error(
                    &connection,
                    "SELECT true AS flag",
                    &path,
                    &format!(", SHEET 'New', MODE 'append'{temporary}")
                ),
                message,
                "{part}: {content}{temporary}"
            );
            assert_eq!(std::fs::read(&path).unwrap(), original, "{part}: {content}");
            assert_eq!(
                directory.file_names_after_failure(&["book.xlsx"]),
                ["book.xlsx"],
                "{part}: {content}"
            );
        }
    }

    let directory = TempDir::new("append_not_a_workbook");
    let path = directory.path("book.xlsx");
    std::fs::write(&path, b"this is not a workbook").unwrap();
    for mode in ["append", "replace"] {
        let error = copy_error(&connection, rows, &path, &format!(", MODE '{mode}'"));
        assert!(error.starts_with("IO Error: "), "{error}");
        assert_eq!(std::fs::read(&path).unwrap(), b"this is not a workbook");
        assert_eq!(
            directory.file_names_after_failure(&["book.xlsx"]),
            ["book.xlsx"]
        );
    }
}

#[test]
fn a_sheet_added_to_a_workbook_that_counts_from_1904_counts_from_1904() {
    let connection = connect();
    let old_sheet = [
        row(1, &[inline("A1", "d"), inline("B1", "ts")]),
        row(2, &[styled("A2", 0, 1), styled("B2", 43538.5, 3)]),
        row(3, &[styled("A3", 59, 1), styled("B3", -1.25, 3)]),
    ]
    .concat();
    let old_rows = [
        "1904-01-01|2023-03-15 12:00:00",
        "1904-02-29|1903-12-30 18:00:00",
    ];
    let rows = "SELECT DATE '1899-12-20' + (i * 9)::INTEGER AS day, \
                TIMESTAMP '1899-12-31 00:00:00.000001' + to_microseconds((i * 7654321987)::BIGINT) AS ts, \
                TIME '00:00:00' + to_microseconds((i * 4321987)::BIGINT) AS tod, \
                (TIMESTAMP '2024-03-15 12:34:56.789' + INTERVAL (i) DAY)::TIMESTAMP_MS AS ts_ms \
                FROM range(400) t(i)";
    let written = query(&connection, rows);
    for (mode, sheet) in [("append", "New"), ("replace", "Other")] {
        for temporary in ["", ", USE_TMP_FILE false"] {
            let directory = TempDir::new("append_1904");
            let path = directory.path("book.xlsx");
            let book = Workbook::new()
                .sheet("Old", &old_sheet)
                .sheet("Other", &row(1, &[number("A1", 1)]))
                .properties("date1904=\"true\"");
            std::fs::write(&path, book.bytes()).unwrap();
            assert_eq!(read_sheet(&connection, &path, "Old"), old_rows);

            copy(
                &connection,
                rows,
                &path,
                &format!(", SHEET '{sheet}', MODE '{mode}', HEADER true{temporary}"),
            );
            assert_eq!(
                read_sheet(&connection, &path, sheet),
                written,
                "{mode}{temporary}"
            );
            assert_eq!(read_sheet(&connection, &path, "Old"), old_rows);
            // 1899-12-20 is 1,472 days before the day the workbook counts from.
            let sheet_file = if mode == "append" {
                "sheet3.xml"
            } else {
                "sheet2.xml"
            };
            let cells = sheet_rows(&path, sheet_file);
            assert!(
                cells.contains(r#"<row r="2"><c r="A2" s="5"><v>-1472.0</v></c>"#),
                "{cells:.600}"
            );
        }
    }

    // A workbook that is made anew counts from 1900, whatever was there before.
    let directory = TempDir::new("create_over_1904");
    let path = directory.path("book.xlsx");
    let book = Workbook::new()
        .sheet("Old", &old_sheet)
        .properties("date1904=\"1\"");
    std::fs::write(&path, book.bytes()).unwrap();
    copy(&connection, "SELECT DATE '2024-03-15' AS day", &path, "");
    assert!(sheet_rows(&path, "sheet1.xml").contains("<v>45366.0</v>"));
    assert_eq!(
        query(
            &connection,
            &format!("FROM read_xlsx('{path}', header = false)")
        ),
        ["2024-03-15"]
    );
}

#[test]
fn parts_in_utf16_are_read_but_not_patched() {
    let connection = connect();
    let complete = || Workbook::new().sheet("Data", &row(1, &[number("A1", 1)]));
    for (part, message) in [
        (
            "[Content_Types].xml",
            "IO Error: Cannot patch [Content_Types].xml: missing </Types>",
        ),
        (
            "xl/workbook.xml",
            "IO Error: Cannot patch xl/workbook.xml: missing </sheets>",
        ),
        (
            "xl/_rels/workbook.xml.rels",
            "IO Error: Cannot patch xl/_rels/workbook.xml.rels: missing </Relationships>",
        ),
    ] {
        let directory = TempDir::new("append_utf16");
        let path = directory.path("book.xlsx");
        let bytes = complete().bytes_with(|name, content| {
            if name == part {
                let text = std::str::from_utf8(content).expect("a part in UTF-8");
                utf16(
                    &text.replace("encoding=\"UTF-8\"", "encoding=\"UTF-16\""),
                    false,
                    true,
                )
            } else {
                content.to_vec()
            }
        });
        std::fs::write(&path, &bytes).unwrap();
        assert_eq!(
            query(
                &connection,
                &format!("FROM read_xlsx('{path}', header = false)")
            ),
            ["1.0"],
            "{part}"
        );
        assert_eq!(
            copy_error(
                &connection,
                "SELECT 2 AS a",
                &path,
                ", SHEET 'New', MODE 'append'"
            ),
            message
        );
        assert_eq!(std::fs::read(&path).unwrap(), bytes, "{part}");
    }
}

#[test]
fn a_sheet_outside_the_worksheets_cannot_be_replaced() {
    let connection = connect();
    let relations = |target: &str| {
        format!(
            concat!(
                r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#,
                r#"<Relationship Id="rId1" Type="http://x/relationships/worksheet" Target="{}"/>"#,
                r#"</Relationships>"#
            ),
            target
        )
    };
    let workbook = Workbook::new()
        .sheet("Data", &row(1, &[number("A1", 1)]))
        .part("xl/_rels/workbook.xml.rels", relations("sheets/first.xml"))
        .write("replace_elsewhere.xlsx");
    assert_eq!(
        copy_error(
            &connection,
            "SELECT 1",
            workbook.path(),
            ", SHEET 'Data', MODE 'replace'"
        ),
        "IO Error: REPLACE not supported: sheet 'Data' is stored at 'xl/sheets/first.xml', expected \
         xl/worksheets/<name>"
    );
    let workbook = Workbook::new()
        .sheet("Data", &row(1, &[number("A1", 1)]))
        .part(
            "xl/_rels/workbook.xml.rels",
            relations("worksheets/sheet1.xml").replace("rId1", "rId2"),
        )
        .write("replace_unresolved.xlsx");
    assert_eq!(
        copy_error(
            &connection,
            "SELECT 1",
            workbook.path(),
            ", SHEET 'Data', MODE 'replace'"
        ),
        "IO Error: Could not resolve relationship 'rId1' for sheet 'Data'"
    );

    // A sheet stored under an absolute path is found where it is.
    let workbook = Workbook::new()
        .sheet("Data", &row(1, &[number("A1", 1)]))
        .part(
            "xl/_rels/workbook.xml.rels",
            relations("/xl/worksheets/sheet1.xml"),
        )
        .write("replace_absolute.xlsx");
    let path = workbook.path();
    copy(
        &connection,
        "SELECT 9 AS nine",
        path,
        ", SHEET 'DATA', MODE 'replace'",
    );
    assert_eq!(sheet_as_text(&connection, path, "Data"), ["9"]);
}
