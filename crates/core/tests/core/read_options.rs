//! The options of `read_xlsx`: how they are passed, checked and combined.

use crate::support::{
    Workbook, connect, describe, fixture, inline, number, numbered_rows, query, query_error, row,
};

fn two_sheets() -> Workbook {
    Workbook::new()
        .sheet(
            "First",
            &[
                row(1, &[inline("A1", "id"), inline("B1", "name")]),
                row(2, &[number("A2", 1), inline("B2", "one")]),
                row(3, &[number("A3", 2), inline("B3", "two")]),
            ]
            .concat(),
        )
        .sheet(
            "A & <B>",
            &[row(1, &[number("A1", 10)]), row(2, &[number("A2", 20)])].concat(),
        )
}

#[test]
fn options_can_be_given_in_any_order_and_spelling() {
    let connection = connect();
    let book = two_sheets().write("options.xlsx");
    let path = book.path();
    let expected = ["1.0|one", "2.0|two"];
    for arguments in [
        String::new(),
        ", header = true".to_owned(),
        ", HEADER := true, sheet = 'First'".to_owned(),
        ", sheet := 'First', all_varchar = false, ignore_errors = false".to_owned(),
        ", normalize_names = false, empty_as_varchar = false, stop_at_empty = true".to_owned(),
    ] {
        let rows = query(&connection, &format!("FROM read_xlsx('{path}'{arguments})"));
        assert_eq!(rows, expected, "{arguments}");
    }
    assert_eq!(
        query(
            &connection,
            &format!("FROM read_xlsx(file_path = '{path}')")
        ),
        expected
    );
}

#[test]
fn options_are_cast_like_explicit_casts() {
    let connection = connect();
    let book = two_sheets().write("option_casts.xlsx");
    let path = book.path();
    for arguments in [
        "header = 0",
        "header = 'false'",
        "header = 'f'",
        "header = 0.0",
    ] {
        let rows = query(
            &connection,
            &format!("FROM read_xlsx('{path}', {arguments})"),
        );
        assert_eq!(rows, ["id|name", "1|one", "2|two"], "{arguments}");
    }
    for arguments in [
        "all_varchar = 1",
        "all_varchar = 'yes'",
        "all_varchar = 2.5",
    ] {
        let rows = query(
            &connection,
            &format!("FROM read_xlsx('{path}', {arguments})"),
        );
        assert_eq!(rows, ["1|one", "2|two"], "{arguments}");
    }
    let error = query_error(
        &connection,
        &format!("FROM read_xlsx('{path}', sheet = 42)"),
    );
    assert_eq!(
        error,
        format!(
            "Binder Error: Sheet \"42\" not found in xlsx file \"{path}\"\nDid you mean: \"A & <B>\", \"First\""
        )
    );
}

#[test]
fn options_that_do_not_cast_are_rejected() {
    let connection = connect();
    let path = fixture("basic.xlsx");
    for (arguments, message) in [
        (
            "header = 'abc'",
            "Could not cast value 'abc' to named parameter \"header\" of type BOOLEAN: Could not convert string 'abc' to BOOL",
        ),
        (
            "all_varchar = 'it''s'",
            "Could not cast value 'it''s' to named parameter \"all_varchar\" of type BOOLEAN: Could not convert string 'it's' to BOOL",
        ),
        (
            "stop_at_empty = DATE '2020-01-01'",
            "Could not cast value '2020-01-01'::DATE to named parameter \"stop_at_empty\" of type BOOLEAN: Unimplemented type for cast (DATE -> BOOLEAN)",
        ),
        (
            "ignore_errors = [true]",
            "Could not cast value [true] to named parameter \"ignore_errors\" of type BOOLEAN: Unimplemented type for cast (BOOLEAN[] -> BOOLEAN)",
        ),
    ] {
        let error = query_error(
            &connection,
            &format!("FROM read_xlsx('{path}', {arguments})"),
        );
        assert_eq!(
            error,
            format!("Invalid Input Error: {message}"),
            "{arguments}"
        );
    }
}

#[test]
fn unknown_and_misplaced_arguments_are_rejected() {
    let connection = connect();
    let path = fixture("basic.xlsx");
    let error = query_error(
        &connection,
        &format!("FROM read_xlsx('{path}', hader = true)"),
    );
    assert!(
        error.starts_with("Binder Error: Invalid named parameter \"hader\" for function read_xlsx"),
        "{error}"
    );
    let error = query_error(
        &connection,
        &format!("FROM read_xlsx('{path}', header = true, header = false)"),
    );
    assert_eq!(
        error,
        "Binder Error: Duplicate named argument \"header\" for function read_xlsx"
    );
    for sql in [
        format!("FROM read_xlsx('{path}', true)"),
        "FROM read_xlsx(42)".to_owned(),
        "FROM read_xlsx()".to_owned(),
        format!("FROM read_xlsx(['{path}'])"),
    ] {
        let error = query_error(&connection, &sql);
        assert!(
            error.starts_with("Binder Error: No function matches"),
            "{sql}: {error}"
        );
    }
}

#[test]
fn a_null_option_is_an_option_not_given() {
    let connection = connect();
    let book = two_sheets().write("null_options.xlsx");
    let path = book.path();
    let expected = ["1.0|one", "2.0|two"];
    for arguments in [
        "header = NULL",
        "sheet = NULL, range = NULL",
        "all_varchar = NULL::BOOLEAN, stop_at_empty = NULL, ignore_errors = NULL",
        "empty_as_varchar = NULL, normalize_names = NULL::VARCHAR",
    ] {
        let rows = query(
            &connection,
            &format!("FROM read_xlsx('{path}', {arguments})"),
        );
        assert_eq!(rows, expected, "{arguments}");
    }
    for sql in ["FROM read_xlsx(NULL)", "FROM read_xlsx(NULL::VARCHAR)"] {
        assert_eq!(
            query_error(&connection, sql),
            "Binder Error: read_xlsx: the file path cannot be NULL"
        );
    }
}

#[test]
fn the_file_is_found_by_glob() {
    let connection = connect();
    let path = fixture("google_sheets.xlsx");
    let pattern = path.replace("google_sheets", "google_*");
    assert_eq!(
        query(&connection, &format!("FROM read_xlsx('{pattern}')")),
        query(&connection, &format!("FROM read_xlsx('{path}')"))
    );
    for missing in [
        "nope.xlsx".to_owned(),
        path.replace("google_sheets", "nothing*"),
        String::new(),
    ] {
        let error = query_error(&connection, &format!("FROM read_xlsx('{missing}')"));
        assert_eq!(
            error,
            format!("IO Error: Cannot open file \"{missing}\": No such file or directory")
        );
    }
}

#[test]
fn sheets_are_chosen_by_name() {
    let connection = connect();
    let book = two_sheets().write("sheets.xlsx");
    let path = book.path();
    assert_eq!(
        query(&connection, &format!("FROM read_xlsx('{path}')")),
        ["1.0|one", "2.0|two"]
    );
    assert_eq!(
        query(
            &connection,
            &format!("FROM read_xlsx('{path}', sheet = '')")
        ),
        ["1.0|one", "2.0|two"]
    );
    assert_eq!(
        query(
            &connection,
            &format!("FROM read_xlsx('{path}', sheet = 'A & <B>')")
        ),
        ["10.0", "20.0"]
    );
    let error = query_error(
        &connection,
        &format!("FROM read_xlsx('{path}', sheet = 'first')"),
    );
    assert_eq!(
        error,
        format!(
            "Binder Error: Sheet \"first\" not found in xlsx file \"{path}\"\nDid you mean: \"First\""
        )
    );
    // The name in the message is the escaped one that is looked for.
    let error = query_error(
        &connection,
        &format!("FROM read_xlsx('{path}', sheet = 'A & B')"),
    );
    assert_eq!(
        error,
        format!(
            "Binder Error: Sheet \"A &amp; B\" not found in xlsx file \"{path}\"\nDid you mean: \"A & <B>\""
        )
    );
}

const WORKSHEET_RELATION: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";

fn workbook_part(sheets: &str) -> String {
    format!(
        "<workbook xmlns=\"http://schemas.openxmlformats.org/spreadsheetml/2006/main\" \
         xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\">\
         <sheets>{sheets}</sheets></workbook>"
    )
}

fn relations_part(relations: &str) -> String {
    format!(
        "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">\
         {relations}</Relationships>"
    )
}

#[test]
fn the_relationship_id_of_a_sheet_is_the_id_in_the_relationships_namespace() {
    let connection = connect();
    let transitional = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    let strict = "http://purl.oclc.org/ooxml/officeDocument/relationships";
    let book = |workbook: &str| {
        Workbook::new()
            .sheet("One", &row(1, &[number("A1", 1)]))
            .sheet("Two", &row(1, &[number("A1", 2)]))
            .part("xl/workbook.xml", workbook)
            .write("relationship_prefix.xlsx")
    };
    let read = |workbook: &str, sheet: &str| {
        let file = book(workbook);
        query(
            &connection,
            &format!(
                "FROM read_xlsx('{}', sheet = '{sheet}', header = false)",
                file.path()
            ),
        )
    };
    for namespace in [transitional, strict] {
        let on_the_root = format!(
            "<x:workbook xmlns:x=\"main\" xmlns:rel=\"{namespace}\"><x:sheets>\
             <x:sheet name=\"One\" sheetId=\"1\" rel:id=\"rId1\"/><x:sheet name=\"Two\" sheetId=\"2\" rel:id=\"rId2\"/>\
             </x:sheets></x:workbook>"
        );
        assert_eq!(read(&on_the_root, "One"), ["1.0"], "{namespace}");
        assert_eq!(read(&on_the_root, "Two"), ["2.0"], "{namespace}");
        let on_every_sheet = format!(
            "<workbook><sheets><sheet xmlns:a=\"{namespace}\" name=\"One\" a:id=\"rId1\"/>\
             <sheet name=\"Two\" xmlns:b=\"{namespace}\" b:id=\"rId2\" r:id=\"rId1\"/></sheets></workbook>"
        );
        assert_eq!(read(&on_every_sheet, "One"), ["1.0"], "{namespace}");
        assert_eq!(read(&on_every_sheet, "Two"), ["2.0"], "{namespace}");
    }
    // Without a prefix for the namespace, the attribute written `r:id` is the one.
    let literal = "<workbook xmlns:r=\"other\" xmlns:o=\"other\"><sheets><sheet name=\"One\" o:id=\"rId2\" r:id=\"rId1\"/>\
                   <sheet name=\"Two\" r:id=\"rId2\" id=\"rId1\"/></sheets></workbook>";
    assert_eq!(read(literal, "One"), ["1.0"]);
    assert_eq!(read(literal, "Two"), ["2.0"]);
    for workbook in [
        "<workbook xmlns:rel=\"other\"><sheets><sheet name=\"One\" rel:id=\"rId1\"/></sheets></workbook>".to_owned(),
        "<workbook><sheets><sheet name=\"One\" rel:id=\"rId1\"/></sheets></workbook>".to_owned(),
        format!(
            "<workbook><other xmlns:rel=\"{transitional}\"/><sheets><sheet name=\"One\" rel:id=\"rId1\"/></sheets>\
             </workbook>"
        ),
    ] {
        let file = book(&workbook);
        assert_eq!(
            query_error(&connection, &format!("FROM read_xlsx('{}')", file.path())),
            "Invalid Input Error: Invalid sheet entry in workbook.xml",
            "{workbook}"
        );
    }
}

#[test]
fn the_last_of_repeated_sheets_and_relationships_counts() {
    let connection = connect();
    let relation = |id: &str, kind: &str, target: &str| {
        format!(r#"<Relationship Id="{id}" Type="{kind}" Target="{target}"/>"#)
    };
    let theme = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
    let chart = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
    let relations = [
        relation("rId1", WORKSHEET_RELATION, "worksheets/sheet1.xml"),
        relation("rId2", WORKSHEET_RELATION, "worksheets/sheet2.xml"),
        relation("rId3", WORKSHEET_RELATION, "worksheets/sheet3.xml"),
        relation("rId2", WORKSHEET_RELATION, "worksheets/sheet4.xml"),
        relation("rId4", chart, "worksheets/sheet1.xml"),
        relation("rId3", theme, "theme/theme1.xml"),
        relation("rId5", theme, "theme/theme1.xml"),
        relation("rId5", WORKSHEET_RELATION, "/xl/worksheets/sheet2.xml"),
    ]
    .concat();
    let sheets = concat!(
        r#"<sheet name="Dup" sheetId="1" r:id="rId1"/><sheet name="Other" sheetId="2" r:id="rId2"/>"#,
        r#"<sheet name="Dup" sheetId="3" r:id="rId3"/><sheet name="NoRelation" sheetId="4" r:id="rId9"/>"#,
        r#"<sheet name="Chart" sheetId="5" r:id="rId4"/><sheet name="Late" sheetId="6" r:id="rId5"/>"#
    );
    let mut book = Workbook::new();
    for value in 1..=4 {
        book = book.sheet(&format!("S{value}"), &row(1, &[number("A1", value)]));
    }
    let book = book
        .part("xl/workbook.xml", workbook_part(sheets))
        .part("xl/_rels/workbook.xml.rels", relations_part(&relations))
        .write("repeated_sheets.xlsx");
    let path = book.path();
    for (sheet, value) in [
        ("Dup", "3.0"),
        ("Other", "4.0"),
        ("Late", "2.0"),
        ("", "3.0"),
    ] {
        assert_eq!(
            query(
                &connection,
                &format!("FROM read_xlsx('{path}', sheet = '{sheet}')")
            ),
            [value],
            "{sheet}"
        );
    }
    assert_eq!(
        query_error(
            &connection,
            &format!("FROM read_xlsx('{path}', sheet = 'Chart')")
        ),
        format!(
            "Binder Error: Sheet \"Chart\" not found in xlsx file \"{path}\"\nDid you mean: \"Late\""
        )
    );
    assert!(
        query_error(
            &connection,
            &format!("FROM read_xlsx('{path}', sheet = 'NoRelation')")
        )
        .starts_with("Binder Error: Sheet \"NoRelation\" not found")
    );
}

#[test]
fn a_sheet_is_found_among_very_many() {
    let connection = connect();
    let count = 30_000;
    let sheets: String = (1..=count)
        .map(|id| format!(r#"<sheet name="Sheet {id}" sheetId="{id}" r:id="rId{id}"/>"#))
        .collect();
    let relations: String = (1..=count)
        .map(|id| {
            let part = if id == 29_999 { 2 } else { 1 };
            format!(
                r#"<Relationship Id="rId{id}" Type="{WORKSHEET_RELATION}" Target="worksheets/sheet{part}.xml"/>"#
            )
        })
        .collect();
    let book = Workbook::new()
        .sheet("S1", &row(1, &[number("A1", 1)]))
        .sheet("S2", &row(1, &[number("A1", 2)]))
        .part("xl/workbook.xml", workbook_part(&sheets))
        .part("xl/_rels/workbook.xml.rels", relations_part(&relations))
        .write("many_sheets.xlsx");
    let path = book.path();
    for (sheet, value) in [
        ("Sheet 1", "1.0"),
        ("Sheet 29999", "2.0"),
        ("Sheet 30000", "1.0"),
    ] {
        assert_eq!(
            query(
                &connection,
                &format!("FROM read_xlsx('{path}', sheet = '{sheet}')")
            ),
            [value],
            "{sheet}"
        );
    }
    assert_eq!(
        query_error(
            &connection,
            &format!("FROM read_xlsx('{path}', sheet = 'Sheet 30001')")
        ),
        format!(
            "Binder Error: Sheet \"Sheet 30001\" not found in xlsx file \"{path}\"\nDid you mean: \
             \"Sheet 3000\", \"Sheet 3001\", \"Sheet 10001\", \"Sheet 13000\", \"Sheet 13001\""
        )
    );
}

#[test]
fn sheet_names_are_unescaped_once_more_than_xml_asks_for() {
    let connection = connect();
    let book = Workbook::new()
        .sheet("&lt;x&gt;", &row(1, &[number("A1", 1)]))
        .sheet("&amp;amp;", &row(1, &[number("A1", 2)]))
        .write("escaped_sheets.xlsx");
    let path = book.path();
    assert_eq!(
        query(
            &connection,
            &format!("FROM read_xlsx('{path}', sheet = '<x>')")
        ),
        ["1.0"]
    );
    assert_eq!(
        query(
            &connection,
            &format!("FROM read_xlsx('{path}', sheet = '&amp;')")
        ),
        ["2.0"]
    );
    assert!(
        query_error(
            &connection,
            &format!("FROM read_xlsx('{path}', sheet = '&lt;x&gt;')")
        )
        .starts_with("Binder Error: Sheet \"&amp;lt;x&amp;gt;\" not found")
    );
}

#[test]
fn ranges() {
    let connection = connect();
    let book = Workbook::new()
        .sheet("S", &numbered_rows(1, 10))
        .write("ranges.xlsx");
    let path = book.path();
    let read = |arguments: &str| {
        query(
            &connection,
            &format!("FROM read_xlsx('{path}', {arguments})"),
        )
    };

    assert_eq!(read("range = 'A2:B3'"), ["2.0|label 2", "3.0|label 3"]);
    assert_eq!(read("range = 'B2:B3'"), ["label 3"]);
    assert_eq!(
        read("range = 'B2:B3', header = false"),
        ["label 2", "label 3"]
    );
    assert_eq!(
        read("range = 'A9:C12'"),
        [
            "9.0|label 9|NULL",
            "10.0|label 10|NULL",
            "NULL|NULL|NULL",
            "NULL|NULL|NULL"
        ]
    );
    assert_eq!(
        read("range = 'A9:C12', stop_at_empty = true"),
        ["9.0|label 9|NULL", "10.0|label 10|NULL"]
    );
    assert_eq!(read("range = 'A20:B21'"), ["NULL|NULL", "NULL|NULL"]);
    assert_eq!(read("range = 'A2:A4 and more'"), ["2.0", "3.0", "4.0"]);
    assert_eq!(
        describe(
            &connection,
            &format!("FROM read_xlsx('{path}', range = 'A20:B21')")
        ),
        ["A20 DOUBLE", "B20 DOUBLE"]
    );
    assert_eq!(
        describe(
            &connection,
            &format!("FROM read_xlsx('{path}', range = 'D2:E3', empty_as_varchar = true)")
        ),
        ["D VARCHAR", "E VARCHAR"]
    );
    assert_eq!(
        query(
            &connection,
            &format!("SELECT count(*), count(A) FROM read_xlsx('{path}', range = 'A:B')")
        ),
        ["1048576|10"]
    );
    assert_eq!(
        query(
            &connection,
            &format!(
                "SELECT count(*) FROM read_xlsx('{path}', range = 'A:B', stop_at_empty = true)"
            )
        ),
        ["10"]
    );
    assert_eq!(
        query(
            &connection,
            &format!("SELECT count(*) FROM (DESCRIBE FROM read_xlsx('{path}', range = '3:4'))")
        ),
        ["16384"]
    );
}

#[test]
fn invalid_ranges() {
    let connection = connect();
    let path = fixture("google_sheets.xlsx");
    for range in [
        "A1:ZZZZZZZZ5",
        "Z1:A5",
        "A5:A1",
        "A1",
        "A1:",
        ":B2",
        "a1:b2",
        "A1-B2",
        "",
        "A1:XFF1",
        "5",
    ] {
        let error = query_error(
            &connection,
            &format!("FROM read_xlsx('{path}', range = '{range}')"),
        );
        assert_eq!(
            error,
            format!("Binder Error: Invalid range '{range}' specified")
        );
    }
}

#[test]
fn a_missing_file_is_reported_before_a_bad_range() {
    let connection = connect();
    let error = query_error(&connection, "FROM read_xlsx('nope.xlsx', range = 'bad')");
    assert!(error.starts_with("IO Error: Cannot open file"), "{error}");
    let error = query_error(&connection, "FROM read_xlsx('nope.xlsx', header = 'bad')");
    assert!(
        error.starts_with("Invalid Input Error: Could not cast value"),
        "{error}"
    );
}

#[test]
fn parameters_of_prepared_statements_are_options() {
    let connection = connect();
    let path = fixture("google_sheets.xlsx");
    crate::support::execute(
        &connection,
        "PREPARE read AS SELECT CAST(COLUMNS(*) AS VARCHAR) FROM read_xlsx($1, header = $2, range = $3)",
    );
    let execute = |arguments: &str| {
        crate::support::try_query_text(&connection, &format!("EXECUTE read('{path}', {arguments})"))
            .unwrap()
    };
    assert_eq!(execute("true, 'A1:B2'"), ["123.0|500,60"]);
    assert_eq!(execute("false, 'A1:B2'"), ["ABC|HELLO", "123.0|500,60"]);
    assert_eq!(execute("NULL, 'A1:A3'"), ["123.0", "456.0"]);
}
