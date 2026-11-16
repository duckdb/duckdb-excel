//! The options of `read_xlsx`: how they are passed, checked and combined.

mod support;

use support::{
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
    support::execute(
        &connection,
        "PREPARE read AS SELECT CAST(COLUMNS(*) AS VARCHAR) FROM read_xlsx($1, header = $2, range = $3)",
    );
    let execute = |arguments: &str| {
        support::try_query_text(&connection, &format!("EXECUTE read('{path}', {arguments})"))
            .unwrap()
    };
    assert_eq!(execute("true, 'A1:B2'"), ["123.0|500,60"]);
    assert_eq!(execute("false, 'A1:B2'"), ["ABC|HELLO", "123.0|500,60"]);
    assert_eq!(execute("NULL, 'A1:A3'"), ["123.0", "456.0"]);
}
