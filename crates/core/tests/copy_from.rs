//! `COPY tbl FROM 'file.xlsx' (FORMAT xlsx)`.

mod support;

use support::{
    Workbook, connect, execute, execute_error, fixture, inline, number, numbered_rows, query, row,
    styled, typed,
};

fn people() -> Workbook {
    Workbook::new()
        .sheet(
            "People",
            &[
                row(
                    1,
                    &[
                        inline("A1", "id"),
                        inline("B1", "name"),
                        inline("C1", "born"),
                        inline("D1", "active"),
                    ],
                ),
                row(
                    2,
                    &[
                        number("A2", 1),
                        inline("B2", "Ada"),
                        styled("C2", 5823, 1),
                        typed("D2", "b", "1"),
                    ],
                ),
                row(
                    3,
                    &[
                        number("A3", 2),
                        inline("B3", "Grace"),
                        styled("C3", 2535, 1),
                        typed("D3", "b", "0"),
                    ],
                ),
                row(
                    4,
                    &[
                        number("A4", 3),
                        inline("B4", "Edsger"),
                        styled("C4", 11089, 1),
                    ],
                ),
            ]
            .concat(),
        )
        .sheet("Numbers", &numbered_rows(1, 5000))
}

#[test]
fn rows_are_cast_to_the_types_of_the_table() {
    let connection = connect();
    let book = people().write("copy_people.xlsx");
    let path = book.path();
    execute(
        &connection,
        "CREATE TABLE people(id INTEGER, name VARCHAR, born DATE, active BOOLEAN)",
    );
    execute(
        &connection,
        &format!("COPY people FROM '{path}' (FORMAT xlsx)"),
    );
    assert_eq!(
        query(&connection, "FROM people"),
        [
            "1|Ada|1915-12-10|true",
            "2|Grace|1906-12-09|false",
            "3|Edsger|1930-05-11|NULL"
        ]
    );

    execute(
        &connection,
        "CREATE TABLE loose(id DECIMAL(5, 2), name VARCHAR, born VARCHAR, active VARCHAR)",
    );
    execute(
        &connection,
        &format!("COPY loose FROM '{path}' (FORMAT xlsx)"),
    );
    assert_eq!(
        query(&connection, "FROM loose"),
        [
            "1.00|Ada|5823|1",
            "2.00|Grace|2535|0",
            "3.00|Edsger|11089|NULL"
        ]
    );

    execute(
        &connection,
        "CREATE TABLE stamps(id UTINYINT, name VARCHAR, born TIMESTAMP, active INTEGER)",
    );
    execute(
        &connection,
        &format!("COPY stamps FROM '{path}' (FORMAT xlsx)"),
    );
    assert_eq!(
        query(&connection, "FROM stamps"),
        [
            "1|Ada|1915-12-10 00:00:00|1",
            "2|Grace|1906-12-09 00:00:00|0",
            "3|Edsger|1930-05-11 00:00:00|NULL"
        ]
    );
}

#[test]
fn the_format_follows_from_the_file_name() {
    let connection = connect();
    execute(
        &connection,
        "CREATE TABLE sheet(abc DOUBLE, hello VARCHAR, world VARCHAR)",
    );
    execute(
        &connection,
        &format!("COPY sheet FROM '{}'", fixture("google_sheets.xlsx")),
    );
    assert_eq!(
        query(&connection, "FROM sheet"),
        ["123.0|500,60|ABC", "456.0|300,40|some text"]
    );
}

#[test]
fn options_choose_sheet_range_and_header() {
    let connection = connect();
    let book = people().write("copy_options.xlsx");
    let path = book.path();
    execute(
        &connection,
        "CREATE TABLE numbers(id BIGINT, label VARCHAR)",
    );
    let copy = |options: &str| {
        execute(&connection, "DELETE FROM numbers");
        execute(
            &connection,
            &format!("COPY numbers FROM '{path}' (FORMAT xlsx, sheet 'Numbers'{options})"),
        );
        query(
            &connection,
            "SELECT count(*), min(id), max(id) FROM numbers",
        )
    };
    assert_eq!(copy(""), ["5000|1|5000"]);
    assert_eq!(copy(", header"), ["4999|2|5000"]);
    assert_eq!(copy(", HEADER true"), ["4999|2|5000"]);
    assert_eq!(copy(", header 0"), ["5000|1|5000"]);
    assert_eq!(copy(", range 'A10:B19'"), ["10|10|19"]);
    assert_eq!(copy(", range 'A4990:B5010'"), ["21|4990|5000"]);
    assert_eq!(
        copy(", range 'A4990:B5010', stop_at_empty"),
        ["11|4990|5000"]
    );
    assert_eq!(
        copy(", range 'A4990:B5010', stop_at_empty false"),
        ["21|4990|5000"]
    );
    assert_eq!(
        copy(", all_varchar, ignore_errors 'true', empty_as_varchar 1, unknown_option 7"),
        ["5000|1|5000"]
    );
}

#[test]
fn cells_that_do_not_fit_the_table_fail_or_become_null() {
    let connection = connect();
    let book = people().write("copy_errors.xlsx");
    let path = book.path();
    execute(
        &connection,
        "CREATE TABLE strict(id INTEGER, name INTEGER, born DATE, active BOOLEAN)",
    );
    assert_eq!(
        execute_error(
            &connection,
            &format!("COPY strict FROM '{path}' (FORMAT xlsx)")
        ),
        "Invalid Input Error: read_xlsx: Failed to parse cell 'B2': Could not convert string 'Ada' to INT32"
    );
    assert_eq!(query(&connection, "SELECT count(*) FROM strict"), ["0"]);
    execute(
        &connection,
        &format!("COPY strict FROM '{path}' (FORMAT xlsx, ignore_errors)"),
    );
    assert_eq!(
        query(&connection, "FROM strict"),
        [
            "1|NULL|1915-12-10|true",
            "2|NULL|1906-12-09|false",
            "3|NULL|1930-05-11|NULL"
        ]
    );
}

#[test]
fn the_sheet_must_have_as_many_columns_as_the_table() {
    let connection = connect();
    let path = fixture("google_sheets.xlsx");
    execute(
        &connection,
        "CREATE TABLE narrow(a VARCHAR, \"b c\" INTEGER)",
    );
    assert_eq!(
        execute_error(
            &connection,
            &format!("COPY narrow FROM '{path}' (FORMAT xlsx)")
        ),
        format!(
            "Conversion Error: Failed to read file(s) \"{path}\" - column count mismatch: expected 2 columns but found 3\n\
             Table schema: a VARCHAR, b c INTEGER\n\
             XLSX schema: ABC DOUBLE, HELLO VARCHAR, WORLD VARCHAR\n\n\
             Possible solutions:\n\
             * Manually specify which columns to insert using \"INSERT INTO tbl SELECT ... FROM read_xlsx(...)\"\n\
             * Provide an explicit range option with the same width as the table schema using e.g.\
             \"COPY tbl FROM ... (FORMAT 'xlsx', range 'A1:Z10')\""
        )
    );
    execute(
        &connection,
        &format!("COPY narrow FROM '{path}' (FORMAT xlsx, range 'A1:B3', header, ignore_errors)"),
    );
    assert_eq!(
        query(&connection, "FROM narrow"),
        ["123.0|NULL", "456.0|NULL"]
    );
}

#[test]
fn options_are_checked() {
    let connection = connect();
    let path = fixture("google_sheets.xlsx");
    execute(
        &connection,
        "CREATE TABLE target(a VARCHAR, b VARCHAR, c VARCHAR)",
    );
    for (option, error) in [
        (
            "header 1.5",
            "Binder Error: '\"header\"' options expects a boolean value (e.g. TRUE or 1)",
        ),
        (
            "ALL_VARCHAR 2.5e0",
            "Binder Error: '\"all_varchar\"' options expects a boolean value (e.g. TRUE or 1)",
        ),
        (
            "header (true, false)",
            "Binder Error: '\"header\"' option must be standalone or a BOOLEAN value",
        ),
        (
            "ignore_errors 'abc'",
            "Invalid Input Error: Failed to cast value: Could not convert string 'abc' to BOOL",
        ),
        (
            "sheet",
            "Binder Error: '\"sheet\"' option must be a single VARCHAR value",
        ),
        (
            "sheet 42",
            "Binder Error: '\"sheet\"' option must be a single VARCHAR value",
        ),
        (
            "sheet ('a', 'b')",
            "Binder Error: '\"sheet\"' option must be a single VARCHAR value",
        ),
        (
            "range 5",
            "Binder Error: '\"range\"' option must be a single VARCHAR value",
        ),
        ("range 'bad'", "Binder Error: Invalid range 'bad' specified"),
    ] {
        let sql = format!("COPY target FROM '{path}' (FORMAT xlsx, {option})");
        assert_eq!(execute_error(&connection, &sql), error, "{option}");
    }
    let error = execute_error(
        &connection,
        &format!("COPY target FROM '{path}' (FORMAT xlsx, sheet 'Nope')"),
    );
    assert_eq!(
        error,
        format!(
            "Binder Error: Sheet \"Nope\" not found in xlsx file \"{path}\"\nDid you mean: \"Sheet1\""
        )
    );
    assert_eq!(query(&connection, "SELECT count(*) FROM target"), ["0"]);
}

#[test]
fn the_path_is_not_a_glob() {
    let connection = connect();
    execute(&connection, "CREATE TABLE target(a VARCHAR)");
    let pattern = fixture("google_*.xlsx");
    let error = execute_error(
        &connection,
        &format!("COPY target FROM '{pattern}' (FORMAT xlsx)"),
    );
    assert_eq!(
        error,
        format!("IO Error: Cannot open file \"{pattern}\": No such file or directory")
    );
}

#[test]
fn big_sheets_are_copied_in_order() {
    let connection = connect();
    let book = Workbook::new()
        .sheet("S", &numbered_rows(1, 7000))
        .write("copy_big.xlsx");
    execute(&connection, "CREATE TABLE big(id INTEGER, label VARCHAR)");
    execute(
        &connection,
        &format!("COPY big FROM '{}' (FORMAT xlsx)", book.path()),
    );
    assert_eq!(
        query(
            &connection,
            "SELECT count(*), sum(id), count(DISTINCT label) FROM big"
        ),
        ["7000|24503500|7000"]
    );
    let unordered = query(
        &connection,
        "SELECT count(*) FROM big WHERE label <> 'label ' || id",
    );
    assert_eq!(unordered, ["0"]);
    assert_eq!(
        query(
            &connection,
            "SELECT count(*) FROM (SELECT id, rowid FROM big) WHERE id <> rowid + 1"
        ),
        ["0"]
    );
}
