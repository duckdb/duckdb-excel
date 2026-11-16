//! What `read_xlsx` makes of the first rows of a sheet: the range, the header, the column types and names.

use crate::support::{
    Workbook, connect, describe, inline, number, query, query_error, row, shared, styled, typed,
};

fn write(name: &str, sheet_data: &str) -> crate::support::TempFile {
    Workbook::new().sheet("S", sheet_data).write(name)
}

#[test]
fn the_range_starts_at_the_first_row_with_content() {
    let connection = connect();
    let sheet = [
        row(1, &[]),
        row(2, &[number("A2", ""), number("B2", "")]),
        row(
            4,
            &[
                number("B4", ""),
                inline("C4", "x"),
                inline("D4", "y"),
                number("E4", ""),
                inline("F4", "z"),
            ],
        ),
        row(
            5,
            &[
                number("A5", 0),
                number("C5", 1),
                number("D5", 2),
                number("F5", 3),
            ],
        ),
        row(6, &[number("C6", 4), number("D6", 5)]),
    ]
    .concat();
    let book = write("range_sniffing.xlsx", &sheet);
    let sql = format!("FROM read_xlsx('{}')", book.path());
    assert_eq!(describe(&connection, &sql), ["x DOUBLE", "y DOUBLE"]);
    assert_eq!(query(&connection, &sql), ["1.0|2.0", "4.0|5.0"]);
}

#[test]
fn the_header_is_detected_forced_or_forbidden() {
    let connection = connect();
    let sheet = [
        row(1, &[inline("A1", "name"), inline("B1", "value")]),
        row(2, &[inline("A2", "a"), number("B2", 1)]),
        row(3, &[inline("A3", "b"), number("B3", 2)]),
    ]
    .concat();
    let book = write("header_modes.xlsx", &sheet);
    let path = book.path();
    let read = |arguments: &str| {
        let sql = format!("FROM read_xlsx('{path}'{arguments})");
        (describe(&connection, &sql), query(&connection, &sql))
    };
    let (columns, rows) = read("");
    assert_eq!(columns, ["name VARCHAR", "value DOUBLE"]);
    assert_eq!(rows, ["a|1.0", "b|2.0"]);
    assert_eq!(read(", header = true"), (columns, rows));

    let (columns, rows) = read(", header = false");
    assert_eq!(columns, ["A1 VARCHAR", "B1 VARCHAR"]);
    assert_eq!(rows, ["name|value", "a|1", "b|2"]);

    let (columns, rows) = read(", range = 'A2:B3'");
    assert_eq!(columns, ["A VARCHAR", "B DOUBLE"]);
    assert_eq!(rows, ["a|1.0", "b|2.0"]);

    let (columns, rows) = read(", range = 'A2:B3', header = true");
    assert_eq!(columns, ["a VARCHAR", "1 DOUBLE"]);
    assert_eq!(rows, ["b|2.0"]);
}

#[test]
fn a_first_row_with_cells_that_are_not_strings_is_not_a_header() {
    let connection = connect();
    for (name, first_row) in [
        ("numbers", row(1, &[inline("A1", "a"), number("B1", 7)])),
        (
            "formula",
            row(1, &[inline("A1", "a"), typed("B1", "str", "b")]),
        ),
        (
            "boolean",
            row(1, &[inline("A1", "a"), typed("B1", "b", "1")]),
        ),
    ] {
        let sheet = format!(
            "{first_row}{}",
            row(2, &[inline("A2", "x"), inline("B2", "y")])
        );
        let book = write(&format!("no_header_{name}.xlsx"), &sheet);
        let columns = describe(&connection, &format!("FROM read_xlsx('{}')", book.path()));
        assert_eq!(columns[0], "A1 VARCHAR", "{name}");
        assert!(columns[1].starts_with("B1 "), "{name}");
    }
}

#[test]
fn an_empty_cell_ends_the_sniffed_range() {
    let connection = connect();
    let sheet = [
        row(1, &[inline("A1", "a"), inline("B1", ""), inline("C1", "c")]),
        row(2, &[number("A2", 1), number("B2", 2), number("C2", 3)]),
    ]
    .concat();
    let book = write("range_end.xlsx", &sheet);
    let sql = format!("FROM read_xlsx('{}')", book.path());
    assert_eq!(describe(&connection, &sql), ["a DOUBLE"]);
    assert_eq!(query(&connection, &sql), ["1.0"]);
    let whole = format!("FROM read_xlsx('{}', range = 'A1:C2')", book.path());
    assert_eq!(
        describe(&connection, &whole),
        ["A VARCHAR", "B VARCHAR", "C VARCHAR"]
    );
    assert_eq!(query(&connection, &whole), ["a||c", "1|2|3"]);
}

#[test]
fn shared_strings_name_the_columns() {
    let connection = connect();
    let sheet = [
        row(1, &[shared("A1", 2), shared("B1", 0), shared("C1", 2)]),
        row(2, &[number("A2", 1), shared("B2", 1), number("C2", 3)]),
    ]
    .concat();
    let book = Workbook::new()
        .sheet("S", &sheet)
        .shared_strings(&["second", "text", "first &amp; last"])
        .write("shared_header.xlsx");
    let sql = format!("FROM read_xlsx('{}')", book.path());
    assert_eq!(
        describe(&connection, &sql),
        [
            "first & last DOUBLE",
            "second VARCHAR",
            "first & last_1 DOUBLE"
        ]
    );
    assert_eq!(query(&connection, &sql), ["1.0|text|3.0"]);
}

#[test]
fn column_types_come_from_the_first_row_of_data() {
    let connection = connect();
    let header: Vec<String> = "ABCDEFGHIJKL"
        .chars()
        .map(|column| inline(&format!("{column}1"), &format!("c{column}")))
        .collect();
    let sheet = [
        row(1, &header),
        row(
            2,
            &[
                number("A2", 1.5),
                styled("B2", 44197, 1),
                styled("C2", 0.5, 2),
                styled("D2", 44197.5, 3),
                styled("E2", 1, 4),
                typed("F2", "b", "1"),
                inline("G2", "text"),
                typed("H2", "str", "formula"),
                typed("I2", "e", "#N/A"),
                typed("J2", "d", "2021-01-01"),
                inline("K2", "TRUE"),
                styled("L2", 7, 99),
            ],
        ),
    ]
    .concat();
    let book = write("column_types.xlsx", &sheet);
    let sql = format!("FROM read_xlsx('{}')", book.path());
    assert_eq!(
        describe(&connection, &sql),
        [
            "cA DOUBLE",
            "cB DATE",
            "cC TIME",
            "cD TIMESTAMP",
            "cE DOUBLE",
            "cF BOOLEAN",
            "cG VARCHAR",
            "cH VARCHAR",
            "cI VARCHAR",
            "cJ DATE",
            "cK BOOLEAN",
            "cL DOUBLE"
        ]
    );
    assert_eq!(
        query(&connection, &sql),
        [
            "1.5|2021-01-01|12:00:00|2021-01-01 12:00:00|1.0|true|text|formula|#N/A|2021-01-01|true|7.0"
        ]
    );
    let all_varchar = format!("FROM read_xlsx('{}', all_varchar = true)", book.path());
    assert!(
        describe(&connection, &all_varchar)
            .iter()
            .all(|column| column.ends_with(" VARCHAR"))
    );
    assert_eq!(
        query(&connection, &all_varchar),
        ["1.5|44197|0.5|44197.5|1|1|text|formula|#N/A|2021-01-01|TRUE|7"]
    );
}

#[test]
fn a_cell_of_an_unknown_type_cannot_give_a_column_its_type() {
    let connection = connect();
    let sheet = [
        row(1, &[inline("A1", "h")]),
        row(2, &[typed("A2", "zzz", "1")]),
    ]
    .concat();
    let book = write("unknown_type.xlsx", &sheet);
    let path = book.path();
    assert_eq!(
        query_error(&connection, &format!("FROM read_xlsx('{path}')")),
        "Binder Error: Unknown cell type in xlsx file"
    );
    assert_eq!(
        query(
            &connection,
            &format!("FROM read_xlsx('{path}', all_varchar = true)")
        ),
        ["1"]
    );
}

#[test]
fn missing_cells_take_the_default_type() {
    let connection = connect();
    let sheet = [
        row(
            1,
            &[inline("A1", "a"), inline("B1", "b"), inline("C1", "c")],
        ),
        row(2, &[number("A2", 1), number("C2", 3)]),
    ]
    .concat();
    let book = write("default_type.xlsx", &sheet);
    let path = book.path();
    assert_eq!(
        describe(&connection, &format!("FROM read_xlsx('{path}')")),
        ["a DOUBLE", "b DOUBLE", "c DOUBLE"]
    );
    assert_eq!(
        describe(
            &connection,
            &format!("FROM read_xlsx('{path}', empty_as_varchar = true)")
        ),
        ["a DOUBLE", "b VARCHAR", "c DOUBLE"]
    );
}

#[test]
fn a_header_without_data_makes_an_empty_table() {
    let connection = connect();
    let book = write(
        "header_only.xlsx",
        &row(1, &[inline("A1", "a"), inline("B1", "b")]),
    );
    let path = book.path();
    let sql = format!("FROM read_xlsx('{path}')");
    assert_eq!(describe(&connection, &sql), ["a DOUBLE", "b DOUBLE"]);
    assert!(query(&connection, &sql).is_empty());
    assert_eq!(
        describe(
            &connection,
            &format!("FROM read_xlsx('{path}', empty_as_varchar = true)")
        ),
        ["a VARCHAR", "b VARCHAR"]
    );
}

#[test]
fn a_sheet_without_rows_cannot_be_read_without_a_range() {
    let connection = connect();
    let book = write("empty_sheet.xlsx", "");
    let path = book.path();
    assert_eq!(
        query_error(&connection, &format!("FROM read_xlsx('{path}')")),
        "Binder Error: No rows found in xlsx file"
    );
    let sql = format!("FROM read_xlsx('{path}', range = 'B2:C3')");
    assert_eq!(describe(&connection, &sql), ["B2 DOUBLE", "C2 DOUBLE"]);
    assert_eq!(query(&connection, &sql), ["NULL|NULL", "NULL|NULL"]);
}

#[test]
fn column_names_are_trimmed_and_made_unique() {
    let connection = connect();
    let names = [
        "  padded  ",
        "dup",
        "DUP",
        "dup",
        "\u{a0}nbsp\u{3000}",
        "dup_1",
        "\ttab\t",
    ];
    let header: Vec<String> = names
        .iter()
        .zip("ABCDEFG".chars())
        .map(|(name, column)| inline(&format!("{column}1"), name))
        .collect();
    let values: Vec<String> = "ABCDEFG"
        .chars()
        .map(|column| number(&format!("{column}2"), 1))
        .collect();
    let book = write(
        "column_names.xlsx",
        &[row(1, &header), row(2, &values)].concat(),
    );
    let columns = query(
        &connection,
        &format!(
            "SELECT column_name FROM (DESCRIBE FROM read_xlsx('{}'))",
            book.path()
        ),
    );
    assert_eq!(
        columns,
        [
            "padded", "dup", "DUP_1", "dup_2", "nbsp", "dup_1_1", "\ttab\t"
        ]
    );
}

#[test]
fn column_names_can_be_normalized() {
    let connection = connect();
    let names = [
        "Hello World",
        "Ünï cödé",
        "select",
        "1st",
        "???",
        "a  b\tc",
        "Hello World",
        "ｗｉｄｅ",
    ];
    let header: Vec<String> = names
        .iter()
        .zip("ABCDEFGH".chars())
        .map(|(name, column)| inline(&format!("{column}1"), name))
        .collect();
    let values: Vec<String> = "ABCDEFGH"
        .chars()
        .map(|column| number(&format!("{column}2"), 1))
        .collect();
    let book = write(
        "normalized_names.xlsx",
        &[row(1, &header), row(2, &values)].concat(),
    );
    let columns = query(
        &connection,
        &format!(
            "SELECT column_name FROM (DESCRIBE FROM read_xlsx('{}', normalize_names = true))",
            book.path()
        ),
    );
    assert_eq!(
        columns,
        [
            "hello_world",
            "uni_code",
            "_select",
            "_1st",
            "_",
            "a_b_c",
            "hello_world_1",
            "wide"
        ]
    );
    let generated = query(
        &connection,
        &format!(
            "SELECT column_name FROM (DESCRIBE FROM read_xlsx('{}', normalize_names = true, header = false, range = 'A2:B2'))",
            book.path()
        ),
    );
    assert_eq!(generated, ["a", "b"]);
}

#[test]
fn cells_that_go_back_in_the_first_rows_are_rejected() {
    let connection = connect();
    let sheet = [
        row(1, &[inline("B1", "b"), inline("A1", "a")]),
        row(2, &[number("A2", 1), number("B2", 2)]),
    ]
    .concat();
    let book = write("backwards_header.xlsx", &sheet);
    let error = query_error(
        &connection,
        &format!("FROM read_xlsx('{}', range = 'A1:B2')", book.path()),
    );
    assert_eq!(
        error,
        "Invalid Input Error: Cell 'A1' does not follow the cell before it in its row (is the file corrupted?)"
    );
}
