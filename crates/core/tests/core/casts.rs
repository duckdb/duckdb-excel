//! How cell text becomes values of the column types, and what happens to cells that do not convert.

use crate::support::{
    Workbook, connect, inline, number, numbered_rows, query, query_error, row, shared, styled,
    typed,
};

fn write(name: &str, sheet_data: &str) -> crate::support::TempFile {
    Workbook::new().sheet("S", sheet_data).write(name)
}

fn headed(rows: &[String]) -> String {
    let header = row(1, &[inline("A1", "a"), inline("B1", "b")]);
    format!("{header}{}", rows.concat())
}

#[test]
fn a_cell_that_does_not_convert_fails_the_scan() {
    let connection = connect();
    let sheet = headed(&[
        row(2, &[number("A2", 1), number("B2", 2)]),
        row(3, &[number("A3", 3), inline("B3", "oops")]),
        row(4, &[inline("A4", "bad"), number("B4", 4)]),
    ]);
    let book = write("cast_failure.xlsx", &sheet);
    let path = book.path();
    // Columns are converted one after the other, so the first column with a failure reports it.
    assert_eq!(
        query_error(&connection, &format!("FROM read_xlsx('{path}')")),
        "Invalid Input Error: read_xlsx: Failed to parse cell 'A4': Could not convert string 'bad' to DOUBLE"
    );
    assert_eq!(
        query_error(
            &connection,
            &format!("FROM read_xlsx('{path}', range = 'B1:B4')")
        ),
        "Invalid Input Error: read_xlsx: Failed to parse cell 'B3': Could not convert string 'oops' to DOUBLE"
    );
    assert_eq!(
        query(
            &connection,
            &format!("FROM read_xlsx('{path}', ignore_errors = true)")
        ),
        ["1.0|2.0", "3.0|NULL", "NULL|4.0"]
    );
    assert_eq!(
        query(
            &connection,
            &format!("FROM read_xlsx('{path}', all_varchar = true)")
        ),
        ["1|2", "3|oops", "bad|4"]
    );
}

#[test]
fn a_failure_in_a_later_batch_names_its_cell() {
    let connection = connect();
    let sheet = [
        row(1, &[inline("C1", "id"), inline("D1", "label")]),
        (2..=5000)
            .map(|index| {
                row(
                    index,
                    &[
                        number(&format!("C{index}"), index),
                        inline(&format!("D{index}"), "x"),
                    ],
                )
            })
            .collect(),
        row(5001, &[inline("C5001", "late"), inline("D5001", "x")]),
    ]
    .concat();
    let book = write("late_failure.xlsx", &sheet);
    assert_eq!(
        query_error(
            &connection,
            &format!("SELECT count(*) FROM read_xlsx('{}')", book.path())
        ),
        "Invalid Input Error: read_xlsx: Failed to parse cell 'C5001': Could not convert string 'late' to DOUBLE"
    );
}

#[test]
fn empty_strings_become_null() {
    let connection = connect();
    let sheet = headed(&[
        row(2, &[number("A2", 1), typed("B2", "b", "1")]),
        row(3, &[inline("A3", ""), inline("B3", "")]),
        row(4, &[shared("A4", 0), shared("B4", 0)]),
        row(5, &[number("A5", 5), typed("B5", "b", "0")]),
    ]);
    let book = Workbook::new()
        .sheet("S", &sheet)
        .shared_strings(&[""])
        .write("empty_strings.xlsx");
    assert_eq!(
        query(
            &connection,
            &format!("FROM read_xlsx('{}', stop_at_empty = false)", book.path())
        ),
        ["1.0|true", "NULL|NULL", "NULL|NULL", "5.0|false"]
    );
}

#[test]
fn every_failure_is_reported_with_the_first_message_of_the_scan() {
    let connection = connect();
    let sheet = headed(&[
        row(2, &[number("A2", 1), number("B2", 2)]),
        row(3, &[inline("A3", ""), number("B3", 3)]),
        row(4, &[inline("A4", "bad"), number("B4", 4)]),
    ]);
    let book = write("first_message.xlsx", &sheet);
    assert_eq!(
        query_error(&connection, &format!("FROM read_xlsx('{}')", book.path())),
        "Invalid Input Error: read_xlsx: Failed to parse cell 'A4': Could not convert string '' to DOUBLE"
    );
}

#[test]
fn numbers_with_a_date_style_are_serial_dates() {
    let connection = connect();
    let header = row(
        1,
        &[inline("A1", "d"), inline("B1", "t"), inline("C1", "ts")],
    );
    let rows = [
        row(
            2,
            &[
                styled("A2", 44197, 1),
                styled("B2", 0.5, 2),
                styled("C2", "44197.524259259", 3),
            ],
        ),
        row(
            3,
            &[
                styled("A3", 44197.99, 1),
                styled("B3", 44197.25, 2),
                styled("C3", 25569, 3),
            ],
        ),
        row(
            4,
            &[
                styled("A4", -1, 1),
                styled("B4", -0.25, 2),
                styled("C4", -1.5, 3),
            ],
        ),
        row(
            5,
            &[
                styled("A5", 0, 1),
                styled("B5", "0.000011574", 2),
                styled("C5", "0.0000001", 3),
            ],
        ),
        row(
            6,
            &[
                styled("A6", 2958465, 1),
                styled("B6", 1, 2),
                styled("C6", "2958465.999988426", 3),
            ],
        ),
        row(
            7,
            &[
                styled("A7", "", 1),
                styled("B7", "", 2),
                styled("C7", "", 3),
            ],
        ),
    ];
    let book = write("serials.xlsx", &format!("{header}{}", rows.concat()));
    assert_eq!(
        query(
            &connection,
            &format!("FROM read_xlsx('{}', stop_at_empty = false)", book.path())
        ),
        [
            "2021-01-01|12:00:00|2021-01-01 12:34:56",
            "2021-01-01|06:00:00|1970-01-01 00:00:00",
            "1899-12-29|18:00:00|1899-12-28 12:00:00",
            "1899-12-30|00:00:01|1899-12-30 00:00:00.00864",
            "9999-12-31|00:00:00|9999-12-31 23:59:59",
            "NULL|NULL|NULL"
        ]
    );
}

#[test]
fn serial_numbers_out_of_range_are_the_ends_of_the_range() {
    let connection = connect();
    let header = row(1, &[inline("A1", "d"), inline("B1", "ts")]);
    let rows = [
        row(2, &[styled("A2", 44197, 1), styled("B2", 44197, 3)]),
        row(3, &[styled("A3", "1e300", 1), styled("B3", "1e300", 3)]),
        row(4, &[styled("A4", "inf", 1), styled("B4", "inf", 3)]),
    ];
    let book = write("serial_range.xlsx", &format!("{header}{}", rows.concat()));
    let rows = query(
        &connection,
        &format!(
            "SELECT d = 'infinity'::DATE, ts = 'infinity'::TIMESTAMP FROM read_xlsx('{}')",
            book.path()
        ),
    );
    assert_eq!(rows, ["false|false", "true|true", "true|true"]);
}

#[test]
fn a_time_cannot_be_infinite() {
    let connection = connect();
    let sheet = [
        row(1, &[inline("A1", "t")]),
        row(2, &[styled("A2", 0.5, 2)]),
        row(3, &[styled("A3", "1e300", 2)]),
    ]
    .concat();
    let book = write("infinite_time.xlsx", &sheet);
    let path = book.path();
    for sql in [
        format!("FROM read_xlsx('{path}')"),
        format!("SELECT count(*) FROM read_xlsx('{path}')"),
    ] {
        assert_eq!(
            query_error(&connection, &sql),
            "Conversion Error: Can't get TIME of infinite TIMESTAMP"
        );
    }
}

#[test]
fn text_in_a_serial_date_column_is_a_failed_number() {
    let connection = connect();
    let sheet = [
        row(1, &[inline("A1", "d"), inline("B1", "ts")]),
        row(2, &[styled("A2", 44197, 1), styled("B2", 44197, 3)]),
        row(3, &[inline("A3", "2021-01-02"), styled("B3", 44198, 3)]),
    ]
    .concat();
    let book = write("text_in_dates.xlsx", &sheet);
    let path = book.path();
    assert_eq!(
        query_error(&connection, &format!("FROM read_xlsx('{path}')")),
        "Invalid Input Error: read_xlsx: Failed to parse cell 'A3': Could not convert string '2021-01-02' to DOUBLE"
    );
    assert_eq!(
        query(
            &connection,
            &format!("FROM read_xlsx('{path}', ignore_errors = true)")
        ),
        ["2021-01-01|2021-01-01 00:00:00", "NULL|2021-01-02 00:00:00"]
    );
}

#[test]
fn cells_of_type_date_and_boolean_are_cast_from_their_text() {
    let connection = connect();
    let sheet = headed(&[
        row(2, &[typed("A2", "d", "2021-01-01"), typed("B2", "b", "1")]),
        row(
            3,
            &[
                typed("A3", "d", "1999-12-31T23:59:59"),
                inline("B3", "FALSE"),
            ],
        ),
        row(4, &[typed("A4", "d", ""), inline("B4", "maybe")]),
    ]);
    let book = write("typed_cells.xlsx", &sheet);
    let path = book.path();
    assert_eq!(
        query_error(&connection, &format!("FROM read_xlsx('{path}')")),
        "Invalid Input Error: read_xlsx: Failed to parse cell 'B4': Could not convert string 'maybe' to BOOL"
    );
    assert_eq!(
        query(
            &connection,
            &format!("FROM read_xlsx('{path}', ignore_errors = true)")
        ),
        ["2021-01-01|true", "1999-12-31|false", "NULL|NULL"]
    );
}

#[test]
fn numbers_keep_their_precision() {
    let connection = connect();
    let values = [
        "0.1",
        "1e-7",
        "123456789012345678",
        "-0",
        "1.7976931348623157e308",
        "4.9e-324",
        "1E5",
        " 7 ",
    ];
    let rows: Vec<String> = values
        .iter()
        .enumerate()
        .map(|(index, value)| {
            row(
                index as u32 + 2,
                &[number(&format!("A{}", index + 2), value)],
            )
        })
        .collect();
    let book = write(
        "numbers.xlsx",
        &format!("{}{}", row(1, &[inline("A1", "n")]), rows.concat()),
    );
    assert_eq!(
        query(&connection, &format!("FROM read_xlsx('{}')", book.path())),
        [
            "0.1",
            "1e-07",
            "1.2345678901234568e+17",
            "-0.0",
            "1.7976931348623157e+308",
            "5e-324",
            "100000.0",
            "7.0"
        ]
    );
}

#[test]
fn many_text_and_number_columns_convert_side_by_side() {
    let connection = connect();
    let book = write(
        "side_by_side.xlsx",
        &format!(
            "{}{}",
            row(1, &[inline("A1", "id"), inline("B1", "label")]),
            numbered_rows(2, 2500)
        ),
    );
    let rows = query(
        &connection,
        &format!(
            "SELECT count(*), sum(id), min(label), max(length(label)) FROM read_xlsx('{}')",
            book.path()
        ),
    );
    assert_eq!(rows, ["2499|3126249.0|label 10|10"]);
}

#[test]
fn columns_a_query_does_not_ask_for_are_converted_too() {
    let connection = connect();
    let sheet = [
        row(
            1,
            &[
                inline("A1", "label"),
                inline("B1", "day"),
                inline("C1", "time"),
            ],
        ),
        row(
            2,
            &[
                inline("A2", "x"),
                styled("B2", 44197, 1),
                styled("C2", 0.5, 2),
            ],
        ),
        row(
            3,
            &[
                inline("A3", "y"),
                inline("B3", "soon"),
                styled("C3", "1e300", 2),
            ],
        ),
    ]
    .concat();
    let book = write("unasked_columns.xlsx", &sheet);
    let path = book.path();
    for select in ["count(*)", "label", "day", "*"] {
        assert_eq!(
            query_error(
                &connection,
                &format!("SELECT {select} FROM read_xlsx('{path}')")
            ),
            "Invalid Input Error: read_xlsx: Failed to parse cell 'B3': Could not convert string 'soon' to DOUBLE",
            "{select}"
        );
        assert_eq!(
            query_error(
                &connection,
                &format!("SELECT {select} FROM read_xlsx('{path}', ignore_errors = true)")
            ),
            "Conversion Error: Can't get TIME of infinite TIMESTAMP",
            "{select}"
        );
    }
    assert_eq!(
        query(
            &connection,
            &format!("SELECT count(*) FROM read_xlsx('{path}', all_varchar = true)")
        ),
        ["2"]
    );
}
