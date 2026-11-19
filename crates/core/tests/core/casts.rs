//! How cell text becomes values of the column types, and what happens to cells that do not convert.

use crate::support::{
    Workbook, connect, execute, execute_error, inline, number, numbered_rows, query, query_error,
    row, shared, styled, typed,
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
fn a_failure_is_reported_with_the_reason_of_its_cell() {
    let connection = connect();
    // An empty string fails to convert without a word, before the cell that is reported.
    let sheet = headed(&[
        row(2, &[number("A2", 1), number("B2", 2)]),
        row(3, &[inline("A3", ""), number("B3", 3)]),
        row(4, &[inline("A4", "bad"), number("B4", 4)]),
        row(5, &[inline("A5", "worse"), inline("B5", "also bad")]),
    ]);
    let book = write("own_reason.xlsx", &sheet);
    let path = book.path();
    assert_eq!(
        query_error(&connection, &format!("FROM read_xlsx('{path}')")),
        "Invalid Input Error: read_xlsx: Failed to parse cell 'A4': Could not convert string 'bad' to DOUBLE"
    );
    assert_eq!(
        query(
            &connection,
            &format!("FROM read_xlsx('{path}', range = 'A1:B3')")
        ),
        ["1.0|2.0", "NULL|3.0"]
    );
    assert_eq!(
        query(
            &connection,
            &format!("FROM read_xlsx('{path}', ignore_errors = true)")
        ),
        ["1.0|2.0", "NULL|3.0", "NULL|4.0", "NULL|NULL"]
    );
}

#[test]
fn the_reason_is_not_that_of_a_cell_in_another_column() {
    let connection = connect();
    let sheet = [
        row(
            1,
            &[inline("A1", "n"), inline("B1", "flag"), inline("C1", "day")],
        ),
        row(
            2,
            &[
                number("A2", 1),
                typed("B2", "b", "1"),
                styled("C2", 44197, 1),
            ],
        ),
        row(
            3,
            &[inline("A3", ""), typed("B3", "b", "0"), inline("C3", "")],
        ),
        row(
            4,
            &[
                number("A4", 4),
                inline("B4", "maybe"),
                styled("C4", 44198, 1),
            ],
        ),
        row(
            5,
            &[number("A5", 5), typed("B5", "b", "1"), inline("C5", "soon")],
        ),
    ]
    .concat();
    let book = write("reason_by_column.xlsx", &sheet);
    let path = book.path();
    assert_eq!(
        query_error(&connection, &format!("FROM read_xlsx('{path}')")),
        "Invalid Input Error: read_xlsx: Failed to parse cell 'B4': Could not convert string 'maybe' to BOOL"
    );
    // A day is a number first, so that is what its text fails to be.
    assert_eq!(
        query_error(
            &connection,
            &format!("FROM read_xlsx('{path}', range = 'C1:C5')")
        ),
        "Invalid Input Error: read_xlsx: Failed to parse cell 'C5': Could not convert string 'soon' to DOUBLE"
    );
    assert_eq!(
        query(
            &connection,
            &format!("FROM read_xlsx('{path}', ignore_errors = true)")
        ),
        [
            "1.0|true|2021-01-01",
            "NULL|false|NULL",
            "4.0|NULL|2021-01-02",
            "5.0|true|NULL"
        ]
    );
}

#[test]
fn the_reason_is_not_that_of_a_cell_in_an_earlier_batch() {
    let connection = connect();
    let sheet = [
        row(1, &[inline("C1", "id"), inline("D1", "label")]),
        row(2, &[number("C2", 2), inline("D2", "x")]),
        row(3, &[inline("C3", ""), inline("D3", "x")]),
        (4..=7000)
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
        row(7001, &[inline("C7001", "late"), inline("D7001", "x")]),
    ]
    .concat();
    let book = write("reason_by_batch.xlsx", &sheet);
    let path = book.path();
    for select in ["*", "count(*)", "label"] {
        assert_eq!(
            query_error(
                &connection,
                &format!("SELECT {select} FROM read_xlsx('{path}')")
            ),
            "Invalid Input Error: read_xlsx: Failed to parse cell 'C7001': Could not convert string 'late' to DOUBLE",
            "{select}"
        );
    }
    execute(&connection, "CREATE TABLE ids(id SMALLINT, label VARCHAR)");
    assert_eq!(
        execute_error(
            &connection,
            &format!("COPY ids FROM '{path}' (FORMAT xlsx)")
        ),
        "Invalid Input Error: read_xlsx: Failed to parse cell 'C7001': Could not convert string 'late' to INT16"
    );
    execute(
        &connection,
        &format!("COPY ids FROM '{path}' (FORMAT xlsx, ignore_errors)"),
    );
    assert_eq!(
        query(&connection, "SELECT count(*), count(id), max(id) FROM ids"),
        ["7000|6998|7000"]
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
            "2021-01-01|12:00:00|2021-01-01 12:34:55.999978",
            "2021-01-01|06:00:00|1970-01-01 00:00:00",
            "1899-12-29|18:00:00|1899-12-28 12:00:00",
            "1899-12-30|00:00:00.999994|1899-12-30 00:00:00.00864",
            "9999-12-31|24:00:00|9999-12-31 23:59:59",
            "NULL|NULL|NULL"
        ]
    );
}

#[test]
fn serial_numbers_that_are_no_value_of_the_column_type_do_not_convert() {
    let connection = connect();
    let header = row(
        1,
        &[inline("A1", "d"), inline("B1", "t"), inline("C1", "ts")],
    );
    let triple = |number: u32, serial: &str| {
        row(
            number,
            &[
                styled(&format!("A{number}"), serial, 1),
                styled(&format!("B{number}"), serial, 2),
                styled(&format!("C{number}"), serial, 3),
            ],
        )
    };
    let rows = [
        triple(2, "44197.5"),
        triple(3, "-1e9"),
        triple(4, "1e300"),
        triple(5, "nan"),
        triple(6, "inf"),
        triple(7, "3e9"),
        triple(8, "44198.25"),
    ];
    let book = write("serial_range.xlsx", &format!("{header}{}", rows.concat()));
    let path = book.path();
    let failure = |range: &str| {
        query_error(
            &connection,
            &format!("FROM read_xlsx('{path}', range = '{range}', header = false)"),
        )
    };
    let reason = "Invalid Input Error: read_xlsx: Failed to parse cell";
    for (range, cell_and_reason) in [
        (
            "C3:C3",
            "'C3': Could not convert the serial number -1000000000.0 to TIMESTAMP: the value is out of range",
        ),
        (
            "A4:A4",
            "'A4': Could not convert the serial number 1e300 to DATE: the value is out of range",
        ),
        (
            "B4:B4",
            "'B4': Could not convert the serial number 1e300 to TIME: the value is out of range",
        ),
        (
            "A5:A5",
            "'A5': Could not convert the serial number NaN to DATE: the value is not a finite number",
        ),
        (
            "B6:B6",
            "'B6': Could not convert the serial number inf to TIME: the value is not a finite number",
        ),
        (
            "C5:C5",
            "'C5': Could not convert the serial number NaN to TIMESTAMP: the value is not a finite number",
        ),
        (
            "A7:A7",
            "'A7': Could not convert the serial number 3000000000.0 to DATE: the value is out of range",
        ),
        // The first column with such a cell reports its first one.
        (
            "A2:C8",
            "'A4': Could not convert the serial number 1e300 to DATE: the value is out of range",
        ),
        (
            "B2:C8",
            "'B4': Could not convert the serial number 1e300 to TIME: the value is out of range",
        ),
    ] {
        assert_eq!(
            failure(range),
            format!("{reason} {cell_and_reason}"),
            "{range}"
        );
    }
    // A billion days back is a day and a time of day, but no TIMESTAMP; an infinite number is an infinite
    // day and timestamp, but no time of day.
    assert_eq!(
        query(
            &connection,
            &format!("FROM read_xlsx('{path}', ignore_errors = true)")
        ),
        [
            "2021-01-01|12:00:00|2021-01-01 12:00:00",
            "2736009-12-27 (BC)|00:00:00|NULL",
            "NULL|NULL|NULL",
            "NULL|NULL|NULL",
            "infinity|NULL|infinity",
            "NULL|00:00:00|NULL",
            "2021-01-02|06:00:00|2021-01-02 06:00:00"
        ]
    );
    assert_eq!(
        query(
            &connection,
            &format!(
                "SELECT count(*), count(d), count(t), count(ts) FROM read_xlsx('{path}', ignore_errors = true)"
            )
        ),
        ["7|4|4|3"]
    );
}

#[test]
fn the_values_a_serial_number_becomes_are_ones_duckdb_accepts() {
    let connection = connect();
    let sheet = [
        row(1, &[inline("A1", "d"), inline("B1", "ts")]),
        row(2, &[styled("A2", 44197, 1), styled("B2", 44197, 3)]),
        // The last day and the last whole day of the types, and the first ones.
        row(
            3,
            &[styled("A3", 2147509215_i64, 1), styled("B3", 106777560, 3)],
        ),
        row(
            4,
            &[
                styled("A4", -2147458077_i64, 1),
                styled("B4", -106726422, 3),
            ],
        ),
    ]
    .concat();
    let book = write("serial_ends.xlsx", &sheet);
    let path = book.path();
    assert_eq!(
        query(
            &connection,
            &format!(
                "SELECT d::VARCHAR, ts::VARCHAR, isfinite(d), isfinite(ts) FROM read_xlsx('{path}')"
            )
        ),
        [
            "2021-01-01|2021-01-01 00:00:00|true|true",
            "5881580-07-10|294247-01-10 00:00:00|true|true",
            "5877642-06-25 (BC)|290309-12-22 (BC) 00:00:00|true|true"
        ]
    );
    for (cell, serial, name) in [
        ("A3", "2147509216", "DATE"),
        ("A3", "-2147458078", "DATE"),
        ("B3", "106777561", "TIMESTAMP"),
        ("B3", "-106726423", "TIMESTAMP"),
    ] {
        let column = &cell[..1];
        let style = if name == "DATE" { 1 } else { 3 };
        let sheet = [
            row(1, &[inline(&format!("{column}1"), "x")]),
            row(2, &[styled(&format!("{column}2"), 44197, style)]),
            row(3, &[styled(cell, serial, style)]),
        ]
        .concat();
        let book = write("serial_past_the_end.xlsx", &sheet);
        let error = query_error(&connection, &format!("FROM read_xlsx('{}')", book.path()));
        assert_eq!(
            error,
            format!(
                "Invalid Input Error: read_xlsx: Failed to parse cell '{cell}': Could not convert the serial number \
                 {serial}.0 to {name}: the value is out of range"
            )
        );
    }
}

fn temporal_header() -> String {
    row(
        1,
        &[inline("A1", "d"), inline("B1", "t"), inline("C1", "ts")],
    )
}

fn temporal_row(number: u32, serial: &str) -> String {
    row(
        number,
        &[
            styled(&format!("A{number}"), serial, 1),
            styled(&format!("B{number}"), serial, 2),
            styled(&format!("C{number}"), serial, 3),
        ],
    )
}

fn temporal_sheet(serials: &[&str]) -> String {
    let rows: String = (2..)
        .zip(serials)
        .map(|(number, serial)| temporal_row(number, serial))
        .collect();
    format!("{}{rows}", temporal_header())
}

#[test]
fn serial_numbers_below_61_are_the_days_a_spreadsheet_shows() {
    let connection = connect();
    let serials = [
        "1", "2", "31.25", "59", "59.75", "60", "60.5", "61", "0.5", "0", "-1.5", "45000",
    ];
    let book = write("serials_below_61.xlsx", &temporal_sheet(&serials));
    assert_eq!(
        query(&connection, &format!("FROM read_xlsx('{}')", book.path())),
        [
            "1900-01-01|24:00:00|1900-01-01 00:00:00",
            "1900-01-02|00:00:00|1900-01-02 00:00:00",
            "1900-01-31|06:00:00|1900-01-31 06:00:00",
            "1900-02-28|00:00:00|1900-02-28 00:00:00",
            "1900-02-28|18:00:00|1900-02-28 18:00:00",
            // The 29th of February 1900 that spreadsheets count is the day before it.
            "1900-02-28|00:00:00|1900-02-28 00:00:00",
            "1900-02-28|12:00:00|1900-02-28 12:00:00",
            "1900-03-01|00:00:00|1900-03-01 00:00:00",
            "1899-12-30|12:00:00|1899-12-30 12:00:00",
            "1899-12-30|00:00:00|1899-12-30 00:00:00",
            "1899-12-28|12:00:00|1899-12-28 12:00:00",
            "2023-03-15|00:00:00|2023-03-15 00:00:00"
        ]
    );
}

#[test]
fn a_workbook_that_counts_from_1904_has_its_dates_four_years_later() {
    let connection = connect();
    let serials = ["0", "1", "59.5", "60", "100.25", "43538.5", "-1"];
    let sheet = temporal_sheet(&serials);
    let expected = [
        "1904-01-01|00:00:00|1904-01-01 00:00:00",
        "1904-01-02|24:00:00|1904-01-02 00:00:00",
        "1904-02-29|12:00:00|1904-02-29 12:00:00",
        "1904-03-01|00:00:00|1904-03-01 00:00:00",
        "1904-04-10|06:00:00|1904-04-10 06:00:00",
        "2023-03-15|12:00:00|2023-03-15 12:00:00",
        "1903-12-31|00:00:00|1903-12-31 00:00:00",
    ];
    for value in ["1", "true", "on"] {
        let book = Workbook::new()
            .sheet("S", &sheet)
            .properties(&format!("date1904=\"{value}\""))
            .write("date1904.xlsx");
        let path = book.path();
        assert_eq!(
            query(&connection, &format!("FROM read_xlsx('{path}')")),
            expected,
            "{value}"
        );
        // As text the cells are the numbers they hold.
        assert_eq!(
            query(
                &connection,
                &format!("SELECT d FROM read_xlsx('{path}', all_varchar = true) LIMIT 2")
            ),
            ["0", "1"]
        );
    }
    for value in ["0", "false", "off", "TRUE", ""] {
        let book = Workbook::new()
            .sheet("S", &sheet)
            .properties(&format!("date1904=\"{value}\""))
            .write("date1900.xlsx");
        assert_eq!(
            query(
                &connection,
                &format!("FROM read_xlsx('{}') LIMIT 2", book.path())
            ),
            [
                "1899-12-30|00:00:00|1899-12-30 00:00:00",
                "1900-01-01|24:00:00|1900-01-01 00:00:00"
            ],
            "{value}"
        );
    }
}

#[test]
fn times_of_day_are_rounded_to_the_microsecond() {
    let connection = connect();
    let serials = [
        // 12:34:56.789 on 2024-03-15, the number next to it, and 12:34:56.789123.
        "45366.5242683912",
        "45366.52426839121",
        "45366.524268392626",
        // A thousandth of a second after midnight and before the next second.
        "45366.00000001157407",
        "45366.00001156250000",
        // One step of a DOUBLE around noon, and a microsecond around it.
        "45366.49999999999",
        "45366.50000000001",
        "45366.50000000001157",
        "45366.49999999998843",
        // The last millisecond of the 28th of February 1900.
        "59.99999998842593",
        "0.5242683912037037",
    ];
    let book = write("serial_fractions.xlsx", &temporal_sheet(&serials));
    assert_eq!(
        query(
            &connection,
            &format!("SELECT ts, t FROM read_xlsx('{}')", book.path())
        ),
        [
            "2024-03-15 12:34:56.789|12:34:56.789",
            "2024-03-15 12:34:56.789|12:34:56.789",
            "2024-03-15 12:34:56.789123|12:34:56.789123",
            "2024-03-15 00:00:00.001|00:00:00.001",
            "2024-03-15 00:00:00.999|00:00:00.999",
            "2024-03-15 12:00:00|12:00:00",
            "2024-03-15 12:00:00|12:00:00",
            "2024-03-15 12:00:00.000001|12:00:00.000001",
            "2024-03-15 11:59:59.999999|11:59:59.999999",
            "1900-02-28 23:59:59.999|23:59:59.999",
            "1899-12-30 12:34:56.789|12:34:56.789"
        ]
    );
}

#[test]
fn infinite_numbers_in_date_columns_are_infinite_dates() {
    let connection = connect();
    let header = row(
        1,
        &[inline("A1", "d"), inline("B1", "ts"), inline("C1", "t")],
    );
    let first = row(
        2,
        &[
            styled("A2", 44197, 1),
            styled("B2", 44197.5, 3),
            styled("C2", 0.5, 2),
        ],
    );
    // As a number cell, a text cell or a formula has it, in the spellings DuckDB reads as infinite.
    let spellings = [
        "infinity",
        "-infinity",
        "inf",
        "-inf",
        "Infinity",
        "-INF",
        " inf ",
    ];
    let rows: String = (3..)
        .zip(spellings)
        .map(|(number, spelling)| {
            row(
                number,
                &[
                    inline(&format!("A{number}"), spelling),
                    typed(&format!("B{number}"), "str", spelling),
                    styled(&format!("C{number}"), 0.25, 2),
                ],
            )
        })
        .collect();
    let last = row(
        10,
        &[
            styled("A10", "inf", 1),
            styled("B10", "-inf", 3),
            styled("C10", 0.75, 2),
        ],
    );
    let book = write(
        "infinite_dates.xlsx",
        &format!("{header}{first}{rows}{last}"),
    );
    let path = book.path();
    let sql = format!("FROM read_xlsx('{path}')");
    assert_eq!(
        crate::support::describe(&connection, &sql),
        ["d DATE", "ts TIMESTAMP", "t TIME"]
    );
    assert_eq!(
        query(
            &connection,
            &format!(
                "SELECT d, ts, isinf(d), isinf(ts), d > DATE '9999-12-31', ts < TIMESTAMP '0001-01-01' {sql}"
            )
        ),
        [
            "2021-01-01|2021-01-01 12:00:00|false|false|false|false",
            "infinity|infinity|true|true|true|false",
            "-infinity|-infinity|true|true|false|true",
            "infinity|infinity|true|true|true|false",
            "-infinity|-infinity|true|true|false|true",
            "infinity|infinity|true|true|true|false",
            "-infinity|-infinity|true|true|false|true",
            "infinity|infinity|true|true|true|false",
            "infinity|-infinity|true|true|true|true"
        ]
    );

    // A time of day is never infinite, and what is not a number is no date.
    let sheet = [
        row(1, &[inline("A1", "d"), inline("B1", "t")]),
        row(2, &[styled("A2", 44197, 1), styled("B2", 0.5, 2)]),
        row(3, &[inline("A3", "nan"), inline("B3", "infinity")]),
        row(4, &[inline("A4", "infinite"), inline("B4", "-inf")]),
    ]
    .concat();
    let book = write("not_infinite.xlsx", &sheet);
    let path = book.path();
    let failure = |range: &str| {
        query_error(
            &connection,
            &format!("FROM read_xlsx('{path}', range = '{range}', header = false)"),
        )
    };
    let reason = "Invalid Input Error: read_xlsx: Failed to parse cell";
    assert_eq!(
        failure("A2:A3"),
        format!(
            "{reason} 'A3': Could not convert the serial number NaN to DATE: the value is not a finite number"
        )
    );
    assert_eq!(
        failure("B2:B3"),
        format!(
            "{reason} 'B3': Could not convert the serial number inf to TIME: the value is not a finite number"
        )
    );
    assert_eq!(
        failure("A2:A4"),
        format!("{reason} 'A4': Could not convert string 'infinite' to DOUBLE")
    );
    assert_eq!(
        query(
            &connection,
            &format!("FROM read_xlsx('{path}', ignore_errors = true)")
        ),
        ["2021-01-01|12:00:00", "NULL|NULL", "NULL|NULL"]
    );
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
    let sheet = |day: String| {
        [
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
            row(3, &[inline("A3", "y"), day, styled("C3", "1e300", 2)]),
        ]
        .concat()
    };
    let text_for_a_day = write("unasked_columns.xlsx", &sheet(inline("B3", "soon")));
    let serial_out_of_range = write("unasked_serials.xlsx", &sheet(styled("B3", 44198, 1)));
    for select in ["count(*)", "label", "day", "*"] {
        assert_eq!(
            query_error(
                &connection,
                &format!(
                    "SELECT {select} FROM read_xlsx('{}')",
                    text_for_a_day.path()
                )
            ),
            "Invalid Input Error: read_xlsx: Failed to parse cell 'B3': Could not convert string 'soon' to DOUBLE",
            "{select}"
        );
        assert_eq!(
            query_error(
                &connection,
                &format!(
                    "SELECT {select} FROM read_xlsx('{}')",
                    serial_out_of_range.path()
                )
            ),
            "Invalid Input Error: read_xlsx: Failed to parse cell 'C3': Could not convert the serial number 1e300 \
             to TIME: the value is out of range",
            "{select}"
        );
    }
    let path = text_for_a_day.path();
    assert_eq!(
        query(
            &connection,
            &format!(
                "SELECT count(*), count(day), count(time) FROM read_xlsx('{path}', ignore_errors = true)"
            )
        ),
        ["2|1|1"]
    );
    assert_eq!(
        query(
            &connection,
            &format!("SELECT count(*) FROM read_xlsx('{path}', all_varchar = true)")
        ),
        ["2"]
    );
}
