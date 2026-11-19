//! The scan of `read_xlsx`: batches, rows the sheet leaves out, ranges past the data, and scans that end early.

use crate::support::{
    Workbook, connect, inline, number, numbered_rows, query, query_error, row, shared, typed,
};

const HEADER: &str = r#"<row r="1"><c r="A1" t="inlineStr"><is><t>id</t></is></c><c r="B1" t="inlineStr"><is><t>label</t></is></c></row>"#;

fn write(name: &str, sheet_data: &str) -> crate::support::TempFile {
    Workbook::new().sheet("S", sheet_data).write(name)
}

#[test]
fn rows_come_in_sheet_order_over_many_batches() {
    let connection = connect();
    let book = write(
        "many_rows.xlsx",
        &format!("{HEADER}{}", numbered_rows(2, 10_001)),
    );
    let path = book.path();
    let rows = query(
        &connection,
        &format!(
            "SELECT count(*), min(id), max(id), sum(id), count(DISTINCT label) FROM read_xlsx('{path}')"
        ),
    );
    assert_eq!(rows, ["10000|2.0|10001.0|50015000.0|10000"]);
    let ordered = query(
        &connection,
        &format!(
            "SELECT count(*) FROM (SELECT id, row_number() OVER () AS position FROM read_xlsx('{path}')) \
             WHERE id <> position + 1 OR position IS NULL"
        ),
    );
    assert_eq!(ordered, ["0"]);
    let labels = query(
        &connection,
        &format!(
            "SELECT label FROM read_xlsx('{path}') WHERE id IN (2, 2049, 2050, 4097, 10001) ORDER BY id"
        ),
    );
    assert_eq!(
        labels,
        [
            "label 2",
            "label 2049",
            "label 2050",
            "label 4097",
            "label 10001"
        ]
    );
}

#[test]
fn batch_boundaries() {
    let connection = connect();
    for rows in [1, 2047, 2048, 2049, 4095, 4096, 4097] {
        let book = write(
            "boundary.xlsx",
            &format!("{HEADER}{}", numbered_rows(2, rows + 1)),
        );
        let counted = query(
            &connection,
            &format!(
                "SELECT count(*), count(id), max(id) FROM read_xlsx('{}')",
                book.path()
            ),
        );
        assert_eq!(counted, [format!("{rows}|{rows}|{}.0", rows + 1)]);
    }
}

#[test]
fn a_scan_stops_at_rows_the_sheet_leaves_out() {
    let connection = connect();
    let sheet = format!("{HEADER}{}{}", numbered_rows(2, 5), numbered_rows(8, 9));
    let book = write("gap.xlsx", &sheet);
    let path = book.path();
    let ids = |arguments: &str| {
        query(
            &connection,
            &format!("SELECT #1::INTEGER FROM read_xlsx('{path}'{arguments})"),
        )
    };
    assert_eq!(ids(""), ["2", "3", "4", "5"]);
    assert_eq!(ids(", stop_at_empty = true"), ["2", "3", "4", "5"]);
    assert_eq!(
        ids(", stop_at_empty = false"),
        ["2", "3", "4", "5", "NULL", "NULL", "8", "9"]
    );
    assert_eq!(
        ids(", range = 'A2:B10'"),
        ["2", "3", "4", "5", "NULL", "NULL", "8", "9", "NULL"]
    );
    assert_eq!(
        ids(", range = 'A2:B10', stop_at_empty = true"),
        ["2", "3", "4", "5"]
    );
}

#[test]
fn a_row_of_empty_cells_ends_a_scan() {
    let connection = connect();
    let sheet = [
        HEADER.to_owned(),
        numbered_rows(2, 3),
        row(4, &[number("A4", ""), inline("B4", "")]),
        numbered_rows(5, 6),
        row(7, &[]),
        numbered_rows(8, 8),
    ]
    .concat();
    let book = write("empty_row.xlsx", &sheet);
    let path = book.path();
    let ids = |arguments: &str| {
        query(
            &connection,
            &format!("SELECT id::INTEGER, label FROM read_xlsx('{path}'{arguments})"),
        )
    };
    assert_eq!(ids(""), ["2|label 2", "3|label 3"]);
    assert_eq!(
        ids(", stop_at_empty = false"),
        [
            "2|label 2",
            "3|label 3",
            "NULL|",
            "5|label 5",
            "6|label 6",
            "NULL|NULL",
            "8|label 8"
        ]
    );
}

#[test]
fn an_empty_row_element_stands_for_the_rows_left_out_before_it() {
    let connection = connect();
    let sheet = [
        row(1, &[number("A1", 1)]),
        r#"<row r="4" ht="30" customHeight="1"/>"#.to_owned(),
        row(5, &[number("A5", 5)]),
        r#"<row r="9"/>"#.to_owned(),
    ]
    .concat();
    let book = write("empty_row_element.xlsx", &sheet);
    let path = book.path();
    assert_eq!(
        query(
            &connection,
            &format!("FROM read_xlsx('{path}', header = false, stop_at_empty = false)")
        ),
        ["1.0", "NULL", "5.0", "NULL"]
    );
    assert_eq!(
        query(
            &connection,
            &format!("FROM read_xlsx('{path}', range = 'A1:A7')")
        ),
        ["1.0", "NULL", "5.0", "NULL", "NULL"]
    );
    assert_eq!(
        query(
            &connection,
            &format!("FROM read_xlsx('{path}', header = false)")
        ),
        ["1.0"]
    );
}

#[test]
fn a_range_whose_first_row_wraps_around_starts_before_the_sheet() {
    let connection = connect();
    let path = crate::support::fixture("google_sheets.xlsx");
    assert_eq!(
        query(
            &connection,
            &format!(
                "FROM read_xlsx('{path}', range = 'A18446744073709551616:B5', header = false)"
            )
        ),
        [
            "NULL|NULL",
            "ABC|HELLO",
            "123.0|500,60",
            "456.0|300,40",
            "NULL|NULL",
            "NULL|NULL"
        ]
    );
}

#[test]
fn long_gaps_span_batches() {
    let connection = connect();
    let sheet = format!(
        "{HEADER}{}{}{}",
        numbered_rows(2, 3),
        numbered_rows(5000, 5001),
        numbered_rows(9000, 9000)
    );
    let book = write("long_gaps.xlsx", &sheet);
    let path = book.path();
    let rows = query(
        &connection,
        &format!(
            "SELECT count(*), count(id), sum(id) FROM read_xlsx('{path}', stop_at_empty = false)"
        ),
    );
    assert_eq!(rows, ["8999|5|19006.0"]);
    let positions = query(
        &connection,
        &format!(
            "SELECT position FROM (SELECT id, row_number() OVER () AS position FROM read_xlsx('{path}', stop_at_empty = false)) \
             WHERE id IS NOT NULL"
        ),
    );
    assert_eq!(positions, ["1", "2", "4999", "5000", "8999"]);
    assert_eq!(
        query(
            &connection,
            &format!("SELECT count(*) FROM read_xlsx('{path}')")
        ),
        ["2"]
    );
}

#[test]
fn a_gap_right_after_a_full_batch() {
    let connection = connect();
    let sheet = format!(
        "{HEADER}{}{}",
        numbered_rows(2, 2049),
        numbered_rows(4100, 4100)
    );
    let book = write("gap_after_batch.xlsx", &sheet);
    let path = book.path();
    assert_eq!(
        query(
            &connection,
            &format!("SELECT count(*) FROM read_xlsx('{path}')")
        ),
        ["2048"]
    );
    assert_eq!(
        query(
            &connection,
            &format!("SELECT count(*), count(id) FROM read_xlsx('{path}', stop_at_empty = false)")
        ),
        ["4099|2049"]
    );
}

#[test]
fn ranges_past_the_data_are_filled_with_nulls() {
    let connection = connect();
    let book = write(
        "past_data.xlsx",
        &format!("{HEADER}{}", numbered_rows(2, 3000)),
    );
    let path = book.path();
    let count = |arguments: &str| {
        query(
            &connection,
            &format!("SELECT count(*), count(#1), count(#2) FROM read_xlsx('{path}', {arguments})"),
        )
    };
    assert_eq!(count("range = 'A1:B7000'"), ["6999|2999|2999"]);
    assert_eq!(
        count("range = 'A1:B7000', stop_at_empty = true"),
        ["2999|2999|2999"]
    );
    assert_eq!(
        count("range = 'A2990:B5049', header = false"),
        ["2060|11|11"]
    );
    assert_eq!(count("range = 'A3001:B5048', header = false"), ["2048|0|0"]);
    assert_eq!(count("range = 'A1:B2049'"), ["2048|2048|2048"]);
    let wide = query(
        &connection,
        &format!(
            "SELECT count(*), count(A), count(D) FROM read_xlsx('{path}', range = 'A2:D4200', header = false)"
        ),
    );
    assert_eq!(wide, ["4199|2999|0"]);
}

#[test]
fn only_cells_inside_the_range_are_read() {
    let connection = connect();
    let sheet = [
        row(
            1,
            &[
                number("A1", 1),
                number("B1", 2),
                number("C1", 3),
                number("D1", 4),
            ],
        ),
        row(
            2,
            &[
                number("A2", 5),
                number("B2", 6),
                number("C2", 7),
                number("D2", 8),
            ],
        ),
        row(3, &[number("A3", 9), number("D3", 12)]),
        row(4, &[number("B4", 14), number("C4", 15)]),
    ]
    .concat();
    let book = write("inside_range.xlsx", &sheet);
    let rows = query(
        &connection,
        &format!(
            "FROM read_xlsx('{}', range = 'B2:C4', header = false)",
            book.path()
        ),
    );
    assert_eq!(rows, ["6.0|7.0", "NULL|NULL", "14.0|15.0"]);
}

#[test]
fn cells_and_rows_out_of_order() {
    let connection = connect();
    let sheet = [
        row(
            1,
            &[inline("A1", "a"), inline("B1", "b"), inline("C1", "c")],
        ),
        row(2, &[number("A2", 1), number("B2", 2), number("C2", 3)]),
        row(3, &[number("C3", 1), number("A3", 2), number("B3", 3)]),
        row(4, &[number("A4", 1), number("A4", 2), number("C4", 3)]),
        row(7, &[number("A7", 7)]),
        row(5, &[number("A5", 5)]),
    ]
    .concat();
    let book = write("out_of_order.xlsx", &sheet);
    let rows = query(
        &connection,
        &format!("FROM read_xlsx('{}', stop_at_empty = false)", book.path()),
    );
    assert_eq!(
        rows,
        [
            "1.0|2.0|3.0",
            "NULL|NULL|NULL",
            "2.0|NULL|3.0",
            "NULL|NULL|NULL",
            "NULL|NULL|NULL",
            "7.0|NULL|NULL",
            "5.0|NULL|NULL"
        ]
    );
}

#[test]
fn rows_and_cells_without_references() {
    let connection = connect();
    let sheet = r#"<row><c t="inlineStr"><is><t>a</t></is></c><c t="inlineStr"><is><t>b</t></is></c></row>
        <row><c><v>1</v></c><c><v>2</v></c></row><row><c><v>3</v></c></row><row r="6"><c r="B6"><v>4</v></c></row><row><c><v>5</v></c></row>"#;
    let book = write("no_references.xlsx", sheet);
    let rows = query(
        &connection,
        &format!("FROM read_xlsx('{}', stop_at_empty = false)", book.path()),
    );
    assert_eq!(
        rows,
        [
            "1.0|2.0",
            "3.0|NULL",
            "NULL|NULL",
            "NULL|NULL",
            "NULL|4.0",
            "5.0|NULL"
        ]
    );
}

#[test]
fn shared_and_inline_strings_fill_text_columns() {
    let connection = connect();
    let long = "long shared string ".repeat(20);
    let sheet = [
        row(1, &[shared("A1", 0), shared("B1", 1)]),
        row(
            2,
            &[
                shared("A2", 2),
                inline("B2", "inline &amp; &lt;escaped&gt;"),
            ],
        ),
        row(3, &[shared("A3", 3), inline("B3", "  spaces  ")]),
        row(4, &[shared("A4", 2), inline("B4", "")]),
        row(5, &[shared("A5", 4)]),
    ]
    .concat();
    let book = Workbook::new()
        .sheet("S", &sheet)
        .shared_strings(&["first", "second", "short", &long, ""])
        .write("strings.xlsx");
    let rows = query(&connection, &format!("FROM read_xlsx('{}')", book.path()));
    assert_eq!(
        rows,
        [
            "short|inline & <escaped>".to_owned(),
            format!("{long}|  spaces  "),
            "short|".to_owned(),
            "|NULL".to_owned()
        ]
    );
}

#[test]
fn wide_sheets_come_out_cell_for_cell() {
    let connection = connect();
    let (width, height) = (300, 400);
    let name = |column: usize| crate::support::column_letters(column);
    let header: Vec<String> = (0..width)
        .map(|column| inline(&format!("{}1", name(column)), &format!("c{column}")))
        .collect();
    let rows: String = (0..height)
        .map(|index| {
            let number_of_row = index + 2;
            let cells: Vec<String> = (0..width)
                .filter(|column| index == 0 || (index + column) % 7 != 0)
                .map(|column| {
                    let reference = format!("{}{number_of_row}", name(column));
                    match column % 10 {
                        9 => inline(&reference, &format!("s{index}_{column}")),
                        8 => shared(&reference, (index + column) % 3),
                        _ => number(&reference, index * 1000 + column),
                    }
                })
                .collect();
            row(number_of_row as u32, &cells)
        })
        .collect();
    let book = Workbook::new()
        .sheet("S", &format!("{}{rows}", row(1, &header)))
        .shared_strings(&["zero", "one", "two"])
        .write("wide.xlsx");
    let path = book.path();
    for arguments in ["", ", all_varchar = true"] {
        let source = format!(
            "(SELECT row_number() OVER () - 1 AS i, * FROM read_xlsx('{path}'{arguments}))"
        );
        for column in [0, 8, 9, 299] {
            let value = match column % 10 {
                9 => format!("'s' || i || '_{column}'"),
                8 => format!("['zero', 'one', 'two'][(i + {column}) % 3 + 1]"),
                _ => format!("i * 1000 + {column}"),
            };
            let expected =
                format!("CASE WHEN i = 0 OR (i + {column}) % 7 <> 0 THEN ({value})::VARCHAR END");
            assert_eq!(
                query(
                    &connection,
                    &format!(
                        "SELECT count(*), count(c{column}), count(*) FILTER (TRY_CAST(c{column} AS BIGINT)::VARCHAR \
                         IS DISTINCT FROM {expected} AND c{column}::VARCHAR IS DISTINCT FROM {expected}) FROM {source}"
                    )
                ),
                [format!(
                    "{height}|{}|0",
                    1 + (1..height)
                        .filter(|index| (index + column) % 7 != 0)
                        .count()
                )],
                "column {column}{arguments}"
            );
        }
    }
    // Asked for one column, the others are still converted.
    let sum: usize = (0..height)
        .filter(|index| index % 7 != 0)
        .map(|index| index * 1000)
        .sum();
    assert_eq!(
        query(
            &connection,
            &format!("SELECT count(*), sum(c0)::BIGINT, max(c299) FROM read_xlsx('{path}')")
        ),
        [format!("{height}|{sum}|s99_299")]
    );
}

#[test]
fn a_shared_string_cell_without_an_index_is_null() {
    let connection = connect();
    let empty = |reference: &str| format!(r#"<c r="{reference}" t="s"/>"#);
    let valueless = |reference: &str| format!(r#"<c r="{reference}" t="s"><v></v></c>"#);
    let sheet = [
        row(1, &[shared("A1", 0), shared("B1", 1)]),
        row(2, &[shared("A2", 2), empty("B2")]),
        row(3, &[valueless("A3"), shared("B3", 2)]),
        row(4, &[empty("A4"), valueless("B4")]),
        row(5, &[shared("A5", 2), shared("B5", 2)]),
    ]
    .concat();
    let book = Workbook::new()
        .sheet("S", &sheet)
        .shared_strings(&["a", "b", "text"])
        .write("shared_without_index.xlsx");
    let path = book.path();
    // A row of nothing but such cells is an empty row, which ends the scan.
    assert_eq!(
        query(&connection, &format!("FROM read_xlsx('{path}')")),
        ["text|NULL", "NULL|text"]
    );
    assert_eq!(
        query(
            &connection,
            &format!("FROM read_xlsx('{path}', stop_at_empty = false)")
        ),
        ["text|NULL", "NULL|text", "NULL|NULL", "text|text"]
    );
    assert_eq!(
        query(
            &connection,
            &format!(
                "SELECT count(*), count(a), count(b) FROM read_xlsx('{path}', range = 'A1:B5', all_varchar = true)"
            )
        ),
        ["4|2|2"]
    );
}

#[test]
fn escapes_in_strings_are_resolved() {
    let connection = connect();
    let sheet = [
        row(1, &[inline("A1", "kind"), inline("B1", "text")]),
        row(2, &[inline("A2", "shared"), shared("B2", 0)]),
        row(3, &[inline("A3", "shared"), shared("B3", 1)]),
        row(4, &[inline("A4", "shared"), shared("B4", 2)]),
        row(5, &[inline("A5", "inline"), inline("B5", "cr_x000D_lf_x000a_")]),
        row(
            6,
            &[
                inline("A6", "inline"),
                r#"<c r="B6" t="inlineStr"><is><r><t>_x00</t></r><r><t>41_ _x0041_</t></r></is></c>"#
                    .to_owned(),
            ],
        ),
        row(
            7,
            &[inline("A7", "formula"), typed("B7", "str", "tab_x0009_bed")],
        ),
        row(
            8,
            &[inline("A8", "inline"), inline("B8", "_xD83D__xDE00_ _xD83D_ _x12G4_ _x0000_")],
        ),
        row(
            9,
            &[inline("A9", "error"), typed("B9", "e", "#N/A_x0041_")],
        ),
    ]
    .concat();
    let book = Workbook::new()
        .sheet("S", &sheet)
        .part(
            "xl/sharedStrings.xml",
            "<sst><si><t>line_x000A_break</t></si><si><t>under_x005F_score _x005F_x000D_</t></si>\
             <si><r><t>rich_x0021_</t></r><r><t>_x0021_ _x00</t></r><r><t>21_</t></r></si></sst>",
        )
        .write("escapes.xlsx");
    let shown = "replace(replace(replace(replace(text, chr(10), '<LF>'), chr(13), '<CR>'), chr(9), '<TAB>'), \
                 chr(0), '<NUL>')";
    for arguments in ["", ", all_varchar = true"] {
        assert_eq!(
            query(
                &connection,
                &format!(
                    "SELECT kind, {shown} FROM read_xlsx('{}'{arguments})",
                    book.path()
                )
            ),
            [
                "shared|line<LF>break",
                "shared|under_score _x000D_",
                "shared|rich!! _x0021_",
                "inline|cr<CR>lf<LF>",
                "inline|_x0041_ A",
                "formula|tab<TAB>bed",
                "inline|\u{1F600} _xD83D_ _x12G4_ <NUL>",
                "error|#N/A_x0041_"
            ],
            "{arguments}"
        );
    }
}

#[test]
fn escapes_in_numbers_are_not_resolved() {
    let connection = connect();
    let sheet = [
        row(1, &[inline("A1", "n")]),
        row(2, &[number("A2", 1)]),
        row(3, &[number("A3", "_x0031_")]),
    ]
    .concat();
    let book = write("escaped_number.xlsx", &sheet);
    let path = book.path();
    assert_eq!(
        query_error(&connection, &format!("FROM read_xlsx('{path}')")),
        "Invalid Input Error: read_xlsx: Failed to parse cell 'A3': Could not convert string '_x0031_' to DOUBLE"
    );
    assert_eq!(
        query(
            &connection,
            &format!("FROM read_xlsx('{path}', all_varchar = true)")
        ),
        ["1", "_x0031_"]
    );
}

#[test]
fn a_scan_that_is_only_counted_still_checks_every_column() {
    let connection = connect();
    let sheet = [
        row(
            1,
            &[
                inline("A1", "id"),
                inline("B1", "value"),
                inline("C1", "text"),
            ],
        ),
        row(2, &[number("A2", 1), number("B2", 2), inline("C2", "x")]),
        row(
            3,
            &[number("A3", 2), inline("B3", "bad"), inline("C3", "y")],
        ),
    ]
    .concat();
    let book = write("counted.xlsx", &sheet);
    let path = book.path();
    let expected = "Invalid Input Error: read_xlsx: Failed to parse cell 'B3': Could not convert string 'bad' to DOUBLE";
    for sql in [
        format!("SELECT count(*) FROM read_xlsx('{path}')"),
        format!("SELECT id FROM read_xlsx('{path}')"),
        format!("SELECT text FROM read_xlsx('{path}')"),
        format!("FROM read_xlsx('{path}')"),
    ] {
        assert_eq!(query_error(&connection, &sql), expected, "{sql}");
    }
    assert_eq!(
        query(
            &connection,
            &format!("SELECT count(*) FROM read_xlsx('{path}', ignore_errors = true)")
        ),
        ["2"]
    );
    assert_eq!(
        query(
            &connection,
            &format!("SELECT count(*), count(text) FROM read_xlsx('{path}', all_varchar = true)")
        ),
        ["2|2"]
    );
}

#[test]
fn scans_can_end_early_and_run_side_by_side() {
    let connection = connect();
    let book = write(
        "early_end.xlsx",
        &format!("{HEADER}{}", numbered_rows(2, 6000)),
    );
    let path = book.path();
    assert_eq!(
        query(
            &connection,
            &format!("SELECT id FROM read_xlsx('{path}') LIMIT 2")
        ),
        ["2.0", "3.0"]
    );
    assert!(query(&connection, &format!("FROM read_xlsx('{path}') LIMIT 0")).is_empty());
    assert_eq!(
        query(
            &connection,
            &format!("SELECT label FROM read_xlsx('{path}') LIMIT 1 OFFSET 4500")
        ),
        ["label 4502"]
    );
    let joined = query(
        &connection,
        &format!(
            "SELECT count(*), sum(a.id) FROM read_xlsx('{path}') a JOIN read_xlsx('{path}') b ON a.id = b.id + 1"
        ),
    );
    assert_eq!(joined, ["5998|18002997.0"]);
    let unioned = query(
        &connection,
        &format!(
            "SELECT count(*) FROM (SELECT id FROM read_xlsx('{path}') UNION ALL SELECT id FROM read_xlsx('{path}'))"
        ),
    );
    assert_eq!(unioned, ["11998"]);
}
