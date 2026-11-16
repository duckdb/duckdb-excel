//! `COPY ... TO 'file.xlsx' (FORMAT xlsx)` into a new workbook.

use crate::support::{
    TempDir, connect, describe, execute, execute_error, query, sheet_rows, workbook_part,
    workbook_parts,
};

fn copy(connection: &duckdb_neo::connection::Connection, rows: &str, path: &str, options: &str) {
    execute(
        connection,
        &format!("COPY ({rows}) TO '{path}' (FORMAT xlsx{options})"),
    );
}

fn copy_error(
    connection: &duckdb_neo::connection::Connection,
    rows: &str,
    path: &str,
    options: &str,
) -> String {
    execute_error(
        connection,
        &format!("COPY ({rows}) TO '{path}' (FORMAT xlsx{options})"),
    )
}

#[test]
fn a_workbook_without_options_has_one_sheet_without_a_header() {
    let connection = connect();
    let directory = TempDir::new("defaults");
    let path = directory.path("out.xlsx");
    copy(
        &connection,
        "SELECT i AS n, 'row ' || i AS label FROM range(3) t(i)",
        &path,
        "",
    );
    assert_eq!(
        query(&connection, &format!("FROM read_xlsx('{path}')")),
        ["0.0|row 0", "1.0|row 1", "2.0|row 2"]
    );
    assert_eq!(
        query(
            &connection,
            &format!("FROM read_xlsx('{path}', sheet = 'Sheet1')")
        )
        .len(),
        3
    );
    assert_eq!(
        describe(&connection, &format!("FROM read_xlsx('{path}')")),
        ["A1 DOUBLE", "B1 VARCHAR"]
    );
    assert_eq!(directory.file_names(), ["out.xlsx"]);
}

#[test]
fn the_format_follows_from_the_file_extension() {
    let connection = connect();
    let directory = TempDir::new("by_extension");
    let path = directory.path("out.xlsx");
    execute(
        &connection,
        &format!("COPY (SELECT 7 AS seven) TO '{path}'"),
    );
    execute(
        &connection,
        &format!("COPY (SELECT 8 AS eight) TO '{path}' WITH (HEADER true)"),
    );
    assert_eq!(
        describe(&connection, &format!("FROM '{path}'")),
        ["eight DOUBLE"]
    );
    assert_eq!(query(&connection, &format!("FROM '{path}'")), ["8.0"]);
}

#[test]
fn the_header_option() {
    let connection = connect();
    let directory = TempDir::new("header");
    let path = directory.path("out.xlsx");
    let rows =
        "SELECT 1 AS \"first & <1st>\", 'x' AS \"it's\", true AS \" \", DATE '2024-01-01' AS d";
    for header in [
        ", HEADER true",
        ", HEADER",
        ", header 1",
        ", HEADER 'yes'",
        ", HEADER 't'",
        ", HEADER 2.5",
    ] {
        copy(&connection, rows, &path, header);
        assert_eq!(
            sheet_rows(&path, "sheet1.xml"),
            concat!(
                r#"<row r="1"><c r="A1" t="inlineStr"><is><t>first &amp; &lt;1st&gt;</t></is></c>"#,
                r#"<c r="B1" t="inlineStr"><is><t>it&apos;s</t></is></c>"#,
                r#"<c r="C1" t="inlineStr"><is><t> </t></is></c>"#,
                r#"<c r="D1" t="inlineStr"><is><t>d</t></is></c></row>"#,
                r#"<row r="2"><c r="A2"><v>1</v></c><c r="B2" t="inlineStr"><is><t>x</t></is></c>"#,
                r#"<c r="C2" t="b" s="5"><v>1</v></c><c r="D2" s="1"><v>45292.0</v></c></row>"#
            ),
            "{header}"
        );
    }
    assert_eq!(
        query(
            &connection,
            &format!(
                "SELECT \"first & <1st>\", \"it's\", d FROM read_xlsx('{path}', header = true)"
            )
        ),
        ["1.0|x|2024-01-01"]
    );
    for header in ["", ", HEADER false", ", HEADER 0", ", HEADER 'no'"] {
        copy(&connection, rows, &path, header);
        assert!(
            sheet_rows(&path, "sheet1.xml").starts_with(r#"<row r="1"><c r="A1"><v>1</v></c>"#),
            "{header}"
        );
    }
}

#[test]
fn the_sheet_option() {
    let connection = connect();
    let directory = TempDir::new("sheet");
    let path = directory.path("out.xlsx");
    for (name, in_workbook) in [
        ("Data", "Data"),
        ("A&B <1> \"q\"", "A&amp;B &lt;1&gt; &quot;q&quot;"),
        ("it''s", "it&apos;s"),
        ("Tabelle Ä 🦆", "Tabelle Ä 🦆"),
        ("", ""),
    ] {
        copy(
            &connection,
            "SELECT 5 AS a",
            &path,
            &format!(", SHEET '{name}'"),
        );
        let workbook = workbook_part(&path, "xl/workbook.xml");
        let element = format!(
            "<sheets><sheet name=\"{in_workbook}\" state=\"visible\" sheetId=\"1\" r:id=\"rId3\"/></sheets>"
        );
        assert!(workbook.contains(&element), "{workbook}");
        if !name.is_empty() {
            assert_eq!(
                query(
                    &connection,
                    &format!("FROM read_xlsx('{path}', sheet = '{name}')")
                ),
                ["5.0"]
            );
        }
    }
}

#[test]
fn options_of_the_wrong_kind_are_binder_errors() {
    let connection = connect();
    let directory = TempDir::new("option_errors");
    let path = directory.path("out.xlsx");
    let header = "Binder Error: Header option must be a single boolean value";
    let sheet = "Binder Error: Sheet name option must be a single string value";
    let limit = "Binder Error: Sheet row limit option must be a single integer value";
    let mode = "Binder Error: MODE option must be a string ('create', 'append', or 'replace')";
    for (options, message) in [
        (", HEADER 'maybe'", header),
        (", HEADER (true, false)", header),
        (", HEADER DATE '2024-01-01'", header),
        (", SHEET", sheet),
        (", SHEET 5", sheet),
        (", SHEET NULL", sheet),
        (", SHEET ('a', 'b')", sheet),
        (", SHEET_ROW_LIMIT", limit),
        (", SHEET_ROW_LIMIT 'abc'", limit),
        (", SHEET_ROW_LIMIT NULL", limit),
        (", SHEET_ROW_LIMIT (1, 2)", limit),
        (", SHEET_ROW_LIMIT 2147483648", limit),
        (", SHEET_ROW_LIMIT true", limit),
        (", MODE", mode),
        (", MODE 1", mode),
        (", MODE NULL", mode),
        (", MODE ('append', 'replace')", mode),
        (
            ", MODE 'BoGuS'",
            "Binder Error: Invalid MODE 'bogus' (expected 'create', 'append', or 'replace')",
        ),
        (
            ", MODE ''",
            "Binder Error: Invalid MODE '' (expected 'create', 'append', or 'replace')",
        ),
    ] {
        assert_eq!(
            copy_error(&connection, "SELECT 1 AS a", &path, options),
            message,
            "{options}"
        );
    }
    assert!(directory.file_names().is_empty());
}

#[test]
fn of_several_wrong_options_the_first_of_a_fixed_order_is_reported() {
    let connection = connect();
    let directory = TempDir::new("option_order");
    let path = directory.path("out.xlsx");
    let header = "Binder Error: Header option must be a single boolean value";
    let sheet = "Binder Error: Sheet name option must be a single string value";
    let limit = "Binder Error: Sheet row limit option must be a single integer value";
    for (options, message) in [
        (", SHEET 5, MODE 'bogus'", sheet),
        (", MODE 'bogus', SHEET 5", sheet),
        (", SHEET_ROW_LIMIT 'x', MODE 1", limit),
        (", MODE 1, SHEET_ROW_LIMIT 'x', SHEET 5", sheet),
        (
            ", MODE 1, SHEET 5, SHEET_ROW_LIMIT 'x', HEADER 'maybe'",
            header,
        ),
        (", SHEET_ROW_LIMIT 'x', HEADER 'maybe'", header),
    ] {
        assert_eq!(
            copy_error(&connection, "SELECT 1 AS a", &path, options),
            message,
            "{options}"
        );
    }
    // A wrong option is reported before a workbook that cannot be added to.
    assert_eq!(
        copy_error(
            &connection,
            "SELECT 1",
            "s3://example-bucket/remote.xlsx",
            ", MODE 'append', SHEET 5"
        ),
        sheet
    );
    assert!(directory.file_names().is_empty());
}

#[test]
fn options_are_found_whatever_their_case_and_others_are_ignored() {
    let connection = connect();
    let directory = TempDir::new("option_names");
    let path = directory.path("out.xlsx");
    copy(
        &connection,
        "SELECT 1 AS a FROM range(3)",
        &path,
        ", header TRUE, Sheet 'S', sheet_ROW_limit '4', mode 'CREATE', sheet_name 'ignored', all_varchar true",
    );
    assert_eq!(
        query(
            &connection,
            &format!("SELECT count(*), sum(a) FROM read_xlsx('{path}', sheet = 'S')")
        ),
        ["3|3.0"]
    );
    for (mode, sheet, in_workbook, sheets) in [
        ("create", "First", "First", 1),
        ("Append", "Second", "Second", 2),
        ("REPLACE", "second", "Second", 2),
        ("CREATE", "Third", "Third", 1),
    ] {
        copy(
            &connection,
            "SELECT 2 AS a",
            &path,
            &format!(", MODE '{mode}', SHEET '{sheet}'"),
        );
        assert_eq!(
            query(
                &connection,
                &format!("FROM read_xlsx('{path}', sheet = '{in_workbook}')")
            ),
            ["2.0"]
        );
        let workbook = workbook_part(&path, "xl/workbook.xml");
        assert_eq!(workbook.matches("<sheet ").count(), sheets, "{mode}");
    }
}

#[test]
fn adding_to_a_remote_workbook_is_not_implemented() {
    let connection = connect();
    for (path, mode) in [
        ("s3://example-bucket/remote.xlsx", "append"),
        ("https://example.com/remote.xlsx", "replace"),
        ("gcs://bucket/remote.xlsx", "APPEND"),
        ("hf://datasets/someone/remote.xlsx", "Replace"),
    ] {
        assert_eq!(
            copy_error(
                &connection,
                "SELECT 1",
                path,
                &format!(", SHEET 'X', MODE '{mode}'")
            ),
            format!(
                "Not implemented Error: XLSX MODE '{}' is only supported for local files; '{path}' is on a \
                 remote filesystem. Write the workbook to a local path and upload it instead.",
                mode.to_lowercase()
            )
        );
    }
}

#[test]
fn the_sheet_row_limit() {
    let connection = connect();
    let directory = TempDir::new("limits");
    let path = directory.path("out.xlsx");
    let short = |limit: i64| {
        format!("Invalid Input Error: XLSX: Sheet row limit of '{limit}' rows exceeded!")
    };
    for (rows, options, limit) in [
        (5000, ", sheet_row_limit 2049", 2049),
        (2050, ", sheet_row_limit 2049", 2049),
        (1, ", sheet_row_limit 0", 0),
        (3, ", sheet_row_limit 3, HEADER true", 3),
        (3, ", sheet_row_limit 2.4", 2),
    ] {
        assert_eq!(
            copy_error(
                &connection,
                &format!("SELECT * FROM range({rows})"),
                &path,
                options
            ),
            short(limit),
            "{rows} rows{options}"
        );
        assert!(directory.file_names().is_empty(), "{rows} rows{options}");
    }
    for (rows, options) in [
        (2049, ", sheet_row_limit 2049"),
        (1000, ", sheet_row_limit 2049"),
        (3, ", sheet_row_limit 4, HEADER true"),
        (3, ", sheet_row_limit '3'"),
        (5000, ", sheet_row_limit -1"),
    ] {
        copy(
            &connection,
            &format!("SELECT i FROM range({rows}) t(i)"),
            &path,
            options,
        );
        assert_eq!(
            query(
                &connection,
                &format!("SELECT count(*) FROM read_xlsx('{path}', header = false)")
            ),
            [(rows + i64::from(options.contains("HEADER"))).to_string()],
            "{rows} rows{options}"
        );
    }
    copy(
        &connection,
        "SELECT i FROM range(0) t(i)",
        &path,
        ", sheet_row_limit 0",
    );
    assert_eq!(sheet_rows(&path, "sheet1.xml"), "");
}

#[test]
fn a_sheet_holds_as_many_rows_as_a_spreadsheet_shows() {
    let connection = connect();
    let directory = TempDir::new("million");
    let path = directory.path("out.xlsx");
    assert_eq!(
        copy_error(&connection, "SELECT 1 FROM range(1048577)", &path, ""),
        "Invalid Input Error: XLSX: Sheet row limit of '1048576' rows exceeded!\n \
         * XLSX files and compatible applications generally have a limit of '1048576' rows\n \
         * You can export larger sheets at your own risk by setting the 'sheet_row_limit' parameter to a \
         higher value"
    );
    assert!(directory.file_names().is_empty());

    copy(&connection, "SELECT i FROM range(1048576) t(i)", &path, "");
    assert_eq!(
        query(
            &connection,
            &format!(
                "SELECT count(*), min(A1), max(A1), sum(A1) FROM read_xlsx('{path}', header = false)"
            )
        ),
        ["1048576|0.0|1048575.0|549755289600.0"]
    );
}

#[test]
fn numbers_are_written_as_duckdb_prints_them() {
    let connection = connect();
    let directory = TempDir::new("numbers");
    let path = directory.path("out.xlsx");
    copy(
        &connection,
        "SELECT (-128)::TINYINT AS a, 32767::SMALLINT AS b, (-2147483648)::INTEGER AS c,
                9223372036854775807::BIGINT AS d, (-170141183460469231731687303715884105728)::HUGEINT AS e,
                255::UTINYINT AS f, 65535::USMALLINT AS g, 4294967295::UINTEGER AS h,
                18446744073709551615::UBIGINT AS i, 340282366920938463463374607431768211455::UHUGEINT AS j,
                0.1::FLOAT AS k, 1e-320::DOUBLE AS l, 123456789.123456789::DOUBLE AS m, 1e22::DOUBLE AS n,
                (-12.5)::DECIMAL(4, 1) AS o, 1234567890.0123456789::DECIMAL(38, 10) AS p,
                'nan'::DOUBLE AS q, '-inf'::FLOAT AS r",
        &path,
        "",
    );
    assert_eq!(
        sheet_rows(&path, "sheet1.xml"),
        concat!(
            r#"<row r="1"><c r="A1"><v>-128</v></c><c r="B1"><v>32767</v></c><c r="C1"><v>-2147483648</v></c>"#,
            r#"<c r="D1"><v>9223372036854775807</v></c>"#,
            r#"<c r="E1"><v>-170141183460469231731687303715884105728</v></c>"#,
            r#"<c r="F1"><v>255</v></c><c r="G1"><v>65535</v></c><c r="H1"><v>4294967295</v></c>"#,
            r#"<c r="I1"><v>18446744073709551615</v></c>"#,
            r#"<c r="J1"><v>340282366920938463463374607431768211455</v></c>"#,
            r#"<c r="K1"><v>0.1</v></c><c r="L1"><v>1e-320</v></c><c r="M1"><v>123456789.12345679</v></c>"#,
            r#"<c r="N1"><v>1e+22</v></c><c r="O1"><v>-12.5</v></c><c r="P1"><v>1234567890.0123456789</v></c>"#,
            r#"<c r="Q1"><v>nan</v></c><c r="R1"><v>-inf</v></c></row>"#
        )
    );
    assert_eq!(
        query(
            &connection,
            &format!("SELECT A1, D1, K1, O1, Q1, R1 FROM read_xlsx('{path}')")
        ),
        ["-128.0|9.223372036854776e+18|0.1|-12.5|nan|-inf"]
    );
}

#[test]
fn booleans_dates_and_times_come_back_as_what_they_were() {
    let connection = connect();
    let directory = TempDir::new("temporal");
    let path = directory.path("out.xlsx");
    execute(
        &connection,
        "CREATE TABLE moments AS SELECT
            i % 3 = 0 AS flag,
            DATE '1899-12-30' + (i * 997)::INTEGER AS day,
            TIME '00:00:00' + INTERVAL (i * 1237) SECOND AS tod,
            TIMESTAMP '1960-01-01 00:00:00' + INTERVAL (i * 100003) MINUTE + INTERVAL (i * 250) MILLISECOND AS ts
         FROM range(200) t(i)",
    );
    copy(
        &connection,
        "SELECT flag, day, tod, ts, ts::TIMESTAMP_S AS ts_s, ts::TIMESTAMP_MS AS ts_ms, ts::TIMESTAMP_NS AS ts_ns,
                ts::TIMESTAMPTZ AS ts_tz, tod::TIMETZ AS tod_tz FROM moments",
        &path,
        ", HEADER true",
    );
    assert_eq!(
        describe(&connection, &format!("FROM read_xlsx('{path}')")),
        [
            "flag BOOLEAN",
            "day DATE",
            "tod TIME",
            "ts TIMESTAMP",
            "ts_s TIMESTAMP",
            "ts_ms TIMESTAMP",
            "ts_ns TIMESTAMP",
            "ts_tz TIMESTAMP",
            "tod_tz TIME"
        ]
    );
    // A serial number holds a timestamp to well under a millisecond.
    let to_the_millisecond = |column: &str| format!("{column}::TIMESTAMP_MS::TIMESTAMP");
    let read = query(
        &connection,
        &format!(
            "SELECT flag, day, tod, {}, ts_s, {}, {}, {}, tod_tz FROM read_xlsx('{path}')",
            to_the_millisecond("ts"),
            to_the_millisecond("ts_ms"),
            to_the_millisecond("ts_ns"),
            to_the_millisecond("ts_tz"),
        ),
    );
    let ts = to_the_millisecond("ts");
    let written = query(
        &connection,
        &format!(
            "SELECT flag, day, tod, {ts}, ts::TIMESTAMP_S::TIMESTAMP, {ts}, {ts}, {ts}, tod FROM moments"
        ),
    );
    assert_eq!(written.len(), 200);
    assert_eq!(read, written);

    let cells = sheet_rows(&path, "sheet1.xml");
    assert!(cells.contains(concat!(
        r#"<row r="3"><c r="A3" t="b" s="5"><v>0</v></c><c r="B3" s="1"><v>997.0</v></c>"#,
        r#"<c r="C3" s="3"><v>0.01431712962962963</v></c><c r="D3" s="4"><v>21985.446530671295</v></c>"#,
        r#"<c r="E3" s="2"><v>21985.446527777778</v></c><c r="F3" s="4"><v>21985.446530671295</v></c>"#,
        r#"<c r="G3" s="4"><v>21985.446530671295</v></c><c r="H3" s="4"><v>21985.446530671295</v></c>"#,
        r#"<c r="I3" s="3"><v>0.01431712962962963</v></c></row>"#
    )));
}

#[test]
fn timestamps_with_a_time_zone_are_written_as_utc() {
    let connection = connect();
    let directory = TempDir::new("zones");
    let path = directory.path("out.xlsx");
    copy(
        &connection,
        "SELECT TIMESTAMPTZ '2024-06-01 12:00:00+02' AS moment, TIMETZ '12:34:56+02:00' AS tod",
        &path,
        "",
    );
    assert_eq!(
        query(&connection, &format!("FROM read_xlsx('{path}')")),
        ["2024-06-01 10:00:00|12:34:56"]
    );
}

#[test]
fn values_without_a_cell_type_of_their_own_are_text() {
    let connection = connect();
    let directory = TempDir::new("text_kinds");
    let path = directory.path("out.xlsx");
    execute(
        &connection,
        "CREATE TYPE mood AS ENUM ('sad', 'ok', 'happy')",
    );
    execute(
        &connection,
        "CREATE TABLE things AS SELECT
            'text' AS s, '\\xAA\\x00<'::BLOB AS blob, '0110'::BIT AS bits,
            'a0eebc99-9c0b-4ef8-bb6d-6bb9bd380a11'::UUID AS id, INTERVAL '1 year 2 days 3.5 seconds' AS span,
            'happy'::mood AS m, [1, NULL, 3] AS list, {'a': 1, 'b': 'x<y'} AS record, MAP {'k': 'v&w'} AS pairs,
            [[1], [2, 3]] AS nested, [1, 2, 3]::INTEGER[3] AS fixed, union_value(n := 5) AS choice,
            TIME_NS '01:02:03.123456789' AS tod_ns, 12345678901234567890123::BIGNUM AS big",
    );
    copy(&connection, "FROM things", &path, ", HEADER true");
    assert_eq!(
        query(
            &connection,
            &format!("FROM read_xlsx('{path}', all_varchar = true)")
        ),
        query(
            &connection,
            "SELECT CAST(COLUMNS(*) AS VARCHAR) FROM things"
        )
    );
    let types = describe(&connection, &format!("FROM read_xlsx('{path}')"));
    assert!(
        types.iter().all(|column| column.ends_with(" VARCHAR")),
        "{types:?}"
    );
    let cells = sheet_rows(&path, "sheet1.xml");
    assert!(cells.contains(r#"<c r="B2" t="inlineStr"><is><t>\xAA\x00&lt;</t></is></c>"#));
    assert!(cells.contains(
        r#"<c r="H2" t="inlineStr"><is><t>{&apos;a&apos;: 1, &apos;b&apos;: x&lt;y}</t></is></c>"#
    ));
}

#[test]
fn every_type_there_is_can_be_written() {
    let connection = connect();
    let directory = TempDir::new("all_types");
    let path = directory.path("out.xlsx");
    let columns = query(
        &connection,
        "SELECT column_name || '|' || column_type FROM (DESCRIBE FROM test_all_types())",
    );
    assert!(columns.len() > 50);
    copy(&connection, "FROM test_all_types()", &path, ", HEADER true");
    let read = describe(&connection, &format!("FROM read_xlsx('{path}')"));
    assert_eq!(read.len(), columns.len());

    let mut compared = 0;
    for column in &columns {
        let (name, type_name) = column.split_once('|').unwrap();
        let is_typed = ["BOOLEAN", "DATE", "TIME", "TIMESTAMP"]
            .iter()
            .any(|typed| type_name.starts_with(typed))
            && !type_name.ends_with(']')
            && type_name != "TIME_NS"
            && type_name != "TIMESTAMP_NS WITH TIME ZONE";
        if is_typed {
            continue;
        }
        copy(
            &connection,
            &format!("SELECT \"{name}\" FROM test_all_types()"),
            &path,
            "",
        );
        let read = query(
            &connection,
            &format!("FROM read_xlsx('{path}', all_varchar = true, header = false)"),
        );
        // NUL cannot be written, and a sheet ends at the last row that has a cell.
        let mut written = query(
            &connection,
            &format!("SELECT replace(\"{name}\"::VARCHAR, chr(0), '') FROM test_all_types()"),
        );
        while written.last().is_some_and(|row| row == "NULL") {
            written.pop();
        }
        assert_eq!(read, written, "{column}");
        compared += 1;
    }
    assert!(compared > 40, "{compared}");
}

#[test]
fn text_is_escaped_and_nul_is_dropped() {
    let connection = connect();
    let directory = TempDir::new("escapes");
    let path = directory.path("out.xlsx");
    execute(
        &connection,
        "CREATE TABLE texts AS SELECT * FROM (VALUES
            (1, 'a&b <tag> \"quoted\" it''s'), (2, ']]> &amp; &#65; <![CDATA[x]]>'), (3, 'line1' || chr(10) || 'line2'),
            (4, chr(9) || ' lead and trail  '), (5, ''), (6, 'Ünïcödé 日本語 🦆'), (7, 'nul' || chr(0) || 'byte' || chr(0)),
            (8, NULL), (9, repeat('long ', 5000))) t(id, s)",
    );
    copy(&connection, "SELECT s FROM texts ORDER BY id", &path, "");
    let cells = sheet_rows(&path, "sheet1.xml");
    assert!(cells.starts_with(concat!(
        r#"<row r="1"><c r="A1" t="inlineStr"><is><t>"#,
        r#"a&amp;b &lt;tag&gt; &quot;quoted&quot; it&apos;s</t></is></c></row>"#,
        r#"<row r="2"><c r="A2" t="inlineStr"><is><t>"#,
        r#"]]&gt; &amp;amp; &amp;#65; &lt;![CDATA[x]]&gt;</t></is></c></row>"#,
        "<row r=\"3\"><c r=\"A3\" t=\"inlineStr\"><is><t>line1\nline2</t></is></c></row>",
        "<row r=\"4\"><c r=\"A4\" t=\"inlineStr\"><is><t>\t lead and trail  </t></is></c></row>",
        r#"<row r="5"><c r="A5" t="inlineStr"><is><t></t></is></c></row>"#,
        r#"<row r="6"><c r="A6" t="inlineStr"><is><t>Ünïcödé 日本語 🦆</t></is></c></row>"#,
        r#"<row r="7"><c r="A7" t="inlineStr"><is><t>nulbyte</t></is></c></row>"#,
        r#"<row r="8"></row><row r="9"><c r="A9" t="inlineStr"><is><t>long long "#
    )));
    assert_eq!(
        query(
            &connection,
            &format!("SELECT A1 FROM read_xlsx('{path}', header = false, stop_at_empty = false)")
        ),
        query(
            &connection,
            "SELECT replace(s, chr(0), '') FROM texts ORDER BY id"
        )
    );
}

#[test]
fn empty_cells_are_left_out() {
    let connection = connect();
    let directory = TempDir::new("nulls");
    let path = directory.path("out.xlsx");
    copy(
        &connection,
        "SELECT CASE WHEN i % 2 = 0 THEN i END AS a, NULL AS b, CASE WHEN i % 3 = 0 THEN 's' || i END AS c,
                CASE WHEN i = 1 THEN DATE '2024-01-01' END AS d, NULL::BOOLEAN AS e,
                CASE WHEN i = 4 THEN true END AS f
         FROM range(5) t(i)",
        &path,
        "",
    );
    assert_eq!(
        sheet_rows(&path, "sheet1.xml"),
        concat!(
            r#"<row r="1"><c r="A1"><v>0</v></c><c r="C1" t="inlineStr"><is><t>s0</t></is></c></row>"#,
            r#"<row r="2"><c r="D2" s="1"><v>45292.0</v></c></row>"#,
            r#"<row r="3"><c r="A3"><v>2</v></c></row>"#,
            r#"<row r="4"><c r="C4" t="inlineStr"><is><t>s3</t></is></c></row>"#,
            r#"<row r="5"><c r="A5"><v>4</v></c><c r="F5" t="b" s="5"><v>1</v></c></row>"#
        )
    );
    copy(
        &connection,
        "SELECT NULL::INTEGER AS a, NULL::VARCHAR AS b FROM range(3)",
        &path,
        "",
    );
    assert_eq!(
        sheet_rows(&path, "sheet1.xml"),
        r#"<row r="1"></row><row r="2"></row><row r="3"></row>"#
    );
}

#[test]
fn a_query_without_rows_gives_a_sheet_without_rows() {
    let connection = connect();
    let directory = TempDir::new("empty");
    let path = directory.path("out.xlsx");
    let rows = "SELECT i AS n, 'x' AS s FROM range(10) t(i) WHERE i > 100";
    copy(&connection, rows, &path, "");
    assert_eq!(sheet_rows(&path, "sheet1.xml"), "");
    assert_eq!(workbook_parts(&path).len(), 9);
    copy(&connection, rows, &path, ", HEADER true");
    assert_eq!(
        sheet_rows(&path, "sheet1.xml"),
        concat!(
            r#"<row r="1"><c r="A1" t="inlineStr"><is><t>n</t></is></c>"#,
            r#"<c r="B1" t="inlineStr"><is><t>s</t></is></c></row>"#
        )
    );
    assert_eq!(
        describe(&connection, &format!("FROM read_xlsx('{path}')")),
        ["n DOUBLE", "s DOUBLE"]
    );
}

/// A table of several row groups, which DuckDB scans on several threads when the rows may go to more than
/// one file.
const ROW_GROUPS_ROWS: u64 = 500_000;

#[test]
fn rows_keep_their_order_over_many_batches() {
    let connection = connect();
    let directory = TempDir::new("order");
    let path = directory.path("out.xlsx");
    execute(&connection, "SET threads = 4");
    execute(
        &connection,
        &format!(
            "CREATE TABLE big AS SELECT i AS n, 'row ' || i AS label FROM range({ROW_GROUPS_ROWS}) t(i)"
        ),
    );
    let rows_in_order = || {
        query(
            &connection,
            &format!(
                "SELECT count(*), count(*) FILTER (A1 <> position OR B1 <> 'row ' || position) FROM \
                 (SELECT A1, B1, row_number() OVER () - 1 AS position FROM read_xlsx('{path}'))"
            ),
        )
    };
    copy(&connection, "FROM big", &path, "");
    assert_eq!(rows_in_order(), [format!("{ROW_GROUPS_ROWS}|0")]);
    copy(
        &connection,
        "FROM big WHERE n < 70000",
        &path,
        ", BATCH_SIZE 25000",
    );
    assert_eq!(rows_in_order(), ["70000|0"]);
    copy(
        &connection,
        "FROM big WHERE n < 70000 AND n % 2 = 1 ORDER BY n DESC",
        &path,
        "",
    );
    assert_eq!(
        query(
            &connection,
            &format!(
                "SELECT count(*), count(*) FILTER (A1 <> 69999 - 2 * position) FROM \
                 (SELECT A1, row_number() OVER () - 1 AS position FROM read_xlsx('{path}'))"
            )
        ),
        ["35000|0"]
    );
}

#[test]
fn wide_sheets_name_their_columns_with_letters() {
    let connection = connect();
    let directory = TempDir::new("wide");
    let path = directory.path("out.xlsx");
    let columns: Vec<String> = (0..750)
        .map(|index| format!("i + {index} AS c{index}"))
        .collect();
    copy(
        &connection,
        &format!("SELECT {} FROM range(3) t(i)", columns.join(", ")),
        &path,
        ", HEADER true",
    );
    let cells = sheet_rows(&path, "sheet1.xml");
    for cell in [
        r#"<c r="Z2"><v>25</v></c><c r="AA2"><v>26</v></c>"#,
        r#"<c r="ZZ3"><v>702</v></c><c r="AAA3"><v>703</v></c>"#,
        r#"<c r="ABV4"><v>751</v></c></row>"#,
    ] {
        assert!(cells.contains(cell), "{cell}");
    }
    assert_eq!(
        query(
            &connection,
            &format!("SELECT c0, c26, c749 FROM read_xlsx('{path}')")
        ),
        ["0.0|26.0|749.0", "1.0|27.0|750.0", "2.0|28.0|751.0"]
    );
}

#[test]
fn constant_and_repeated_values_fill_every_row() {
    let connection = connect();
    let directory = TempDir::new("layouts");
    let path = directory.path("out.xlsx");
    execute(
        &connection,
        "CREATE TABLE repeated AS SELECT (i % 4)::VARCHAR AS s, DATE '2020-01-01' + (i % 3)::INTEGER AS d,
                i % 2 = 0 AS b, CASE WHEN i % 5 < 4 THEN TIMESTAMP_S '2020-01-01 00:00:00' END AS ts
         FROM range(20000) t(i)",
    );
    copy(
        &connection,
        "SELECT 42 AS answer, 'same' AS s0, DATE '2024-01-01' AS d0, true AS b0, NULL::TIME AS t0, * FROM repeated",
        &path,
        ", HEADER true",
    );
    assert_eq!(
        query(
            &connection,
            &format!(
                "SELECT count(*), sum(answer), count(DISTINCT s0), min(d0), max(d0), bool_and(b0), count(t0),
                        sum(s::INTEGER), count(DISTINCT d), count(*) FILTER (b), count(ts), min(ts), max(ts)
                 FROM read_xlsx('{path}')"
            )
        ),
        ["20000|840000.0|1|2024-01-01|2024-01-01|true|0|30000|3|10000|16000|2020-01-01 00:00:00|2020-01-01 00:00:00"]
    );
}

#[test]
fn threads_that_write_at_the_same_time_each_get_their_rows_into_a_workbook() {
    let connection = connect();
    let directory = TempDir::new("many_files");
    execute(&connection, "SET threads = 4");
    execute(
        &connection,
        &format!(
            "CREATE TABLE grouped AS SELECT i AS n, i % 3 AS part FROM range({ROW_GROUPS_ROWS}) t(i)"
        ),
    );

    // Every thread writes a file of its own, with the rows in the order the thread read them.
    let per_thread = directory.path("per_thread");
    execute(
        &connection,
        &format!("COPY grouped TO '{per_thread}' (FORMAT xlsx, PER_THREAD_OUTPUT true)"),
    );
    let files = std::fs::read_dir(&per_thread).unwrap().count();
    assert!(files > 1, "{files} file");
    let mut rows = 0;
    for file in 0..files {
        let counts = query(
            &connection,
            &format!(
                "SELECT count(*), count(*) FILTER (A1 <= before) FROM (SELECT A1, lag(A1) OVER () AS before \
                 FROM read_xlsx('{per_thread}/data_{file}.xlsx', header = false))"
            ),
        );
        let (count, out_of_order) = counts[0].split_once('|').unwrap();
        assert_eq!(out_of_order, "0", "file {file}");
        rows += count.parse::<u64>().unwrap();
    }
    assert_eq!(rows, ROW_GROUPS_ROWS);

    // The threads share the files of the partitions; the batches of a file are written one at a time.
    let partitioned = directory.path("partitioned");
    execute(
        &connection,
        &format!("COPY grouped TO '{partitioned}' (FORMAT xlsx, PARTITION_BY (part), HEADER true)"),
    );
    let mut rows = 0;
    for part in 0..3 {
        let path = format!("{partitioned}/part={part}/data_0.xlsx");
        let counts = query(
            &connection,
            &format!(
                "SELECT count(*), count(DISTINCT n), count(*) FILTER (n % 3 <> {part}) FROM read_xlsx('{path}')"
            ),
        );
        let count: u64 = counts[0].split('|').next().unwrap().parse().unwrap();
        assert_eq!(counts, [format!("{count}|{count}|0")], "part {part}");
        let cells = sheet_rows(&path, "sheet1.xml");
        assert_eq!(
            cells.matches("<row r=").count() as u64,
            count + 1,
            "part {part}"
        );
        let last_row = format!("<row r=\"{}\">", count + 1);
        assert!(cells.contains(&last_row), "part {part}");
        rows += count;
    }
    assert_eq!(rows, ROW_GROUPS_ROWS);
}

#[test]
fn a_workbook_written_over_a_larger_one_has_nothing_of_it() {
    let connection = connect();
    let directory = TempDir::new("overwrite");
    let path = directory.path("out.xlsx");
    copy(
        &connection,
        "SELECT i, 'padding ' || i FROM range(20000) t(i)",
        &path,
        "",
    );
    let large = std::fs::metadata(&path).unwrap().len();
    for options in ["", ", USE_TMP_FILE false", ", USE_TMP_FILE true"] {
        copy(&connection, "SELECT 1 AS a", &path, options);
        let small = std::fs::metadata(&path).unwrap().len();
        assert!(small < large / 10, "{options}: {small} of {large} bytes");
        assert_eq!(workbook_parts(&path).len(), 9, "{options}");
        assert_eq!(query(&connection, &format!("FROM '{path}'")), ["1.0"]);
        assert_eq!(directory.file_names(), ["out.xlsx"], "{options}");
    }
}

#[test]
fn a_failing_query_leaves_no_workbook() {
    let connection = connect();
    let directory = TempDir::new("failing");
    let path = directory.path("out.xlsx");
    let error = copy_error(
        &connection,
        "SELECT i::UTINYINT AS small FROM range(100000) t(i)",
        &path,
        "",
    );
    assert!(error.starts_with("Conversion Error"), "{error}");
    assert!(directory.file_names().is_empty());

    let error = copy_error(
        &connection,
        "SELECT 1",
        &directory.path("missing/out.xlsx"),
        "",
    );
    assert!(error.starts_with("IO Error: Cannot open file"), "{error}");
}
