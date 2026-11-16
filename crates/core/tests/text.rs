//! `text` and `excel_text`.

mod support;

use support::{connect, execute, query, query_error};

#[test]
fn text_and_excel_text_format_numbers() {
    let connection = connect();
    for function in ["text", "excel_text"] {
        let rows = query(
            &connection,
            &format!(
                "SELECT {function}(1234567.891, '#,##0.00'), {function}(0.5, '0%'), {function}(42, '000')"
            ),
        );
        assert_eq!(rows, ["1,234,567.89|50%|042"]);
    }
}

#[test]
fn null_arguments_give_null() {
    let connection = connect();
    for function in ["text", "excel_text"] {
        let rows = query(
            &connection,
            &format!("SELECT {function}(NULL, '0'), {function}(1, NULL), {function}(NULL, NULL)"),
        );
        assert_eq!(rows, ["NULL|NULL|NULL"]);
    }
}

#[test]
fn columns_of_numbers_and_formats() {
    let connection = connect();
    execute(
        &connection,
        "CREATE TABLE numbers AS SELECT * FROM (VALUES (1.5, '0.00'), (NULL, '0'), (2.5, NULL), (1234, '#,##0'), (0.25, '0%')) t(n, f)",
    );
    assert_eq!(
        query(&connection, "SELECT text(n, f) FROM numbers"),
        ["1.50", "NULL", "NULL", "1,234", "25%"]
    );
    assert_eq!(
        query(&connection, "SELECT text(n, '0.0') FROM numbers"),
        ["1.5", "NULL", "2.5", "1234.0", "0.3"]
    );
    assert_eq!(
        query(&connection, "SELECT text(2, f) FROM numbers"),
        ["2.00", "2", "NULL", "2", "200%"]
    );
    let rows = query(
        &connection,
        "SELECT count(*), count(text(i, '0.00')), max(text(i, '0')) FROM range(5000) r(i)",
    );
    assert_eq!(rows, ["5000|5000|999"]);
}

#[test]
fn other_numeric_types_are_cast_to_double() {
    let connection = connect();
    let rows = query(
        &connection,
        "SELECT text(7::TINYINT, '00'), text(1.5::DECIMAL(5, 2), '0.000'), text(3::HUGEINT, '0')",
    );
    assert_eq!(rows, ["07|1.500|3"]);
}

#[test]
fn wrong_arguments_are_rejected() {
    let connection = connect();
    for sql in [
        "SELECT text(1)",
        "SELECT text(true, '0')",
        "SELECT excel_text(1, 2)",
        "SELECT text()",
    ] {
        let error = query_error(&connection, sql);
        assert!(
            error.starts_with("Binder Error: No function matches"),
            "{error}"
        );
    }
    let error = query_error(&connection, "SELECT text('abc', '0')");
    assert!(error.starts_with("Conversion Error:"), "{error}");
}

#[test]
fn the_signature_is_positional() {
    let connection = connect();
    let rows = query(
        &connection,
        "SELECT function_name, parameters::VARCHAR, parameter_types::VARCHAR, return_type FROM duckdb_functions() \
         WHERE function_name IN ('text', 'excel_text') ORDER BY function_name",
    );
    assert_eq!(
        rows,
        [
            "excel_text|[col0, col1]|[DOUBLE, VARCHAR]|VARCHAR",
            "text|[col0, col1]|[DOUBLE, VARCHAR]|VARCHAR"
        ]
    );
}

#[test]
fn formats_that_cannot_be_processed_are_errors_the_connection_survives() {
    let connection = connect();
    for sql in [
        "SELECT text(1, '[$')",
        "SELECT text(-48.5, '@')",
        "SELECT text(i, CASE WHEN i = 3 THEN '0E' ELSE '0' END) FROM range(5) r(i)",
    ] {
        let error = query_error(&connection, sql);
        assert!(
            error.contains("Invalid Input Error: Unexpected result for number format"),
            "{sql}: {error}"
        );
    }
    assert_eq!(query(&connection, "SELECT text(1, '0.0')"), ["1.0"]);
}

#[test]
fn formats_that_show_nothing_give_an_empty_string() {
    let connection = connect();
    let rows = query(
        &connection,
        "SELECT text(1234567.897, '#,##0' || repeat('0', 71)) = '', text(1, '') = ''",
    );
    assert_eq!(rows, ["true|true"]);
}

#[test]
fn formats_that_change_from_row_to_row() {
    let connection = connect();
    let rows = query(
        &connection,
        "SELECT string_agg(text(i, CASE i % 3 WHEN 0 THEN '0.0' WHEN 1 THEN '0%' ELSE '0.0' END), ',' ORDER BY i) FROM range(6) r(i)",
    );
    assert_eq!(rows, ["0.0,100%,2.0,3.0,400%,5.0"]);
}
