//! `SELECT * FROM 'file.xlsx'`.

mod support;

use support::{connect, describe, fixture, query, query_error};

#[test]
fn a_path_that_ends_in_xlsx_is_read() {
    let connection = connect();
    let path = fixture("google_sheets.xlsx");
    let direct = query(&connection, &format!("FROM read_xlsx('{path}')"));
    assert_eq!(query(&connection, &format!("FROM '{path}'")), direct);
    assert_eq!(direct, ["123.0|500,60|ABC", "456.0|300,40|some text"]);
    assert_eq!(
        describe(&connection, &format!("FROM '{path}'")),
        ["ABC DOUBLE", "HELLO VARCHAR", "WORLD VARCHAR"]
    );
    assert_eq!(
        query(&connection, &format!("SELECT count(*) FROM '{path}'")),
        ["2"]
    );
}

#[test]
fn the_extension_matches_in_any_letter_case() {
    let connection = connect();
    for name in ["missing.XLSX", "missing.XlSx", "dir/missing.xlsx"] {
        let error = query_error(&connection, &format!("FROM '{name}'"));
        assert_eq!(
            error,
            format!("IO Error: Cannot open file \"{name}\": No such file or directory")
        );
    }
}

#[test]
fn qualified_names_are_joined_with_dots() {
    let connection = connect();
    for (sql, path) in [
        ("FROM missing.xlsx", "missing.xlsx"),
        ("FROM a.b.xlsx", "a.b.xlsx"),
        ("FROM \"a.b\".xlsx", "a.b.xlsx"),
        ("FROM \"weird name.xlsx\"", "weird name.xlsx"),
        ("FROM main.\"x.xlsx\"", "main.x.xlsx"),
    ] {
        let error = query_error(&connection, sql);
        assert_eq!(
            error,
            format!("IO Error: Cannot open file \"{path}\": No such file or directory"),
            "{sql}"
        );
    }
}

#[test]
fn other_names_are_left_alone() {
    let connection = connect();
    for sql in [
        "FROM xlsx",
        "FROM \"file.xlsx.bak\"",
        "FROM \"file.xls\"",
        "FROM \"xlsx.\"",
    ] {
        let error = query_error(&connection, sql);
        assert!(
            error.starts_with("Catalog Error: Table with name"),
            "{sql}: {error}"
        );
    }
}

#[test]
fn tables_win_over_files() {
    let connection = connect();
    support::execute(
        &connection,
        "CREATE TABLE \"data.xlsx\" AS SELECT 42 AS answer",
    );
    assert_eq!(query(&connection, "FROM \"data.xlsx\""), ["42"]);
}
