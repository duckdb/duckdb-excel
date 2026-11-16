//! The workbooks of the SQL tests, which come from Excel, LibreOffice, Google Sheets and GDAL.

mod support;

use std::path::Path;

use support::{connect, describe, fixture, query, try_query};

fn workbooks(directory: &Path, found: &mut Vec<String>) {
    for entry in std::fs::read_dir(directory).unwrap() {
        let path = entry.unwrap().path();
        if path.is_dir() {
            workbooks(&path, found);
        } else if path
            .extension()
            .is_some_and(|extension| extension == "xlsx")
        {
            found.push(path.to_str().unwrap().to_owned());
        }
    }
}

#[test]
fn every_fixture_reads_in_every_shape() {
    let connection = connect();
    let mut paths = Vec::new();
    workbooks(Path::new(&fixture("")), &mut paths);
    assert!(paths.len() >= 25, "{paths:?}");
    for path in paths {
        let text = query(
            &connection,
            &format!(
                "SELECT count(*) FROM read_xlsx('{path}', header = false, all_varchar = true)"
            ),
        );
        let lenient = query(
            &connection,
            &format!("SELECT count(*) FROM read_xlsx('{path}', ignore_errors = true)"),
        );
        let (text, lenient): (usize, usize) =
            (text[0].parse().unwrap(), lenient[0].parse().unwrap());
        assert!(
            text == lenient || text == lenient + 1,
            "{path}: {text} and {lenient}"
        );
        // Without help a sheet either reads the same rows or has a cell that does not fit its column.
        match try_query(
            &connection,
            &format!("SELECT count(*) FROM read_xlsx('{path}')"),
        ) {
            Ok(strict) => assert_eq!(strict, [lenient.to_string()], "{path}"),
            Err(error) => assert!(
                error
                    .message
                    .starts_with("Invalid Input Error: read_xlsx: Failed to parse cell"),
                "{path}: {}",
                error.message
            ),
        }
        let ranged = query(
            &connection,
            &format!(
                "SELECT count(*) FROM read_xlsx('{path}', range = 'A1:H40', ignore_errors = true)"
            ),
        );
        assert!(ranged == ["39"] || ranged == ["40"], "{path}: {ranged:?}");
    }
}

#[test]
fn excel_dates_and_numbers() {
    let connection = connect();
    let sql = format!("FROM read_xlsx('{}')", fixture("gdal/datetime.xlsx"));
    assert_eq!(describe(&connection, &sql), ["time TIMESTAMP"]);
    assert_eq!(
        query(
            &connection,
            &format!("SELECT min(time), max(time), count(*) FROM ({sql})")
        ),
        ["2020-04-07 09:58:00|2020-04-07 10:42:00|10"]
    );
}

#[test]
fn a_sheet_with_thousands_of_rows() {
    let connection = connect();
    let sql = format!("FROM read_xlsx('{}')", fixture("2x3000.xlsx"));
    assert_eq!(describe(&connection, &sql), ["Col1 DOUBLE", "Col2 VARCHAR"]);
    assert_eq!(
        query(
            &connection,
            &format!("SELECT count(*), sum(Col1), count(DISTINCT Col2) FROM ({sql})")
        ),
        ["2999|4498500.0|3"]
    );
}

#[test]
fn the_sheets_of_a_workbook() {
    let connection = connect();
    let path = fixture("two_sheets.xlsx");
    assert_eq!(
        query(&connection, &format!("FROM read_xlsx('{path}')")),
        ["42.0|1337.0"]
    );
    assert_eq!(
        query(
            &connection,
            &format!("FROM read_xlsx('{path}', sheet = 'My Sheet')")
        ),
        query(
            &connection,
            &format!("FROM read_xlsx('{path}', sheet = 'My Sheet', header = true)")
        )
    );
}
