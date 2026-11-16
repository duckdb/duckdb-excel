//! The keyword list that `normalize_names` uses.

use crate::support::{connect, query};
use excel_core::duckdb_compat::reserved_keywords;

#[test]
fn the_embedded_keywords_are_the_reserved_keywords_of_duckdb() {
    let connection = connect();
    let keywords = query(
        &connection,
        "SELECT DISTINCT keyword_name FROM duckdb_keywords() \
         WHERE keyword_category IN ('reserved', 'type_function') ORDER BY keyword_name",
    );
    assert_eq!(reserved_keywords(), keywords);
}

#[test]
fn no_keyword_of_another_category_shadows_a_reserved_one() {
    let connection = connect();
    let shadowed = query(
        &connection,
        "SELECT keyword_name FROM duckdb_keywords() WHERE keyword_category = 'unreserved' AND keyword_name IN \
         (SELECT keyword_name FROM duckdb_keywords() WHERE keyword_category IN ('reserved', 'type_function'))",
    );
    assert!(shadowed.is_empty(), "{shadowed:?}");
}
