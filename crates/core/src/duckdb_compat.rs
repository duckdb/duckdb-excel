//! Helpers of DuckDB that its C API does not offer, written to give the results DuckDB v2.0.0 gives.

use std::collections::HashMap;

/// The keywords DuckDB v2.0.0 lists as `reserved` or `type_function`, sorted; an identifier that is one of them
/// has to be quoted.
const RESERVED_KEYWORDS: [&str; 111] = [
    "all",
    "analyse",
    "analyze",
    "and",
    "anti",
    "any",
    "array",
    "as",
    "asc",
    "asof",
    "asymmetric",
    "at",
    "authorization",
    "binary",
    "both",
    "by",
    "case",
    "cast",
    "check",
    "collate",
    "collation",
    "column",
    "columns",
    "concurrently",
    "constraint",
    "create",
    "cross",
    "default",
    "deferrable",
    "desc",
    "describe",
    "distinct",
    "do",
    "else",
    "end",
    "except",
    "false",
    "fetch",
    "for",
    "foreign",
    "freeze",
    "from",
    "full",
    "generated",
    "glob",
    "group",
    "having",
    "ilike",
    "in",
    "initially",
    "inner",
    "intersect",
    "into",
    "is",
    "isnull",
    "join",
    "lambda",
    "lateral",
    "leading",
    "left",
    "like",
    "limit",
    "map",
    "natural",
    "not",
    "notnull",
    "null",
    "offset",
    "on",
    "only",
    "or",
    "order",
    "outer",
    "overlaps",
    "pivot",
    "pivot_longer",
    "pivot_wider",
    "placing",
    "positional",
    "primary",
    "qualify",
    "references",
    "returning",
    "right",
    "select",
    "semi",
    "show",
    "similar",
    "some",
    "struct",
    "summarize",
    "symmetric",
    "table",
    "tablesample",
    "then",
    "to",
    "trailing",
    "true",
    "try_cast",
    "tuple",
    "union",
    "unique",
    "unpack",
    "unpivot",
    "using",
    "variadic",
    "verbose",
    "when",
    "where",
    "window",
    "with",
];

/// The reserved and type-function keywords of DuckDB v2.0.0, sorted.
pub fn reserved_keywords() -> &'static [&'static str] {
    &RESERVED_KEYWORDS
}

/// Whether `name`, in any letter case, is a keyword that cannot be used as a column name without quotes.
pub fn is_reserved_keyword(name: &str) -> bool {
    let name = name.to_ascii_lowercase();
    RESERVED_KEYWORDS.binary_search(&name.as_str()).is_ok()
}

/// Gives every name that repeats an earlier one, in any letter case, a suffix `_1`, `_2`, ... that makes it unique.
pub fn deduplicate_columns(names: &mut [String]) {
    let mut counts: HashMap<String, usize> = HashMap::new();
    for name in names {
        let key = name.to_ascii_lowercase();
        let Some(mut count) = counts.get(&key).copied() else {
            counts.insert(key, 1);
            continue;
        };
        let renamed = loop {
            let candidate = format!("{name}_{count}");
            if !counts.contains_key(&candidate.to_ascii_lowercase()) {
                break candidate;
            }
            count += 1;
        };
        counts.insert(key, count);
        counts.insert(renamed.to_ascii_lowercase(), 1);
        *name = renamed;
    }
}

fn similarity_score(left: &[u8], right: &[u8]) -> usize {
    if left.is_empty() {
        return right.len();
    }
    if right.is_empty() {
        return left.len();
    }
    let mut previous: Vec<usize> = (0..=left.len()).collect();
    let mut current = vec![0; left.len() + 1];
    for (row, right_byte) in right.iter().enumerate() {
        current[0] = row + 1;
        for (column, left_byte) in left.iter().enumerate() {
            let replace_cost = if left_byte.eq_ignore_ascii_case(right_byte) {
                0
            } else {
                3
            };
            current[column + 1] = (previous[column] + replace_cost)
                .min(previous[column + 1] + 1)
                .min(current[column] + 1);
        }
        std::mem::swap(&mut previous, &mut current);
    }
    previous[left.len()]
}

fn case_insensitive_less(left: &str, right: &str) -> bool {
    let upper = |text: &str| {
        text.bytes()
            .map(|byte| byte.to_ascii_uppercase())
            .collect::<Vec<_>>()
    };
    upper(left) < upper(right)
}

/// The suggestions DuckDB appends to "not found" errors: `\n<prefix>: "a", "b"` with up to five of `candidates`
/// that resemble `target`, or nothing when there are no candidates.
pub fn candidates_error_message(candidates: &[String], target: &str, prefix: &str) -> String {
    const MAX_SUGGESTIONS: usize = 5;
    // Distances count 3 for a replaced character and 1 for an inserted or deleted one.
    const THRESHOLD: usize = 5;

    let scores: Vec<(&str, usize)> = candidates
        .iter()
        .map(|candidate| {
            let bytes = candidate.as_bytes();
            let compared = bytes.get(..target.len()).unwrap_or(bytes);
            (
                candidate.as_str(),
                similarity_score(compared, target.as_bytes()),
            )
        })
        .collect();
    let max_score = scores
        .iter()
        .map(|(_, score)| *score)
        .fold(THRESHOLD, usize::max);
    let rating = |score: usize| 1.0 - score as f64 / max_score as f64;
    let mut rated: Vec<(&str, f64)> = scores
        .iter()
        .map(|(candidate, score)| (*candidate, rating(*score)))
        .collect();
    rated.sort_by(|left, right| {
        right.1.total_cmp(&left.1).then_with(|| {
            if case_insensitive_less(left.0, right.0) {
                std::cmp::Ordering::Less
            } else if case_insensitive_less(right.0, left.0) {
                std::cmp::Ordering::Greater
            } else {
                std::cmp::Ordering::Equal
            }
        })
    });
    let threshold = rating(THRESHOLD);
    let suggestions: Vec<String> = rated
        .iter()
        .enumerate()
        .take(MAX_SUGGESTIONS)
        .take_while(|(index, (_, rating))| *index == 0 || *rating >= threshold)
        .map(|(_, (candidate, _))| format!("\"{candidate}\""))
        .collect();
    if suggestions.is_empty() {
        String::new()
    } else {
        format!("\n{prefix}: {}", suggestions.join(", "))
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn deduplicated(names: &[&str]) -> Vec<String> {
        let mut names: Vec<String> = names.iter().map(|name| (*name).to_owned()).collect();
        deduplicate_columns(&mut names);
        names
    }

    #[test]
    fn keywords_are_sorted_and_found() {
        assert!(reserved_keywords().is_sorted());
        assert!(is_reserved_keyword("select") && is_reserved_keyword("SELECT"));
        assert!(is_reserved_keyword("pivot_wider") && is_reserved_keyword("with"));
        assert!(
            !is_reserved_keyword("name") && !is_reserved_keyword("") && !is_reserved_keyword("_")
        );
    }

    #[test]
    fn duplicate_columns_get_a_counter() {
        assert_eq!(
            deduplicated(&["a", "b", "a", "a"]),
            ["a", "b", "a_1", "a_2"]
        );
        assert_eq!(deduplicated(&["a", "A", "a"]), ["a", "A_1", "a_2"]);
        assert_eq!(
            deduplicated(&["a", "a_1", "a", "a"]),
            ["a", "a_1", "a_2", "a_3"]
        );
        assert_eq!(deduplicated(&["a", "a", "a_1"]), ["a", "a_1", "a_1_1"]);
        assert_eq!(deduplicated(&["", ""]), ["", "_1"]);
    }

    #[test]
    fn similarity_scores() {
        for (left, right, score) in [
            ("", "abc", 3),
            ("abc", "", 3),
            ("Sheet", "sheet", 0),
            ("ab", "abc", 1),
            ("abc", "abd", 2),
            ("abc", "xyz", 6),
            ("My", "Sh", 4),
        ] {
            assert_eq!(
                similarity_score(left.as_bytes(), right.as_bytes()),
                score,
                "{left} {right}"
            );
        }
    }

    #[test]
    fn suggestions() {
        let sheets = ["Sheet1".to_owned(), "My Sheet".to_owned()];
        assert_eq!(
            candidates_error_message(&sheets, "Sh", "Did you mean"),
            "\nDid you mean: \"Sheet1\", \"My Sheet\""
        );
        assert_eq!(
            candidates_error_message(&sheets[..1], "FooBar", "Did you mean"),
            "\nDid you mean: \"Sheet1\""
        );
        assert_eq!(candidates_error_message(&[], "x", "Did you mean"), "");
    }

    #[test]
    fn suggestions_keep_the_five_closest() {
        let names: Vec<String> = [
            "data7", "data1", "other", "data3", "DATA2", "data5", "data4", "data6",
        ]
        .iter()
        .map(|name| (*name).to_owned())
        .collect();
        assert_eq!(
            candidates_error_message(&names, "data", "Did you mean"),
            "\nDid you mean: \"data1\", \"DATA2\", \"data3\", \"data4\", \"data5\""
        );
    }
}
