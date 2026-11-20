//! Workbooks whose parts are not in UTF-8: UTF-16 and ISO-8859-1 parts read like UTF-8 ones, and no other
//! encoding is read.

use crate::support::{
    Workbook, connect, describe, execute, inline, latin1, number, query, query_error, row, shared,
    styled, utf16,
};

/// `content` with `declaration` in place of the XML declaration it has, if any.
fn declared(content: &[u8], declaration: &str) -> String {
    let text = std::str::from_utf8(content).expect("a part in UTF-8");
    let body = match text.strip_prefix("<?xml") {
        Some(rest) => rest.split_once("?>").expect("the end of the declaration").1,
        None => text,
    };
    format!("{declaration}{body}")
}

fn book(sheet_name: &str, texts: [&str; 4]) -> Workbook {
    let [header, first, second, third] = texts;
    let sheet = [
        row(
            1,
            &[shared("A1", 0), inline("B1", header), inline("C1", "day")],
        ),
        row(
            2,
            &[shared("A2", 1), number("B2", 1.5), styled("C2", 44197, 1)],
        ),
        row(
            3,
            &[
                inline("A3", second),
                number("B3", -2),
                styled("C3", 44198, 1),
            ],
        ),
        row(
            4,
            &[shared("A4", 2), number("B4", "1e3"), styled("C4", 44199, 1)],
        ),
    ]
    .concat();
    Workbook::new()
        .sheet("Other", &row(1, &[number("A1", 0)]))
        .sheet(sheet_name, &sheet)
        .shared_strings(&["name", first, third])
}

fn read(path: &str, sheet_name: &str) -> (Vec<String>, Vec<String>, Vec<String>) {
    let connection = connect();
    let sql = format!("FROM read_xlsx('{path}', sheet = '{sheet_name}')");
    let as_text = format!("FROM read_xlsx('{path}', sheet = '{sheet_name}', all_varchar = true)");
    (
        describe(&connection, &sql),
        query(&connection, &sql),
        query(&connection, &as_text),
    )
}

#[test]
fn workbooks_in_utf16_read_like_those_in_utf8() {
    let sheet_name = "表 \u{1F600} Größe";
    let texts = [
        "wärt \u{1F600}",
        "żółw &amp; 漢字",
        "\u{10FFFF}\u{E000}\u{D7FF} &lt;",
        "\u{1F600}\u{1F601}",
    ];
    let book = book(sheet_name, texts);
    let plain = book.write("utf8.xlsx");
    let expected = read(plain.path(), sheet_name);
    assert_eq!(
        expected.0,
        ["name VARCHAR", "wärt \u{1F600} DOUBLE", "day DATE"]
    );
    assert_eq!(
        expected.1,
        [
            "żółw & 漢字|1.5|2021-01-01",
            "\u{10FFFF}\u{E000}\u{D7FF} <|-2.0|2021-01-02",
            "\u{1F600}\u{1F601}|1000.0|2021-01-03"
        ]
    );
    for is_big_endian in [false, true] {
        for with_mark in [true, false] {
            for declaration in [
                "",
                "<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"yes\"?>\r\n",
                "<?xml version=\"1.0\"?>",
            ] {
                let encoded = book.write_with("utf16.xlsx", |_, content| {
                    utf16(&declared(content, declaration), is_big_endian, with_mark)
                });
                assert_eq!(
                    read(encoded.path(), sheet_name),
                    expected,
                    "{is_big_endian} {with_mark} {declaration:?}"
                );
            }
        }
    }
}

#[test]
fn workbooks_in_latin1_read_like_those_in_utf8() {
    let sheet_name = "Größe × ÿ";
    let texts = [
        "wärt",
        "café &amp; crème",
        "\u{a0}\u{80}\u{ff} &lt;",
        "naïve",
    ];
    let book = book(sheet_name, texts);
    let plain = book.write("utf8.xlsx");
    let expected = read(plain.path(), sheet_name);
    assert_eq!(expected.0, ["name VARCHAR", "wärt DOUBLE", "day DATE"]);
    assert_eq!(
        expected.2,
        [
            "café & crème|1.5|44197",
            "\u{a0}\u{80}\u{ff} <|-2|44198",
            "naïve|1e3|44199"
        ]
    );
    for declaration in [
        "<?xml version=\"1.0\" encoding=\"ISO-8859-1\"?>",
        "<?xml version='1.0' encoding='iso-8859-1' standalone='yes'?>\n",
    ] {
        let encoded = book.write_with("latin1.xlsx", |_, content| {
            latin1(&declared(content, declaration))
        });
        assert_eq!(read(encoded.path(), sheet_name), expected, "{declaration}");
    }
}

#[test]
fn every_part_has_its_own_encoding() {
    let sheet_name = "Größe";
    let book = book(sheet_name, ["wärt", "café", "crème", "naïve"]);
    let plain = book.write("utf8.xlsx");
    let expected = read(plain.path(), sheet_name);
    let parts = [
        "[Content_Types].xml",
        "xl/workbook.xml",
        "xl/_rels/workbook.xml.rels",
        "xl/styles.xml",
        "xl/sharedStrings.xml",
        "xl/worksheets/sheet2.xml",
    ];
    for (index, part) in parts.iter().enumerate() {
        let encoded = book.write_with("mixed.xlsx", |name, content| {
            if name != *part {
                content.to_vec()
            } else if index % 2 == 0 {
                utf16(&declared(content, ""), index % 4 == 0, true)
            } else {
                latin1(&declared(
                    content,
                    "<?xml version=\"1.0\" encoding=\"ISO-8859-1\"?>",
                ))
            }
        });
        assert_eq!(read(encoded.path(), sheet_name), expected, "{part}");
    }
}

fn long_sheet(rows: u32) -> String {
    let header = row(1, &[inline("A1", "id"), inline("B1", "label")]);
    let rows: String = (2..=rows)
        .map(|index| {
            // Characters of one, two and three bytes and pairs of surrogates, at every offset in a buffer.
            let label = format!("{}é漢\u{1F600}{index}", "x".repeat(index as usize % 7));
            row(
                index,
                &[
                    number(&format!("A{index}"), index),
                    inline(&format!("B{index}"), &label),
                ],
            )
        })
        .collect();
    format!("{header}{rows}")
}

#[test]
fn long_sheets_in_other_encodings_are_read_in_pieces() {
    let connection = connect();
    let rows = 30_000;
    let book = Workbook::new().sheet("S", &long_sheet(rows));
    let summary = |path: &str| {
        query(
            &connection,
            &format!(
                "SELECT count(*), sum(id), sum(hash(label)), min(label), max(length(label)) FROM read_xlsx('{path}')"
            ),
        )
    };
    let plain = book.write("long_utf8.xlsx");
    let expected = summary(plain.path());
    assert!(
        expected[0].starts_with("29999|450014999.0|"),
        "{expected:?}"
    );
    for is_big_endian in [false, true] {
        let encoded = book.write_with("long_utf16.xlsx", |_, content| {
            utf16(&declared(content, ""), is_big_endian, true)
        });
        assert_eq!(summary(encoded.path()), expected, "{is_big_endian}");
    }

    let latin = |index: u32| format!("{}é\u{ff}{index}", "x".repeat(index as usize % 7));
    let sheet: String = (1..=rows)
        .map(|index| row(index, &[inline(&format!("A{index}"), &latin(index))]))
        .collect();
    let book = Workbook::new().sheet("S", &sheet);
    let labels = |path: &str| {
        query(
            &connection,
            &format!(
                "SELECT count(*), sum(hash(A1)), max(length(A1)) FROM read_xlsx('{path}', header = false)"
            ),
        )
    };
    let plain = book.write("long_utf8.xlsx");
    let encoded = book.write_with("long_latin1.xlsx", |_, content| {
        latin1(&declared(
            content,
            "<?xml version=\"1.0\" encoding=\"ISO-8859-1\"?>",
        ))
    });
    assert_eq!(labels(encoded.path()), labels(plain.path()));
}

#[test]
fn copy_from_reads_other_encodings() {
    let connection = connect();
    let sheet_name = "Größe";
    let book = book(sheet_name, ["wärt", "café", "crème", "naïve"]);
    let encoded = book.write_with("copy_utf16.xlsx", |_, content| {
        utf16(&declared(content, ""), true, false)
    });
    execute(
        &connection,
        "CREATE TABLE people (name VARCHAR, amount DECIMAL(9, 2), day DATE)",
    );
    execute(
        &connection,
        &format!(
            "COPY people FROM '{}' (FORMAT xlsx, SHEET '{sheet_name}')",
            encoded.path()
        ),
    );
    assert_eq!(
        query(&connection, "FROM people"),
        [
            "café|1.50|2021-01-01",
            "crème|-2.00|2021-01-02",
            "naïve|1000.00|2021-01-03"
        ]
    );
}

fn sheet_error(sheet: Vec<u8>) -> String {
    let connection = connect();
    let file = Workbook::new()
        .sheet("S", "")
        .part("xl/worksheets/sheet1.xml", sheet)
        .write("bad_encoding.xlsx");
    query_error(&connection, &format!("FROM read_xlsx('{}')", file.path()))
}

#[test]
fn text_that_is_not_in_the_encoding_of_its_part_is_an_xml_error() {
    let sheet = |cell: &str| {
        format!(
            "<worksheet><sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>{cell}</t></is></c></row>\
             </sheetData></worksheet>"
        )
    };
    let (before, after) = sheet("a|b")
        .split_once('|')
        .map(|(before, after)| (before.to_owned(), after.to_owned()))
        .unwrap();
    let with_units = |units: &[u16]| {
        let middle: Vec<u8> = units.iter().flat_map(|unit| unit.to_le_bytes()).collect();
        [
            utf16(&before, false, true),
            middle,
            utf16(&after, false, false),
        ]
        .concat()
    };
    assert!(sheet_error(with_units(&[0xD83D])).ends_with("half a surrogate pair in UTF-16 text"));
    assert!(
        sheet_error(with_units(&[0xDE00, 0xD83D]))
            .ends_with("half a surrogate pair in UTF-16 text")
    );
    assert!(sheet_error(with_units(&[0x1])).ends_with("invalid character"));
    // Half a character at the end of the part.
    let mut cut = utf16(&before, false, true);
    cut.push(b'b');
    assert!(sheet_error(cut).ends_with("the document ends in the middle of a character"));
    let not_utf8 = latin1(&sheet("caf\u{e9}"));
    let error = sheet_error(not_utf8);
    assert!(error.starts_with("IO Error: XML parse error"), "{error}");
}

#[test]
fn a_part_must_be_in_the_encoding_it_declares_and_other_encodings_are_not_read() {
    let body =
        "<worksheet><sheetData><row r=\"1\"><c r=\"A1\"><v>1</v></c></row></sheetData></worksheet>";
    let declaration =
        |encoding: &str| format!("<?xml version=\"1.0\" encoding=\"{encoding}\"?>{body}");
    for (encoding, bytes) in [
        ("UTF-8", utf16(&declaration("UTF-8"), false, true)),
        ("ISO-8859-1", utf16(&declaration("ISO-8859-1"), true, false)),
        ("UTF-16LE", utf16(&declaration("UTF-16LE"), true, true)),
        ("UTF-16", declaration("UTF-16").into_bytes()),
    ] {
        let error = sheet_error(bytes);
        assert!(
            error.starts_with("IO Error: XML parse error at byte ")
                && error.ends_with(&format!(
                    "the document is not in the encoding '{encoding}' it declares"
                )),
            "{error}"
        );
    }
    for encoding in [
        "windows-1252",
        "ISO-8859-15",
        "Shift_JIS",
        "UTF-32",
        "latin1",
    ] {
        for bytes in [
            declaration(encoding).into_bytes(),
            utf16(&declaration(encoding), false, true),
        ] {
            assert_eq!(
                sheet_error(bytes),
                format!(
                    "IO Error: XML parse error: unsupported encoding '{encoding}', only UTF-8, UTF-16 and \
                     ISO-8859-1 are supported"
                )
            );
        }
    }
    let connection = connect();
    for encoding in ["UTF-8", "utf-8", "US-ASCII"] {
        let file = Workbook::new()
            .sheet("S", "")
            .part("xl/worksheets/sheet1.xml", declaration(encoding))
            .write("declared.xlsx");
        assert_eq!(
            query(
                &connection,
                &format!("FROM read_xlsx('{}', header = false)", file.path())
            ),
            ["1.0"]
        );
    }
}
