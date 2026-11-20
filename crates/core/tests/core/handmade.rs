//! The hand-made workbooks of the SQL tests, in `test/data/xlsx/handmade`: what each of them is made of, and
//! that the files are what this module makes. Running the tests with `EXCEL_WRITE_FIXTURES=1` writes the
//! files anew.

use std::path::Path;

use crate::support::{
    HANDMADE, Workbook, archive_entries, fixture, inline, latin1, number, row, shared, styled,
    typed, utf16,
};

/// A row of cells that hold `serial`, shown as a date, a time, a date and time, and a number.
fn serial_row(number: u32, serial: &str) -> String {
    row(
        number,
        &[
            styled(&format!("A{number}"), serial, 1),
            styled(&format!("B{number}"), serial, 2),
            styled(&format!("C{number}"), serial, 3),
            styled(&format!("D{number}"), serial, 0),
        ],
    )
}

fn serial_sheet(serials: &[&str]) -> String {
    let header = row(
        1,
        &[
            inline("A1", "d"),
            inline("B1", "t"),
            inline("C1", "ts"),
            inline("D1", "n"),
        ],
    );
    let rows: String = (2..)
        .zip(serials)
        .map(|(number, serial)| serial_row(number, serial))
        .collect();
    format!("{header}{rows}")
}

fn xf_without_numfmtid() -> Vec<u8> {
    let styles = r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<cellXfs count="3"><xf/><xf numFmtId="14"/><xf fontId="0" applyFont="1"/></cellXfs></styleSheet>"#;
    let sheet = [
        row(
            1,
            &[
                inline("A1", "plain"),
                inline("B1", "day"),
                inline("C1", "bold"),
            ],
        ),
        row(
            2,
            &[
                styled("A2", 1.5, 0),
                styled("B2", 44197, 1),
                styled("C2", 44197, 2),
            ],
        ),
        row(
            3,
            &[
                styled("A3", 2.5, 0),
                styled("B3", 44198, 1),
                styled("C3", 7, 2),
            ],
        ),
    ]
    .concat();
    Workbook::new()
        .sheet("Sheet1", &sheet)
        .part("xl/styles.xml", styles)
        .bytes()
}

fn no_text() -> Vec<u8> {
    let sheet = [
        row(1, &[styled("A1", "", 1), styled("B1", "", 3)]),
        row(
            2,
            &[
                r#"<c r="A2" s="4"/>"#.to_owned(),
                r#"<c r="B2" t="s"/>"#.to_owned(),
            ],
        ),
        row(3, &[]),
    ]
    .concat();
    Workbook::new()
        .sheet("Sheet1", &sheet)
        .shared_strings(&["unused"])
        .bytes()
}

fn date1904() -> Vec<u8> {
    Workbook::new()
        .sheet("Sheet1", &serial_sheet(&["0", "100.25", "43538.5", "1"]))
        .properties("date1904=\"1\"")
        .bytes()
}

fn serials() -> Vec<u8> {
    let serials = [
        "1",
        "2",
        "59",
        "59.75",
        "60",
        "60.5",
        "61",
        "0.5",
        "45000",
        // 12:34:56.789, a thousandth of a second after midnight and one before the next second.
        "45366.5242683912",
        "45366.000000011576",
        "45366.0000115625",
        // One step of a DOUBLE below noon, and the last millisecond of February 1900.
        "45366.49999999999",
        "59.99999998842593",
    ];
    Workbook::new()
        .sheet("Sheet1", &serial_sheet(&serials))
        .bytes()
}

fn serial_range() -> Vec<u8> {
    let serials = ["44197.5", "-1e9", "1e300", "nan", "inf", "44198.25"];
    Workbook::new()
        .sheet("Sheet1", &serial_sheet(&serials))
        .bytes()
}

fn cast_reasons() -> Vec<u8> {
    let boolean = |reference: &str, value: &str| typed(reference, "b", value);
    let mut rows = vec![
        row(
            1,
            &[inline("A1", "n"), inline("B1", "flag"), inline("C1", "day")],
        ),
        row(
            2,
            &[number("A2", 1), boolean("B2", "1"), styled("C2", 44197, 1)],
        ),
        // Empty strings do not convert either, without a word.
        row(3, &[inline("A3", ""), boolean("B3", "0"), inline("C3", "")]),
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
            &[number("A5", 5), boolean("B5", "1"), inline("C5", "soon")],
        ),
    ];
    rows.extend((6..=2100).map(|index| {
        row(
            index,
            &[
                number(&format!("A{index}"), index),
                boolean(&format!("B{index}"), "1"),
                styled(&format!("C{index}"), 44197 + index, 1),
            ],
        )
    }));
    rows.push(row(2101, &[inline("A2101", "late")]));
    Workbook::new().sheet("Sheet1", &rows.concat()).bytes()
}

fn escapes() -> Vec<u8> {
    let sheet = [
        row(1, &[shared("A1", 0), inline("B1", "ki_x006E_d")]),
        row(2, &[shared("A2", 1), inline("B2", "shared")]),
        row(3, &[shared("A3", 2), inline("B3", "shared")]),
        row(4, &[shared("A4", 3), inline("B4", "shared")]),
        row(5, &[inline("A5", "cr_x000D_lf"), inline("B5", "inline")]),
        row(
            6,
            &[typed("A6", "str", "tab_x0009_bed"), inline("B6", "formula")],
        ),
        row(
            7,
            &[inline("A7", "not_x12_an_escape"), inline("B7", "inline")],
        ),
        row(
            8,
            &[
                inline("A8", "_xD83E__xDD86_ and _xD83E_"),
                inline("B8", "inline"),
            ],
        ),
    ]
    .concat();
    let strings = "<sst><si><t>head_x0020_er</t></si><si><t>line_x000A_break</t></si>\
                   <si><t>under_x005F_score _x005F_x0041_</t></si>\
                   <si><r><t>two_x0020_</t></r><r><t>runs _x00</t></r><r><t>41_</t></r></si></sst>";
    Workbook::new()
        .sheet("Sheet1", &sheet)
        .part("xl/sharedStrings.xml", strings)
        .bytes()
}

fn empty_shared_string() -> Vec<u8> {
    let empty = |reference: &str| format!(r#"<c r="{reference}" t="s"/>"#);
    let valueless = |reference: &str| format!(r#"<c r="{reference}" t="s"><v></v></c>"#);
    let sheet = [
        row(1, &[inline("A1", "a"), inline("B1", "b")]),
        row(2, &[shared("A2", 0), empty("B2")]),
        row(3, &[empty("A3"), valueless("B3")]),
        row(4, &[shared("A4", 0), shared("B4", 1)]),
    ]
    .concat();
    Workbook::new()
        .sheet("Sheet1", &sheet)
        .shared_strings(&["zero", "one"])
        .bytes()
}

fn relationship_prefix() -> Vec<u8> {
    let workbook = r#"<x:workbook xmlns:x="http://purl.oclc.org/ooxml/spreadsheetml/main"
 xmlns:rel="http://purl.oclc.org/ooxml/officeDocument/relationships"><x:sheets>
<x:sheet name="Data" sheetId="1" rel:id="rId1"/><x:sheet name="Other" sheetId="2" rel:id="rId2"/>
</x:sheets></x:workbook>"#;
    Workbook::new()
        .sheet("Data", &row(1, &[number("A1", 1), number("B1", 2)]))
        .sheet("Other", &row(1, &[number("A1", 3)]))
        .part("xl/workbook.xml", workbook)
        .bytes()
}

/// A workbook whose every part is encoded by `encode`; `texts` are a sheet name, a column name and two values.
fn encoded(texts: [&str; 4], encode: impl Fn(&str) -> Vec<u8>) -> Vec<u8> {
    let [sheet_name, column_name, first, second] = texts;
    let sheet = [
        row(1, &[shared("A1", 0), inline("B1", "n")]),
        row(2, &[shared("A2", 1), number("B2", 1)]),
        row(3, &[inline("A3", second), number("B3", 2)]),
    ]
    .concat();
    Workbook::new()
        .sheet(sheet_name, &sheet)
        .shared_strings(&[column_name, first])
        .bytes_with(|_, content| {
            let text = std::str::from_utf8(content).expect("a part in UTF-8");
            let body = match text.strip_prefix("<?xml") {
                Some(rest) => rest.split_once("?>").expect("the end of the declaration").1,
                None => text,
            };
            encode(body)
        })
}

const WIDE_TEXTS: [&str; 4] = [
    "Größe 表",
    "wärt 漢字",
    "żółw \u{1F986}",
    "naïve &amp; café",
];
const LATIN_TEXTS: [&str; 4] = ["Größe", "wärt", "crème brûlée", "naïve &amp; café"];

fn utf16le() -> Vec<u8> {
    encoded(WIDE_TEXTS, |body| {
        let declared =
            format!("<?xml version=\"1.0\" encoding=\"UTF-16\" standalone=\"yes\"?>\r\n{body}");
        utf16(&declared, false, true)
    })
}

fn utf16be() -> Vec<u8> {
    encoded(WIDE_TEXTS, |body| utf16(body, true, false))
}

fn iso_8859_1() -> Vec<u8> {
    encoded(LATIN_TEXTS, |body| {
        latin1(&format!(
            "<?xml version=\"1.0\" encoding=\"ISO-8859-1\"?>{body}"
        ))
    })
}

/// The file name of a fixture and what makes its bytes.
type Fixture = (&'static str, fn() -> Vec<u8>);

const WORKBOOKS: [Fixture; 12] = [
    ("cast_reasons.xlsx", cast_reasons),
    ("xf_without_numfmtid.xlsx", xf_without_numfmtid),
    ("no_text.xlsx", no_text),
    ("date1904.xlsx", date1904),
    ("serials.xlsx", serials),
    ("serial_range.xlsx", serial_range),
    ("escapes.xlsx", escapes),
    ("empty_shared_string.xlsx", empty_shared_string),
    ("relationship_prefix.xlsx", relationship_prefix),
    ("utf16le.xlsx", utf16le),
    ("utf16be.xlsx", utf16be),
    ("iso_8859_1.xlsx", iso_8859_1),
];

fn parts(bytes: &[u8]) -> Vec<(String, Vec<u8>)> {
    let file = crate::support::TempFile::new("handmade.xlsx");
    std::fs::write(file.path(), bytes).expect("a writable temporary directory");
    archive_entries(file.path())
        .into_iter()
        .map(|entry| (entry.name, entry.content))
        .collect()
}

#[test]
fn the_fixtures_are_the_workbooks_made_here() {
    let directory = fixture(HANDMADE);
    let rewrite = std::env::var_os("EXCEL_WRITE_FIXTURES").is_some();
    for (name, make) in WORKBOOKS {
        let path = Path::new(&directory).join(name);
        let made = make();
        if rewrite {
            std::fs::create_dir_all(&directory).expect("a writable fixture directory");
            std::fs::write(&path, &made).expect("a writable fixture");
        }
        let stored = std::fs::read(&path).unwrap_or_else(|error| panic!("{name}: {error}"));
        assert_eq!(parts(&stored), parts(&made), "{name}");
    }
    let mut names: Vec<String> = std::fs::read_dir(&directory)
        .expect("the fixture directory")
        .map(|entry| entry.expect("a directory entry").file_name())
        .map(|name| name.to_str().expect("a UTF-8 file name").to_owned())
        .collect();
    names.sort();
    let mut expected: Vec<&str> = WORKBOOKS.iter().map(|(name, _)| *name).collect();
    expected.sort_unstable();
    assert_eq!(names, expected);
}
