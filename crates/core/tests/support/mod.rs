//! Shared by the integration tests: an in-process DuckDB with the extension registered, and workbooks made of
//! hand-written XML parts.

#![allow(dead_code)]

use std::{
    io::Write,
    path::{Path, PathBuf},
    sync::atomic::{AtomicUsize, Ordering},
};

use duckdb_neo::{
    Parameters, Result,
    connection::Connection,
    environment::{Environment, StorageLocation},
};
use excel_zip::ZipWriter;

fn try_connect() -> Result<Connection> {
    let connection = Environment::new()?
        .open(StorageLocation::InMemory)?
        .connect()?;
    excel_core::register_functions!(&connection)?;
    Ok(connection)
}

pub fn connect() -> Connection {
    try_connect().expect("an in-memory DuckDB with the extension registered")
}

/// The rows of `sql` as text, columns joined with `|` and NULL shown as `NULL`.
pub fn try_query(connection: &Connection, sql: &str) -> Result<Vec<String>> {
    try_query_text(
        connection,
        &format!("SELECT CAST(COLUMNS(*) AS VARCHAR) FROM ({sql})"),
    )
}

/// The rows of the statement `sql`, whose columns must all be VARCHAR.
pub fn try_query_text(connection: &Connection, sql: &str) -> Result<Vec<String>> {
    let mut result = connection.query(sql, Parameters::None)?;
    let mut rows = Vec::new();
    while let Some(chunk) = result.next_chunk()? {
        for row in 0..chunk.row_count()? {
            let cells = (0..chunk.col_count())
                .map(|column| {
                    let vector = chunk.get_vector_at::<String>(column)?;
                    Ok(vector.reader()?.get(row)?.unwrap_or("NULL").to_owned())
                })
                .collect::<Result<Vec<_>>>()?;
            rows.push(cells.join("|"));
        }
    }
    Ok(rows)
}

pub fn query(connection: &Connection, sql: &str) -> Vec<String> {
    try_query(connection, sql).unwrap_or_else(|error| panic!("{sql}\nfailed: {error}"))
}

/// The error of `sql`, which must fail: its class and message as DuckDB reports them, without the quoted query.
pub fn query_error(connection: &Connection, sql: &str) -> String {
    match try_query(connection, sql) {
        Ok(rows) => panic!("{sql}\nshould fail but returned {rows:?}"),
        Err(error) => without_location(&error.message),
    }
}

/// The error of the statement `sql`, which must fail.
pub fn execute_error(connection: &Connection, sql: &str) -> String {
    match connection.execute(sql, Parameters::None) {
        Ok(_) => panic!("{sql}\nshould fail"),
        Err(error) => without_location(&error.message),
    }
}

fn without_location(message: &str) -> String {
    let message = message.split("\n\nLINE ").next().unwrap_or(message);
    message.trim_end().to_owned()
}

/// The columns of `sql` as `name TYPE`.
pub fn describe(connection: &Connection, sql: &str) -> Vec<String> {
    query(
        connection,
        &format!("SELECT column_name || ' ' || column_type FROM (DESCRIBE {sql})"),
    )
}

pub fn execute(connection: &Connection, sql: &str) {
    connection
        .execute(sql, Parameters::None)
        .unwrap_or_else(|error| panic!("{sql}\nfailed: {error}"));
}

/// A fixture of the SQL tests, e.g. `basic.xlsx` or `gdal/test.xlsx`.
pub fn fixture(name: &str) -> String {
    Path::new(env!("CARGO_MANIFEST_DIR"))
        .join("../../test/data/xlsx")
        .join(name)
        .to_str()
        .expect("a UTF-8 path")
        .to_owned()
}

/// A file in the temporary directory that is removed when the value is dropped.
pub struct TempFile(PathBuf);

impl TempFile {
    pub fn new(name: &str) -> TempFile {
        static COUNTER: AtomicUsize = AtomicUsize::new(0);
        let unique = COUNTER.fetch_add(1, Ordering::Relaxed);
        let file_name = format!("excel_core_{}_{unique}_{name}", std::process::id());
        TempFile(std::env::temp_dir().join(file_name))
    }

    pub fn path(&self) -> &str {
        self.0.to_str().expect("a UTF-8 path")
    }
}

impl Drop for TempFile {
    fn drop(&mut self) {
        let _ = std::fs::remove_file(&self.0);
    }
}

pub const CONTENT_TYPES: &str = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
</Types>"#;

const MAIN_NAMESPACE: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const RELATIONSHIP_TYPES: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/// Styles 0 to 4 are a number, a date, a time, a timestamp and a number with two decimals.
pub const STYLES: &str = r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd hh:mm:ss"/></numFmts>
<cellXfs count="5"><xf numFmtId="0"/><xf numFmtId="14"/><xf numFmtId="21"/><xf numFmtId="164"/><xf numFmtId="2"/></cellXfs>
</styleSheet>"#;

/// A workbook put together from XML parts.
#[derive(Default)]
pub struct Workbook {
    sheets: Vec<(String, String)>,
    shared_strings: Option<String>,
    parts: Vec<(String, Vec<u8>)>,
    without: Vec<String>,
}

impl Workbook {
    pub fn new() -> Workbook {
        Workbook::default()
    }

    /// Adds a sheet with `sheet_data` as the content of its `sheetData` element.
    pub fn sheet(self, name: &str, sheet_data: &str) -> Workbook {
        let xml = format!(
            r#"<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{MAIN_NAMESPACE}"><sheetData>{sheet_data}</sheetData></worksheet>"#
        );
        self.sheet_xml(name, &xml)
    }

    /// Adds a sheet whose part is exactly `xml`.
    pub fn sheet_xml(mut self, name: &str, xml: &str) -> Workbook {
        self.sheets.push((name.to_owned(), xml.to_owned()));
        self
    }

    pub fn shared_strings(mut self, strings: &[&str]) -> Workbook {
        let items: String = strings
            .iter()
            .map(|string| format!("<si><t>{string}</t></si>"))
            .collect();
        self.shared_strings = Some(format!(
            r#"<sst xmlns="{MAIN_NAMESPACE}" count="{0}" uniqueCount="{0}">{items}</sst>"#,
            strings.len()
        ));
        self
    }

    /// Adds a part, or replaces the one the workbook would write under `name`.
    pub fn part(mut self, name: &str, content: impl AsRef<[u8]>) -> Workbook {
        self.parts
            .push((name.to_owned(), content.as_ref().to_vec()));
        self
    }

    /// Leaves out the part the workbook would write under `name`.
    pub fn without(mut self, name: &str) -> Workbook {
        self.without.push(name.to_owned());
        self
    }

    pub fn bytes(&self) -> Vec<u8> {
        let mut sheets = String::new();
        let mut relations = String::new();
        for (index, (name, _)) in self.sheets.iter().enumerate() {
            let id = index + 1;
            let name = name
                .replace('&', "&amp;")
                .replace('<', "&lt;")
                .replace('"', "&quot;");
            sheets.push_str(&format!(
                r#"<sheet name="{name}" sheetId="{id}" r:id="rId{id}"/>"#
            ));
            relations.push_str(&format!(
                r#"<Relationship Id="rId{id}" Type="{RELATIONSHIP_TYPES}/worksheet" Target="worksheets/sheet{id}.xml"/>"#
            ));
        }
        let workbook = format!(
            r#"<workbook xmlns="{MAIN_NAMESPACE}" xmlns:r="{RELATIONSHIP_TYPES}"><sheets>{sheets}</sheets></workbook>"#
        );
        let workbook_relations = format!(
            r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">{relations}</Relationships>"#
        );

        let mut parts: Vec<(String, Vec<u8>)> = vec![
            ("[Content_Types].xml".to_owned(), CONTENT_TYPES.into()),
            ("xl/workbook.xml".to_owned(), workbook.into_bytes()),
            (
                "xl/_rels/workbook.xml.rels".to_owned(),
                workbook_relations.into_bytes(),
            ),
            ("xl/styles.xml".to_owned(), STYLES.into()),
        ];
        if let Some(shared_strings) = &self.shared_strings {
            parts.push((
                "xl/sharedStrings.xml".to_owned(),
                shared_strings.clone().into_bytes(),
            ));
        }
        for (index, (_, xml)) in self.sheets.iter().enumerate() {
            parts.push((
                format!("xl/worksheets/sheet{}.xml", index + 1),
                xml.clone().into_bytes(),
            ));
        }
        for (name, content) in &self.parts {
            match parts.iter_mut().find(|(known, _)| known == name) {
                Some((_, known_content)) => known_content.clone_from(content),
                None => parts.push((name.clone(), content.clone())),
            }
        }
        parts.retain(|(name, _)| !self.without.contains(name));

        let mut zip = ZipWriter::new(Vec::new());
        for (name, content) in &parts {
            zip.begin_file(name).expect("a new ZIP entry");
            zip.write_all(content).expect("ZIP entry data");
            zip.end_file().expect("the end of a ZIP entry");
        }
        zip.finish().expect("a finished ZIP archive")
    }

    /// Writes the workbook to a temporary file.
    pub fn write(&self, name: &str) -> TempFile {
        let file = TempFile::new(name);
        std::fs::write(file.path(), self.bytes()).expect("a writable temporary directory");
        file
    }
}

/// A `row` element with the cells `cells`.
pub fn row(number: u32, cells: &[String]) -> String {
    format!(r#"<row r="{number}">{}</row>"#, cells.concat())
}

/// A cell with a number, or with whatever `value` is, in the default style.
pub fn number(reference: &str, value: impl std::fmt::Display) -> String {
    format!(r#"<c r="{reference}"><v>{value}</v></c>"#)
}

/// A cell with a number in cell style `style`; see [`STYLES`].
pub fn styled(reference: &str, value: impl std::fmt::Display, style: u32) -> String {
    format!(r#"<c r="{reference}" s="{style}"><v>{value}</v></c>"#)
}

/// A cell with an inline string; `text` is XML.
pub fn inline(reference: &str, text: &str) -> String {
    format!(r#"<c r="{reference}" t="inlineStr"><is><t>{text}</t></is></c>"#)
}

/// A cell that refers to a shared string.
pub fn shared(reference: &str, index: usize) -> String {
    format!(r#"<c r="{reference}" t="s"><v>{index}</v></c>"#)
}

/// A cell of type `cell_type` with `value` as its `v` element.
pub fn typed(reference: &str, cell_type: &str, value: &str) -> String {
    format!(r#"<c r="{reference}" t="{cell_type}"><v>{value}</v></c>"#)
}

/// Rows `first..=last` with the row number in column A and `label <row>` as an inline string in column B.
pub fn numbered_rows(first: u32, last: u32) -> String {
    (first..=last)
        .map(|index| {
            row(
                index,
                &[
                    number(&format!("A{index}"), index),
                    inline(&format!("B{index}"), &format!("label {index}")),
                ],
            )
        })
        .collect()
}
