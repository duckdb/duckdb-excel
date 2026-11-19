//! Shared by the integration tests: an in-process DuckDB with the extension registered, and workbooks made of
//! hand-written XML parts.

use std::{
    io::{Read, Write},
    path::{Path, PathBuf},
    sync::atomic::{AtomicUsize, Ordering},
    time::{Duration, Instant},
};

use duckdb_neo::{
    Parameters, Result,
    connection::Connection,
    environment::{Environment, StorageLocation},
};
use excel_zip::{ZipReader, ZipWriter};

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

/// The directory, among the fixtures, of the workbooks that `handmade.rs` makes.
pub const HANDMADE: &str = "handmade";

fn collect_workbooks(directory: &Path, found: &mut Vec<String>) {
    for entry in std::fs::read_dir(directory).expect("a readable directory") {
        let path = entry.expect("a directory entry").path();
        if path.is_dir() {
            if !path.ends_with(HANDMADE) {
                collect_workbooks(&path, found);
            }
        } else if path
            .extension()
            .is_some_and(|extension| extension == "xlsx")
        {
            found.push(path.to_str().expect("a UTF-8 path").to_owned());
        }
    }
}

/// The fixtures that spreadsheet programs and other libraries wrote, sorted.
pub fn produced_fixtures() -> Vec<String> {
    let mut found = Vec::new();
    collect_workbooks(Path::new(&fixture("")), &mut found);
    found.sort();
    found
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

/// A directory in the temporary directory that is removed, with what it holds, when the value is dropped.
pub struct TempDir(PathBuf);

impl TempDir {
    pub fn new(name: &str) -> TempDir {
        static COUNTER: AtomicUsize = AtomicUsize::new(0);
        let unique = COUNTER.fetch_add(1, Ordering::Relaxed);
        let directory_name = format!("excel_core_{}_dir{unique}_{name}", std::process::id());
        let directory = std::env::temp_dir().join(directory_name);
        std::fs::create_dir_all(&directory).expect("a writable temporary directory");
        TempDir(directory)
    }

    /// The path of `name` in the directory.
    pub fn path(&self, name: &str) -> String {
        self.0.join(name).to_str().expect("a UTF-8 path").to_owned()
    }

    /// Copies a fixture of the SQL tests into the directory as `name` and returns its path.
    pub fn copy_fixture(&self, fixture_path: &str, name: &str) -> String {
        let path = self.path(name);
        std::fs::copy(fixture_path, &path).expect("a readable fixture");
        path
    }

    /// The names of the files in the directory, sorted.
    pub fn file_names(&self) -> Vec<String> {
        let mut names: Vec<String> = std::fs::read_dir(&self.0)
            .expect("a readable directory")
            .map(|entry| entry.expect("a directory entry").file_name())
            .map(|name| name.to_str().expect("a UTF-8 file name").to_owned())
            .collect();
        names.sort();
        names
    }

    /// The file names once only `expected` are left, or what is there after ten seconds. DuckDB can release
    /// what a failed statement held on another thread, after the statement has returned.
    pub fn file_names_after_failure(&self, expected: &[&str]) -> Vec<String> {
        let deadline = Instant::now() + Duration::from_secs(10);
        loop {
            let names = self.file_names();
            if names == expected || Instant::now() >= deadline {
                return names;
            }
            std::thread::sleep(Duration::from_millis(5));
        }
    }
}

impl Drop for TempDir {
    fn drop(&mut self) {
        let _ = std::fs::remove_dir_all(&self.0);
    }
}

/// An entry of an archive as it is stored, and its content.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StoredEntry {
    pub name: String,
    pub crc32: u32,
    pub compressed_size: u64,
    pub content: Vec<u8>,
}

/// The entries of the archive at `path` in the order of its directory; every entry must match its checksum.
pub fn archive_entries(path: &str) -> Vec<StoredEntry> {
    let file = std::fs::File::open(path).unwrap_or_else(|error| panic!("{path}: {error}"));
    let mut zip = ZipReader::new(file).unwrap_or_else(|error| panic!("{path}: {error}"));
    let entries = zip.entries().to_vec();
    entries
        .into_iter()
        .enumerate()
        .map(|(index, entry)| {
            let mut content = Vec::new();
            zip.open_entry(index).expect("an entry that opens");
            zip.read_to_end(&mut content).expect("a readable entry");
            zip.close_entry()
                .expect("an entry that matches its checksum");
            StoredEntry {
                name: entry.name,
                crc32: entry.crc32,
                compressed_size: entry.compressed_size,
                content,
            }
        })
        .collect()
}

/// The name and the content, as text, of every part of the workbook at `path`.
pub fn workbook_parts(path: &str) -> Vec<(String, String)> {
    archive_entries(path)
        .into_iter()
        .map(|entry| {
            let content = String::from_utf8(entry.content).expect("a part in UTF-8");
            (entry.name, content)
        })
        .collect()
}

/// The content of the part `name` of the workbook at `path`.
pub fn workbook_part(path: &str, name: &str) -> String {
    workbook_parts(path)
        .into_iter()
        .rfind(|(part, _)| part == name)
        .map(|(_, content)| content)
        .unwrap_or_else(|| panic!("{path} has no part {name}"))
}

/// What a sheet part holds between the start and the end of its `sheetData`.
pub fn sheet_rows(path: &str, sheet_file: &str) -> String {
    let part = workbook_part(path, &format!("xl/worksheets/{sheet_file}"));
    let (_, rest) = part.split_once("<sheetData>\n\t").expect("sheet data");
    let (rows, _) = rest
        .rsplit_once("</sheetData>")
        .expect("the end of sheet data");
    rows.to_owned()
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
    workbook_properties: String,
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

    /// Gives the workbook a `workbookPr` element with `attributes`, e.g. `date1904="1"`.
    pub fn properties(mut self, attributes: &str) -> Workbook {
        self.workbook_properties = format!("<workbookPr {attributes}/>");
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
        self.bytes_with(|_, content| content.to_vec())
    }

    /// The workbook with every part as `encode` returns it for the name and the content of the part.
    pub fn bytes_with(&self, encode: impl Fn(&str, &[u8]) -> Vec<u8>) -> Vec<u8> {
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
        let properties = &self.workbook_properties;
        let workbook = format!(
            r#"<workbook xmlns="{MAIN_NAMESPACE}" xmlns:r="{RELATIONSHIP_TYPES}">{properties}<sheets>{sheets}</sheets></workbook>"#
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
            zip.write_all(&encode(name, content))
                .expect("ZIP entry data");
            zip.end_file().expect("the end of a ZIP entry");
        }
        zip.finish().expect("a finished ZIP archive")
    }

    /// Writes the workbook to a temporary file.
    pub fn write(&self, name: &str) -> TempFile {
        self.write_with(name, |_, content| content.to_vec())
    }

    /// Writes the workbook of [`Self::bytes_with`] to a temporary file.
    pub fn write_with(&self, name: &str, encode: impl Fn(&str, &[u8]) -> Vec<u8>) -> TempFile {
        let file = TempFile::new(name);
        std::fs::write(file.path(), self.bytes_with(encode))
            .expect("a writable temporary directory");
        file
    }
}

/// `text` in UTF-16, with a byte order mark when `with_mark`.
pub fn utf16(text: &str, is_big_endian: bool, with_mark: bool) -> Vec<u8> {
    let mark = if with_mark { "\u{feff}" } else { "" };
    mark.encode_utf16()
        .chain(text.encode_utf16())
        .flat_map(|unit| {
            if is_big_endian {
                unit.to_be_bytes()
            } else {
                unit.to_le_bytes()
            }
        })
        .collect()
}

/// `text`, which must have no character above U+00FF, in ISO-8859-1.
pub fn latin1(text: &str) -> Vec<u8> {
    text.chars()
        .map(|character| u8::try_from(u32::from(character)).expect("a Latin-1 character"))
        .collect()
}

/// The letters of column `index`, counted from 0: `A`, `B`, ... `Z`, `AA`, ...
pub fn column_letters(index: usize) -> String {
    let mut letters = Vec::new();
    let mut rest = index + 1;
    while rest > 0 {
        letters.push(char::from(b'A' + ((rest - 1) % 26) as u8));
        rest = (rest - 1) / 26;
    }
    letters.iter().rev().collect()
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
