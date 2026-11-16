//! The DuckDB file adapters and the archive on top of them.

mod support;

use std::io::{Read, Seek, SeekFrom, Write};

use duckdb_neo::{error::DuckDBError, file::FileSystem};
use excel_core::{
    file::{FileReader, FileWriter, first_match},
    xlsx::archive::Archive,
};
use excel_zip::{ZipReader, ZipWriter};
use support::{TempFile, Workbook, connect, fixture, inline, number, row};

#[test]
fn a_reader_reads_and_seeks() {
    let connection = connect();
    let file_system = FileSystem::new(&connection).unwrap();
    let file = TempFile::new("reader.bin");
    let content: Vec<u8> = (0..100_000u32).flat_map(u32::to_le_bytes).collect();
    std::fs::write(file.path(), &content).unwrap();

    let mut reader = FileReader::open(&file_system, file.path()).unwrap();
    let mut all = Vec::new();
    reader.read_to_end(&mut all).unwrap();
    assert_eq!(all, content);
    assert_eq!(reader.read(&mut [0; 16]).unwrap(), 0);

    assert_eq!(reader.seek(SeekFrom::End(0)).unwrap(), content.len() as u64);
    assert_eq!(
        reader.seek(SeekFrom::End(-8)).unwrap(),
        content.len() as u64 - 8
    );
    let mut tail = [0; 8];
    reader.read_exact(&mut tail).unwrap();
    assert_eq!(tail, content[content.len() - 8..]);

    assert_eq!(reader.seek(SeekFrom::Start(40)).unwrap(), 40);
    let mut word = [0; 4];
    reader.read_exact(&mut word).unwrap();
    assert_eq!(u32::from_le_bytes(word), 10);
    assert_eq!(reader.seek(SeekFrom::Current(-4)).unwrap(), 40);
    assert_eq!(reader.stream_position().unwrap(), 40);
    assert_eq!(reader.seek(SeekFrom::Current(8)).unwrap(), 48);
    reader.read_exact(&mut word).unwrap();
    assert_eq!(u32::from_le_bytes(word), 12);

    assert!(reader.seek(SeekFrom::Current(-1000)).is_err());
    assert!(reader.seek(SeekFrom::End(i64::MIN)).is_err());
}

#[test]
fn a_missing_file_is_reported_as_duckdb_reports_it() {
    let connection = connect();
    let file_system = FileSystem::new(&connection).unwrap();
    let error = FileReader::open(&file_system, "/nonexistent/dir/book.xlsx").unwrap_err();
    assert_eq!(error.code, DuckDBError::IoGeneral);
    assert_eq!(
        error.message,
        "Cannot open file \"/nonexistent/dir/book.xlsx\": No such file or directory"
    );
    let error = Archive::open(&file_system, "/nonexistent/dir/book.xlsx").unwrap_err();
    assert_eq!(error.code, DuckDBError::IoGeneral);
    assert!(error.message.starts_with("Cannot open file"));
}

#[test]
fn globs_give_their_first_match() {
    let connection = connect();
    let file_system = FileSystem::new(&connection).unwrap();
    let directory = fixture("");
    let first = first_match(&file_system, &format!("{directory}/t*.xlsx")).unwrap();
    let name = first.rsplit('/').next().unwrap();
    assert!(name.starts_with('t') && name.ends_with(".xlsx"), "{first}");
    let only = first_match(&file_system, &format!("{directory}/goo*.xlsx")).unwrap();
    assert!(only.ends_with("/google_sheets.xlsx"), "{only}");
    let exact = fixture("basic.xlsx");
    assert_eq!(first_match(&file_system, &exact).unwrap(), exact);

    let pattern = format!("{directory}/nothing*.xlsx");
    let error = first_match(&file_system, &pattern).unwrap_err();
    assert_eq!(error.code, DuckDBError::IoGeneral);
    assert_eq!(
        error.message,
        format!("Cannot open file \"{pattern}\": No such file or directory")
    );
}

#[test]
fn a_writer_creates_or_empties_the_file_and_writes_in_order() {
    let connection = connect();
    let file_system = FileSystem::new(&connection).unwrap();
    let file = TempFile::new("writer.bin");
    let mut writer = FileWriter::create(&file_system, file.path()).unwrap();
    writer.write_all(b"hello, ").unwrap();
    writer.write_all(&[b'x'; 200_000]).unwrap();
    writer.flush().unwrap();
    writer.write_all(b" world").unwrap();
    writer.close().unwrap();

    let written = std::fs::read(file.path()).unwrap();
    assert_eq!(written.len(), 7 + 200_000 + 6);
    assert!(written.starts_with(b"hello, xxx") && written.ends_with(b"xxx world"));

    let mut writer = FileWriter::create(&file_system, file.path()).unwrap();
    writer.write_all(b"HELLO").unwrap();
    writer.close().unwrap();
    assert_eq!(std::fs::read(file.path()).unwrap(), b"HELLO");

    let error = FileWriter::create(&file_system, "/nonexistent/dir/out.xlsx").unwrap_err();
    assert_eq!(error.code, DuckDBError::IoGeneral);
    assert!(
        error.message.starts_with("Cannot open file"),
        "{}",
        error.message
    );
}

#[test]
fn an_archive_goes_through_duckdb_files_both_ways() {
    let connection = connect();
    let file_system = FileSystem::new(&connection).unwrap();
    let file = TempFile::new("archive.zip");
    let payload = "payload ".repeat(50_000);

    let mut zip = ZipWriter::new(FileWriter::create(&file_system, file.path()).unwrap());
    zip.begin_file("dir/").unwrap();
    zip.end_file().unwrap();
    zip.begin_file("dir/part.xml").unwrap();
    zip.write_all(payload.as_bytes()).unwrap();
    zip.end_file().unwrap();
    zip.begin_file("empty.xml").unwrap();
    zip.end_file().unwrap();
    zip.finish().unwrap().close().unwrap();

    let mut zip = ZipReader::new(FileReader::open(&file_system, file.path()).unwrap()).unwrap();
    let names: Vec<&str> = zip
        .entries()
        .iter()
        .map(|entry| entry.name.as_str())
        .collect();
    assert_eq!(names, ["dir/", "dir/part.xml", "empty.xml"]);
    zip.open_entry(1).unwrap();
    let mut read_back = String::new();
    zip.read_to_string(&mut read_back).unwrap();
    zip.close_entry().unwrap();
    assert_eq!(read_back, payload);

    let mut archive = Archive::open(&file_system, file.path()).unwrap();
    assert!(!archive.try_open_entry("missing.xml").unwrap());
    assert!(archive.try_open_entry("dir/part.xml").unwrap());
    assert_eq!(archive.entry_len(), payload.len() as u64);
    let mut start = [0; 8];
    archive.read_exact(&mut start).unwrap();
    assert_eq!(&start, b"payload ");
    archive.close_entry().unwrap();
    assert!(archive.try_open_entry("empty.xml").unwrap());
    assert_eq!(archive.entry_len(), 0);
    archive.close_entry().unwrap();
}

#[test]
fn the_last_of_several_parts_with_one_name_is_read() {
    let connection = connect();
    let file_system = FileSystem::new(&connection).unwrap();
    let workbook = Workbook::new()
        .sheet("S", &row(1, &[inline("A1", "first"), number("B1", 1)]))
        .write("duplicate_parts.xlsx");
    let mut bytes_with_second_sheet = ZipWriter::new(Vec::new());
    let mut source = ZipReader::new(std::io::Cursor::new(
        std::fs::read(workbook.path()).unwrap(),
    ))
    .unwrap();
    for index in 0..source.entries().len() {
        bytes_with_second_sheet
            .copy_entry(&mut source, index)
            .unwrap();
    }
    bytes_with_second_sheet
        .begin_file("xl/worksheets/sheet1.xml")
        .unwrap();
    bytes_with_second_sheet
        .write_all(b"<replacement/>")
        .unwrap();
    bytes_with_second_sheet.end_file().unwrap();
    std::fs::write(workbook.path(), bytes_with_second_sheet.finish().unwrap()).unwrap();

    let mut archive = Archive::open(&file_system, workbook.path()).unwrap();
    assert!(archive.try_open_entry("xl/worksheets/sheet1.xml").unwrap());
    let mut content = String::new();
    archive.read_to_string(&mut content).unwrap();
    assert_eq!(content, "<replacement/>");
}

#[test]
fn something_that_is_not_an_archive_is_an_io_error() {
    let connection = connect();
    let file_system = FileSystem::new(&connection).unwrap();
    let file = TempFile::new("not_a_zip.xlsx");
    for content in [&b""[..], b"just text", &[0u8; 5000]] {
        std::fs::write(file.path(), content).unwrap();
        let error = Archive::open(&file_system, file.path()).unwrap_err();
        assert_eq!(error.code, DuckDBError::IoGeneral);
        assert!(error.message.contains("ZIP"), "{}", error.message);
    }
}
