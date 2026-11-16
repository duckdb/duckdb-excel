//! Calls made in the wrong state, and the limits the writer enforces.

use std::io::{self, ErrorKind, Read, Write};

use excel_zip::ZipWriter;

use crate::support::{kind, open, read_archive, read_entry, write_archive};

#[test]
fn reader_calls_without_an_open_entry() {
    let archive = write_archive(&[("a.txt", b"abc")]);
    let mut reader = open(&archive).unwrap();
    assert!(!reader.is_entry_open());
    assert_eq!(
        (reader.entry_len(), reader.entry_pos(), reader.is_done()),
        (0, 0, true)
    );
    assert_eq!(kind(reader.read(&mut [0; 8])), ErrorKind::InvalidInput);
    assert_eq!(kind(reader.close_entry()), ErrorKind::InvalidInput);
    assert_eq!(kind(reader.open_entry(1)), ErrorKind::InvalidInput);
    assert_eq!(read_entry(&mut reader, 0, 8).unwrap(), b"abc");
}

#[test]
fn reader_allows_one_open_entry() {
    let archive = write_archive(&[("a.txt", b"abc"), ("b.txt", b"defg")]);
    let mut reader = open(&archive).unwrap();
    reader.open_entry(0).unwrap();
    assert_eq!(kind(reader.open_entry(1)), ErrorKind::InvalidInput);
    assert_eq!(kind(reader.open_entry(0)), ErrorKind::InvalidInput);
    assert_eq!(reader.entry_len(), 3);
    let mut content = Vec::new();
    reader.read_to_end(&mut content).unwrap();
    assert_eq!(content, b"abc");
    reader.close_entry().unwrap();
    assert_eq!(kind(reader.close_entry()), ErrorKind::InvalidInput);
    assert_eq!(read_entry(&mut reader, 1, 8).unwrap(), b"defg");
}

#[test]
fn writer_calls_without_an_open_file() {
    let mut writer = ZipWriter::new(Vec::new());
    assert_eq!(kind(writer.write(b"abc")), ErrorKind::InvalidInput);
    assert_eq!(kind(writer.end_file()), ErrorKind::InvalidInput);
    assert_eq!(
        kind(writer.add_directory("no-slash")),
        ErrorKind::InvalidInput
    );
    assert_eq!(kind(writer.add_directory("")), ErrorKind::InvalidInput);
    writer.flush().unwrap();
    assert_eq!(writer.finish().unwrap().len(), 22);
}

#[test]
fn writer_allows_one_open_file() {
    let source = write_archive(&[("source.txt", b"source")]);
    let mut source = open(&source).unwrap();
    let mut writer = ZipWriter::new(Vec::new());
    writer.begin_file("a.txt").unwrap();
    assert_eq!(kind(writer.begin_file("b.txt")), ErrorKind::InvalidInput);
    assert_eq!(kind(writer.add_directory("dir/")), ErrorKind::InvalidInput);
    assert_eq!(
        kind(writer.copy_entry(&mut source, 0)),
        ErrorKind::InvalidInput
    );
    writer.write_all(b"abc").unwrap();
    writer.end_file().unwrap();
    assert_eq!(kind(writer.end_file()), ErrorKind::InvalidInput);
    let archive = writer.finish().unwrap();
    assert_eq!(
        read_archive(&archive).unwrap(),
        [("a.txt".to_string(), b"abc".to_vec())]
    );
}

#[test]
fn writer_cannot_finish_with_an_open_file() {
    let mut writer = ZipWriter::new(Vec::new());
    writer.begin_file("a.txt").unwrap();
    writer.write_all(b"abc").unwrap();
    assert_eq!(kind(writer.finish()), ErrorKind::InvalidInput);
}

#[test]
fn writer_refuses_names_that_do_not_fit() {
    let mut writer = ZipWriter::new(Vec::new());
    let fitting = "n".repeat(65_535);
    let long = "n".repeat(65_536);
    assert_eq!(kind(writer.begin_file(&long)), ErrorKind::InvalidInput);
    assert_eq!(
        kind(writer.add_directory(&format!("{long}/"))),
        ErrorKind::InvalidInput
    );
    writer.begin_file(&fitting).unwrap();
    writer.end_file().unwrap();
    let archive = writer.finish().unwrap();
    assert_eq!(open(&archive).unwrap().entries()[0].name, fitting);
}

#[test]
fn writer_refuses_more_entries_than_the_classic_format_holds() {
    let mut writer = ZipWriter::new(Vec::new());
    for index in 0..65_534 {
        writer.add_directory(&format!("d{index}/")).unwrap();
    }
    assert_eq!(
        kind(writer.add_directory("one-too-many/")),
        ErrorKind::Unsupported
    );
    assert_eq!(
        kind(writer.begin_file("one-too-many.txt")),
        ErrorKind::Unsupported
    );
    let archive = writer.finish().unwrap();
    let reader = open(&archive).unwrap();
    assert_eq!(reader.entries().len(), 65_534);
    assert_eq!(reader.find("d65533/"), Some(65_533));
}

#[test]
#[ignore = "deflates 4 GiB: run with `--release -- --ignored`"]
fn writer_refuses_a_file_beyond_4_gib() {
    let mut writer = ZipWriter::new(io::sink());
    writer.begin_file("huge.bin").unwrap();
    let zeros = vec![0; 1 << 20];
    for _ in 0..4095 {
        writer.write_all(&zeros).unwrap();
    }
    writer.write_all(&zeros[..(1 << 20) - 3]).unwrap();
    assert_eq!(kind(writer.write(&zeros[..2])), ErrorKind::Unsupported);
    writer.write_all(&zeros[..1]).unwrap();
    writer.end_file().unwrap();
    writer.finish().unwrap();
}
