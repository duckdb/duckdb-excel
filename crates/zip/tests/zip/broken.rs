//! A writer whose output or source fails half-way must refuse everything afterwards.

use std::io::{Cursor, ErrorKind, Write};

use excel_zip::{ZipReader, ZipWriter};

use crate::support::{
    Appended, RawArchive, RawEntry, Watched, kind, noise, open, read_archive, rows,
};

fn source_archive() -> Vec<u8> {
    RawArchive::of(vec![RawEntry::stored("stored.bin", &noise(300_000, 4))]).build()
}

fn every_call_fails<W: Write>(mut writer: ZipWriter<W>) {
    let source = source_archive();
    let mut source = open(&source).unwrap();
    assert_eq!(kind(writer.write(b"more")), ErrorKind::Other);
    assert_eq!(kind(writer.flush()), ErrorKind::Other);
    assert_eq!(kind(writer.end_file()), ErrorKind::Other);
    assert_eq!(kind(writer.begin_file("next.txt")), ErrorKind::Other);
    assert_eq!(kind(writer.copy_entry(&mut source, 0)), ErrorKind::Other);
    assert_eq!(kind(writer.copy_entry(&mut source, 9)), ErrorKind::Other);
    assert_eq!(kind(writer.finish()), ErrorKind::Other);
}

#[test]
fn failed_flush_breaks_the_writer() {
    let mut writer = ZipWriter::new(Appended::failing(0));
    writer.begin_file("a.txt").unwrap();
    writer.write_all(b"abc").unwrap();
    assert_eq!(kind(writer.flush()), ErrorKind::StorageFull);
    every_call_fails(writer);
}

#[test]
fn failed_write_inside_a_file_breaks_the_writer() {
    let content = noise(500_000, 1);
    let mut writer = ZipWriter::new(Appended::failing(2));
    writer.begin_file("noise.bin").unwrap();
    assert_eq!(kind(writer.write_all(&content)), ErrorKind::StorageFull);
    every_call_fails(writer);
}

#[test]
fn failed_write_at_the_end_of_a_file_breaks_the_writer() {
    let mut writer = ZipWriter::new(Appended::failing(0));
    writer.begin_file("noise.bin").unwrap();
    for seed in 0..3 {
        writer.write_all(&noise(30_000, seed)).unwrap();
    }
    assert_eq!(kind(writer.end_file()), ErrorKind::StorageFull);
    every_call_fails(writer);
}

#[test]
fn failed_write_during_a_copy_breaks_the_writer() {
    let source = source_archive();
    let mut source = open(&source).unwrap();
    let mut writer = ZipWriter::new(Appended::failing(1));
    assert_eq!(
        kind(writer.copy_entry(&mut source, 0)),
        ErrorKind::StorageFull
    );
    every_call_fails(writer);
}

#[test]
fn source_that_fails_during_a_copy_breaks_the_writer() {
    let archive = source_archive();
    let mut source = ZipReader::new(Watched::failing(Cursor::new(&archive), 3)).unwrap();
    let mut writer = ZipWriter::new(Vec::new());
    writer.begin_file("before.xml").unwrap();
    writer.write_all(b"<before/>").unwrap();
    writer.end_file().unwrap();
    assert_eq!(kind(writer.copy_entry(&mut source, 0)), ErrorKind::TimedOut);
    every_call_fails(writer);
}

#[test]
fn source_that_fails_before_a_copy_starts_leaves_the_writer_usable() {
    let archive = source_archive();
    let mut source = ZipReader::new(Watched::failing(Cursor::new(&archive), 1)).unwrap();
    let mut writer = ZipWriter::new(Vec::new());
    assert_eq!(kind(writer.copy_entry(&mut source, 0)), ErrorKind::TimedOut);
    writer.copy_entry(&mut source, 0).unwrap();
    let copy = writer.finish().unwrap();
    assert_eq!(
        read_archive(&copy).unwrap(),
        read_archive(&archive).unwrap()
    );
}

#[test]
fn archive_is_complete_or_an_error_was_reported() {
    let sheet = rows(400_000);
    let image = noise(200_000, 7);
    let source = source_archive();
    let run = |output: Appended| -> std::io::Result<Appended> {
        let mut source = open(&source).unwrap();
        let mut writer = ZipWriter::new(output);
        writer.begin_file("sheet.xml")?;
        writer.write_all(&sheet)?;
        writer.flush()?;
        writer.end_file()?;
        writer.copy_entry(&mut source, 0)?;
        writer.begin_file("image.bin")?;
        writer.write_all(&image)?;
        writer.end_file()?;
        writer.finish()
    };
    let complete = run(Appended::default()).unwrap();
    let expected = read_archive(&complete.bytes).unwrap();
    assert_eq!(expected.len(), 3);
    for failing_write in 0..complete.writes.len() {
        assert_eq!(
            kind(run(Appended::failing(failing_write))),
            ErrorKind::StorageFull,
            "write {failing_write} failed"
        );
    }
    let untouched = run(Appended::failing(complete.writes.len())).unwrap();
    assert_eq!(untouched.bytes, complete.bytes);
}
