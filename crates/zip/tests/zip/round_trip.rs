//! Archives written by `ZipWriter` and read back by `ZipReader`.

use std::io::{Cursor, Read, Write};

use excel_zip::{ZipEntry, ZipReader, ZipWriter};

use crate::support::{
    Appended, FLAG_DATA_DESCRIPTOR, FLAG_UTF8, METHOD_DEFLATED, RawArchive, RawEntry, Watched,
    central_directory, crc32, noise, open, read_archive, read_entry, rows, write_archive,
};

const LARGE_LEN: usize = 24 * 1024 * 1024;

fn round_trip(files: &[(&str, &[u8])]) {
    let archive = write_archive(files);
    let read = read_archive(&archive).unwrap();
    assert_eq!(read.len(), files.len());
    for ((name, content), (read_name, read_content)) in files.iter().zip(&read) {
        assert_eq!(name, read_name);
        assert!(content == read_content, "content of '{name}' differs");
    }
}

#[test]
fn many_small_entries() {
    let contents: Vec<(String, Vec<u8>)> = (0..3000usize)
        .map(|index| {
            let content = if index % 3 == 0 {
                rows(index % 700)
            } else {
                noise(index % 500, index as u64)
            };
            (format!("part/{}/item{index}.xml", index % 17), content)
        })
        .collect();
    let files: Vec<(&str, &[u8])> = contents
        .iter()
        .map(|(name, content)| (name.as_str(), content.as_slice()))
        .collect();
    round_trip(&files);

    let archive = write_archive(&files);
    let reader = open(&archive).unwrap();
    assert_eq!(reader.entries().len(), 3000);
    assert_eq!(reader.find("part/5/item2997.xml"), Some(2997));
    assert_eq!(reader.find("part/5/item2998.xml"), None);
}

#[test]
fn large_compressible_entry() {
    let content = rows(LARGE_LEN);
    let archive = write_archive(&[("xl/worksheets/sheet1.xml", &content)]);
    assert!(archive.len() < content.len() / 4);
    let mut reader = open(&archive).unwrap();
    assert_eq!(reader.entries()[0].uncompressed_size, LARGE_LEN as u64);
    assert!(read_entry(&mut reader, 0, 64 * 1024).unwrap() == content);
}

#[test]
fn large_incompressible_entry() {
    let content = noise(LARGE_LEN, 7);
    let archive = write_archive(&[("media/image.bin", &content)]);
    assert!(archive.len() >= content.len());
    let mut reader = open(&archive).unwrap();
    assert!(read_entry(&mut reader, 0, 100_000).unwrap() == content);
}

#[test]
fn empty_file_and_empty_archive() {
    round_trip(&[("empty.xml", b"")]);
    round_trip(&[("a", b""), ("b", b"x"), ("c", b"")]);

    let archive = write_archive(&[]);
    assert_eq!(archive.len(), 22);
    assert!(open(&archive).unwrap().entries().is_empty());
}

#[test]
fn names_ending_with_a_slash_are_directories() {
    let written = write_archive(&[("xl/", b""), ("xl/workbook.xml", b"<workbook/>")]);
    let built = RawArchive::of(vec![
        RawEntry::stored("xl/", b""),
        RawEntry::deflated("xl/workbook.xml", b"<workbook/>"),
    ])
    .build();
    for archive in [written, built] {
        let reader = open(&archive).unwrap();
        let directories: Vec<bool> = reader
            .entries()
            .iter()
            .map(ZipEntry::is_directory)
            .collect();
        assert_eq!(directories, [true, false]);
        assert_eq!(
            (
                &*reader.entries()[0].name,
                reader.entries()[0].uncompressed_size
            ),
            ("xl/", 0)
        );
        let files = read_archive(&archive).unwrap();
        assert!(files[0].1.is_empty());
        assert_eq!(files[1].1, b"<workbook/>");
    }
}

#[test]
fn non_ascii_names() {
    let files: [(&str, &[u8]); 3] = [
        ("données/", b""),
        ("données/feuille ☃.xml", b"<neige/>"),
        ("plain.xml", b"<plain/>"),
    ];
    round_trip(&files);

    let archive = write_archive(&files);
    let utf8_flags: Vec<bool> = central_directory(&archive)
        .iter()
        .map(|record| record.flags & FLAG_UTF8 != 0)
        .collect();
    assert_eq!(utf8_flags, [true, true, false]);
    assert_eq!(
        open(&archive).unwrap().find("données/feuille ☃.xml"),
        Some(1)
    );
}

#[test]
fn small_and_odd_chunk_sizes() {
    let deflated = rows(200_000);
    let stored = noise(150_001, 3);
    let source = RawArchive::of(vec![
        RawEntry::deflated("deflated.xml", &deflated),
        RawEntry::stored("stored.bin", &stored),
    ])
    .build();
    for chunk_len in [1, 2, 3, 7, 255, 4097, 65_535, 65_536, 65_537, 1_000_000] {
        let mut reader = open(&source).unwrap();
        assert!(
            read_entry(&mut reader, 0, chunk_len).unwrap() == deflated,
            "chunks of {chunk_len}"
        );
        assert!(
            read_entry(&mut reader, 1, chunk_len).unwrap() == stored,
            "chunks of {chunk_len}"
        );
    }
}

#[test]
fn file_that_returns_short_reads() {
    let content = rows(300_000);
    let noisy = noise(70_000, 11);
    let archive = write_archive(&[("sheet.xml", &content), ("noise.bin", &noisy)]);
    for limit in [1, 13, 4096] {
        let mut reader = ZipReader::new(Watched::trickling(Cursor::new(&archive), limit)).unwrap();
        assert!(read_entry(&mut reader, 0, 10_000).unwrap() == content);
        assert!(read_entry(&mut reader, 1, 10_000).unwrap() == noisy);
    }
}

#[test]
fn duplicate_names_resolve_to_the_last() {
    let archive = write_archive(&[
        ("xl/workbook.xml", b"first"),
        ("other.xml", b"other"),
        ("xl/workbook.xml", b"second"),
        ("xl/workbook.xml", b"third"),
        ("tail.xml", b"tail"),
    ]);
    let mut reader = open(&archive).unwrap();
    assert_eq!(reader.entries().len(), 5);
    let index = reader.find("xl/workbook.xml").unwrap();
    assert_eq!(index, 3);
    assert_eq!(read_entry(&mut reader, index, 16).unwrap(), b"third");
    assert_eq!(reader.find("missing.xml"), None);
}

#[test]
fn closing_early_skips_the_checksum() {
    let content = rows(500_000);
    let mut entry = RawEntry::deflated("sheet.xml", &content);
    entry.crc32 ^= 1;
    let archive = RawArchive::of(vec![entry, RawEntry::stored("next.txt", b"next")]).build();
    let mut reader = open(&archive).unwrap();

    reader.open_entry(0).unwrap();
    assert_eq!(reader.entry_len(), 500_000);
    let mut start = [0; 1000];
    reader.read_exact(&mut start).unwrap();
    assert_eq!(start, content[..1000]);
    reader.close_entry().unwrap();
    assert_eq!(reader.entry_len(), 0);

    assert_eq!(read_entry(&mut reader, 1, 100).unwrap(), b"next");
    assert!(read_entry(&mut reader, 0, 100_000).is_err());
    assert_eq!(read_entry(&mut reader, 1, 100).unwrap(), b"next");
}

#[test]
fn reading_stops_at_the_stated_size() {
    let archive = write_archive(&[("a.txt", b"0123456789")]);
    let mut reader = open(&archive).unwrap();
    reader.open_entry(0).unwrap();
    let mut buffer = [0; 64];
    assert_eq!(reader.read(&mut buffer).unwrap(), 10);
    assert_eq!(reader.read(&mut buffer).unwrap(), 0);
    assert_eq!(reader.read(&mut []).unwrap(), 0);
    reader.close_entry().unwrap();
}

#[test]
fn output_is_reproducible_and_laid_out_for_streaming() {
    let files: [(&str, &[u8]); 3] = [
        ("first.txt", b"first"),
        ("dir/a.txt", b"hello hello hello"),
        ("b.txt", b""),
    ];
    let archive = write_archive(&files);
    assert_eq!(archive, write_archive(&files));

    let records = central_directory(&archive);
    assert_eq!(records.len(), 3);
    for record in &records {
        assert_eq!(
            (record.dos_time, record.dos_date),
            (0, 0x0021),
            "1980-01-01 00:00:00"
        );
        assert_eq!(
            (
                record.version_made_by,
                record.version_needed,
                record.extra_len
            ),
            (20, 20, 0)
        );
        assert_eq!(
            (record.method, record.external_attributes),
            (METHOD_DEFLATED, 0)
        );
    }
    let file = &records[1];
    assert_eq!(
        (file.crc32, file.uncompressed_size),
        (crc32(b"hello hello hello"), 17)
    );
    let local = &archive[file.header_offset as usize..];
    assert_eq!(&local[..4], b"PK\x03\x04");
    assert_eq!(
        u16::from_le_bytes([local[6], local[7]]),
        FLAG_DATA_DESCRIPTOR
    );
    assert_eq!(&local[14..26], [0; 12], "sizes and CRC follow the data");
    let descriptor = &local[30 + file.name.len() + file.compressed_size as usize..][..16];
    assert_eq!(&descriptor[..4], b"PK\x07\x08");
    assert_eq!(descriptor[4..8], file.crc32.to_le_bytes());
    assert_eq!(descriptor[8..12], file.compressed_size.to_le_bytes());
    assert_eq!(descriptor[12..16], file.uncompressed_size.to_le_bytes());
}

#[test]
fn writes_reach_the_output_in_large_pieces() {
    let content = noise(3_000_000, 5);
    let stored = noise(1_000_000, 6);
    let source = RawArchive::of(vec![RawEntry::stored("stored.bin", &stored)]).build();
    let mut source = open(&source).unwrap();

    let mut writer = ZipWriter::new(Appended::default());
    writer.begin_file("small.txt").unwrap();
    writer.write_all(b"small").unwrap();
    writer.end_file().unwrap();
    writer.begin_file("dir/noise.bin").unwrap();
    for piece in content.chunks(1000) {
        writer.write_all(piece).unwrap();
    }
    writer.end_file().unwrap();
    writer.copy_entry(&mut source, 0).unwrap();
    writer.begin_file("last.txt").unwrap();
    writer.write_all(b"last").unwrap();
    writer.end_file().unwrap();
    let output = writer.finish().unwrap();

    assert_eq!(output.flushes, 1);
    let (_, full) = output.writes.split_last().unwrap();
    assert!(full.len() >= 60, "{:?}", output.writes);
    assert!(
        full.iter().all(|len| *len >= 64 * 1024),
        "{:?}",
        output.writes
    );
    let files = read_archive(&output.bytes).unwrap();
    assert_eq!(files[0].1, b"small");
    assert!(files[1].1 == content);
    assert!(files[2].1 == stored);
    assert_eq!(files[3].1, b"last");
}

#[test]
fn file_written_in_pieces_of_any_size() {
    let content = rows(1_500_000);
    let sizes = [
        1, 2, 0, 7, 100, 3000, 32_767, 32_768, 32_769, 1, 70_000, 5, 200_000, 11,
    ];
    let mut writer = ZipWriter::new(Vec::new());
    writer.begin_file("sheet.xml").unwrap();
    let mut rest = content.as_slice();
    for size in sizes.iter().cycle() {
        let (piece, tail) = rest.split_at((*size).min(rest.len()));
        assert_eq!(writer.write(piece).unwrap(), piece.len());
        rest = tail;
        if rest.is_empty() {
            break;
        }
    }
    writer.end_file().unwrap();
    writer.begin_file("after.xml").unwrap();
    writer.write_all(b"<after/>").unwrap();
    writer.end_file().unwrap();
    let archive = writer.finish().unwrap();

    let files = read_archive(&archive).unwrap();
    assert!(files[0].1 == content);
    assert_eq!(files[1].1, b"<after/>");
}

#[test]
fn flush_pushes_the_buffered_bytes() {
    let mut writer = ZipWriter::new(Appended::default());
    writer.flush().unwrap();
    writer.begin_file("a.txt").unwrap();
    writer.write_all(b"abc").unwrap();
    writer.flush().unwrap();
    writer.end_file().unwrap();
    let output = writer.finish().unwrap();
    assert_eq!(output.writes.len(), 2, "one write per flush that had bytes");
    assert!(output.writes[0] >= 30 + "a.txt".len());
    assert_eq!(output.flushes, 3);
    assert_eq!(read_archive(&output.bytes).unwrap()[0].1, b"abc");
}

#[test]
fn entries_are_fetched_in_few_reads_that_stay_inside_them() {
    let sheet = noise(1_000_000, 9);
    let archive = write_archive(&[
        ("small.xml", b"<small/>"),
        ("sheet.bin", &sheet),
        ("last.xml", b"<last/>"),
    ]);
    for (index, record) in central_directory(&archive).iter().enumerate() {
        let (file, log) = Watched::new(Cursor::new(&archive));
        let mut reader = ZipReader::new(file).unwrap();
        read_entry(&mut reader, index, 5000).unwrap();
        let log = log.borrow();

        let start = u64::from(record.header_offset);
        let end = start + 30 + record.name.len() as u64 + u64::from(record.compressed_size);
        let (directory_read, entry_reads) = log.reads.split_first().unwrap();
        assert_eq!(
            directory_read.0 + directory_read.1 as u64,
            archive.len() as u64
        );
        assert_eq!(entry_reads.len() as u64, (end - start).div_ceil(64 * 1024));
        let mut offset = start;
        for (read_offset, count) in entry_reads {
            assert_eq!(*read_offset, offset);
            offset += *count as u64;
        }
        assert_eq!(offset, end);
        assert_eq!(log.seeks, 3);
    }
}
