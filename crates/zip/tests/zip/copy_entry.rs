//! Entries copied from one archive into another without recompression.

use std::fs;
use std::io::{self, Cursor, ErrorKind, Write};

use excel_zip::{ZipReader, ZipWriter};

use crate::fixtures::fixture_paths;
use crate::support::{
    Central, FLAG_DATA_DESCRIPTOR, FLAG_ENCRYPTED, FLAG_UTF8, METHOD_DEFLATED, METHOD_STORED,
    RawArchive, RawEntry, Sparse, central_directory, end_records, kind, noise, open, read_archive,
    read_entry, rows, write_archive,
};

fn copy_all(source: &[u8]) -> Vec<u8> {
    let mut reader = open(source).unwrap();
    let mut writer = ZipWriter::new(Vec::new());
    for index in 0..reader.entries().len() {
        writer.copy_entry(&mut reader, index).unwrap();
    }
    writer.finish().unwrap()
}

#[test]
fn copies_an_archive_of_this_writer() {
    let sheet = rows(400_000);
    let image = noise(200_000, 21);
    let source = write_archive(&[
        ("xl/", b""),
        ("xl/worksheets/sheet1.xml", &sheet),
        ("xl/media/image.bin", &image),
        ("empty.xml", b""),
        ("données/été.xml", b"<chaud/>"),
    ]);
    let copy = copy_all(&source);
    assert_eq!(read_archive(&copy).unwrap(), read_archive(&source).unwrap());

    let (source_records, copy_records) = (central_directory(&source), central_directory(&copy));
    for (from, to) in source_records.iter().zip(&copy_records) {
        assert_eq!(to.flags, from.flags & !FLAG_DATA_DESCRIPTOR);
        assert_eq!(
            (
                &to.name,
                to.method,
                to.crc32,
                to.compressed_size,
                to.uncompressed_size
            ),
            (
                &from.name,
                from.method,
                from.crc32,
                from.compressed_size,
                from.uncompressed_size
            )
        );
        let local = &copy[to.header_offset as usize..][..30];
        assert_eq!(local[14..18], to.crc32.to_le_bytes());
        assert_eq!(local[18..22], to.compressed_size.to_le_bytes());
        assert_eq!(local[22..26], to.uncompressed_size.to_le_bytes());
    }
    assert_eq!(
        copy.len(),
        source.len() - 16 * 4,
        "the four data descriptors are gone"
    );
    assert_eq!(copy_all(&copy), copy);
}

#[test]
fn keeps_method_time_and_attributes() {
    let mut script = RawEntry::deflated("bin/run.sh", &rows(5000));
    script.version_made_by = 0x031E;
    script.external_attributes = 0o100755 << 16;
    (script.dos_time, script.dos_date) = (0x7B5D, 0x5962);
    script.descriptor = true;
    script.local_extra = vec![0; 28];
    let mut stored = RawEntry::stored("data.bin", &noise(3000, 8));
    stored.external_attributes = 0x20;
    stored.flags = 0x0006;
    stored.descriptor = true;
    let mut directory = RawEntry::stored("bin/", b"");
    directory.external_attributes = 0x10;
    let mut wide = RawEntry::deflated("wide.xml", &rows(100));
    (wide.zip64_sizes, wide.zip64_offset) = (true, true);
    let source = RawArchive {
        zip64: true,
        ..RawArchive::of(vec![script, stored, directory, wide])
    }
    .build();

    let copy = copy_all(&source);
    assert_eq!(read_archive(&copy).unwrap(), read_archive(&source).unwrap());
    let records = central_directory(&copy);
    let expected = |name: &str, content: &[u8], compressed_size: usize| Central {
        name: name.as_bytes().to_vec(),
        version_made_by: 20,
        flags: 0,
        method: METHOD_DEFLATED,
        dos_time: 0x6000,
        dos_date: 0x5A21,
        crc32: crc32fast::hash(content),
        compressed_size: compressed_size as u32,
        uncompressed_size: content.len() as u32,
        extra_len: 0,
        external_attributes: 0,
        header_offset: 0,
    };
    let offsets: Vec<u32> = records.iter().map(|record| record.header_offset).collect();
    assert_eq!(
        records[0],
        Central {
            version_made_by: 0x031E,
            external_attributes: 0o100755 << 16,
            dos_time: 0x7B5D,
            dos_date: 0x5962,
            ..expected(
                "bin/run.sh",
                &rows(5000),
                records[0].compressed_size as usize
            )
        }
    );
    assert_eq!(
        records[1],
        Central {
            method: METHOD_STORED,
            flags: 0x0006,
            external_attributes: 0x20,
            header_offset: offsets[1],
            ..expected("data.bin", &noise(3000, 8), 3000)
        }
    );
    assert_eq!(
        records[2],
        Central {
            method: METHOD_STORED,
            external_attributes: 0x10,
            header_offset: offsets[2],
            ..expected("bin/", b"", 0)
        }
    );
    assert_eq!(
        records[3],
        Central {
            header_offset: offsets[3],
            ..expected("wide.xml", &rows(100), records[3].compressed_size as usize)
        }
    );
    assert!(offsets.is_sorted() && offsets[1] > 0);
}

#[test]
fn copies_every_fixture() {
    for path in fixture_paths() {
        let source = fs::read(&path).unwrap();
        let copy = copy_all(&source);
        assert!(copy.len() <= source.len(), "{}", path.display());
        assert_eq!(
            read_archive(&copy).unwrap(),
            read_archive(&source).unwrap(),
            "{}",
            path.display()
        );
        assert!(
            central_directory(&copy)
                .iter()
                .all(|record| record.flags & FLAG_DATA_DESCRIPTOR == 0)
        );
    }
}

#[test]
fn mixes_copied_and_new_entries() {
    let first = write_archive(&[
        ("xl/workbook.xml", b"<old/>"),
        ("xl/styles.xml", b"<styles/>"),
    ]);
    let second =
        RawArchive::of(vec![RawEntry::stored("xl/media/logo.bin", &noise(100, 1))]).build();
    let (mut first, mut second) = (open(&first).unwrap(), open(&second).unwrap());

    let mut writer = ZipWriter::new(Vec::new());
    writer.add_directory("xl/").unwrap();
    writer.copy_entry(&mut first, 1).unwrap();
    writer.begin_file("xl/workbook.xml").unwrap();
    writer.write_all(b"<new/>").unwrap();
    writer.end_file().unwrap();
    writer.copy_entry(&mut second, 0).unwrap();
    writer.copy_entry(&mut first, 0).unwrap();
    writer.copy_entry(&mut first, 1).unwrap();
    let archive = writer.finish().unwrap();

    let files = read_archive(&archive).unwrap();
    let names: Vec<&str> = files.iter().map(|(name, _)| name.as_str()).collect();
    assert_eq!(
        names,
        [
            "xl/",
            "xl/styles.xml",
            "xl/workbook.xml",
            "xl/media/logo.bin",
            "xl/workbook.xml",
            "xl/styles.xml"
        ]
    );
    assert_eq!(files[2].1, b"<new/>");
    assert_eq!(files[4].1, b"<old/>");
    assert_eq!(files[3].1, noise(100, 1));

    assert_eq!(
        read_entry(&mut first, 0, 8).unwrap(),
        b"<old/>",
        "the source stays usable"
    );
}

#[test]
fn keeps_names_byte_for_byte() {
    let mut latin = RawEntry::stored("placeholder", b"latin");
    latin.name = b"caf\xE9.txt".to_vec();
    let unflagged = RawEntry::stored("crème.txt", b"utf8");
    let source = RawArchive::of(vec![
        latin,
        unflagged,
        RawEntry::stored("plain.txt", b"plain"),
    ])
    .build();

    let copy = copy_all(&source);
    let records = central_directory(&copy);
    assert_eq!(
        (records[0].name.as_slice(), records[0].flags),
        (&b"caf\xE9.txt"[..], 0)
    );
    assert_eq!(
        (records[1].name.as_slice(), records[1].flags),
        ("crème.txt".as_bytes(), FLAG_UTF8)
    );
    assert_eq!(
        (records[2].name.as_slice(), records[2].flags),
        (&b"plain.txt"[..], 0)
    );
    assert_eq!(read_archive(&copy).unwrap(), read_archive(&source).unwrap());
}

#[test]
fn refuses_what_it_cannot_copy() {
    let mut encrypted = RawEntry::stored("secret.bin", &noise(64, 2));
    encrypted.flags = FLAG_ENCRYPTED;
    let mut huge = RawEntry::deflated("huge.xml", &rows(1000));
    huge.uncompressed_size = 5 << 30;
    huge.zip64_sizes = true;
    let mut cut_short = RawEntry::stored("cut.bin", &noise(64, 3));
    cut_short.compressed_size = 1 << 20;
    let source = RawArchive::of(vec![
        RawEntry::stored("fine.txt", b"fine"),
        encrypted,
        huge,
        cut_short,
    ])
    .build();
    let mut reader = open(&source).unwrap();

    let mut writer = ZipWriter::new(Vec::new());
    assert_eq!(
        kind(writer.copy_entry(&mut reader, 1)),
        ErrorKind::Unsupported
    );
    assert_eq!(
        kind(writer.copy_entry(&mut reader, 2)),
        ErrorKind::Unsupported
    );
    assert_eq!(
        kind(writer.copy_entry(&mut reader, 4)),
        ErrorKind::InvalidInput
    );
    reader.open_entry(0).unwrap();
    assert_eq!(
        kind(writer.copy_entry(&mut reader, 0)),
        ErrorKind::InvalidInput
    );
    reader.close_entry().unwrap();
    assert_eq!(
        kind(writer.copy_entry(&mut reader, 3)),
        ErrorKind::InvalidData
    );
    writer.copy_entry(&mut reader, 0).unwrap();
    let archive = writer.finish().unwrap();

    let mut only_fine = ZipWriter::new(Vec::new());
    only_fine.copy_entry(&mut reader, 0).unwrap();
    assert_eq!(
        archive,
        only_fine.finish().unwrap(),
        "refused copies leave no trace"
    );
    assert_eq!(
        read_archive(&archive).unwrap(),
        [("fine.txt".to_string(), b"fine".to_vec())]
    );
}

#[test]
fn copies_an_entry_the_reader_cannot_decode() {
    let mut exotic = RawEntry::stored("exotic.bin", &noise(500, 6));
    exotic.method = 93;
    exotic.uncompressed_size = 2000;
    let source = RawArchive::of(vec![exotic]).build();
    let mut reader = open(&source).unwrap();
    assert_eq!(kind(reader.open_entry(0)), ErrorKind::Unsupported);

    let copy = copy_all(&source);
    let records = central_directory(&copy);
    assert_eq!(
        (
            records[0].method,
            records[0].compressed_size,
            records[0].uncompressed_size
        ),
        (93, 500, 2000)
    );
    let data_start = 30 + "exotic.bin".len();
    assert_eq!(copy[data_start..data_start + 500], noise(500, 6));
}

#[test]
fn refuses_to_grow_beyond_4_gib() {
    let piece_len: u64 = 3 << 29;
    let mut file = Sparse::new();
    let mut directory = Vec::new();
    for name in ["a.bin", "b.bin", "c.bin"] {
        let mut entry = RawEntry::stored(name, b"");
        (entry.compressed_size, entry.uncompressed_size) = (piece_len, piece_len);
        (entry.zip64_sizes, entry.zip64_offset) = (true, true);
        directory.extend_from_slice(&entry.central_header(file.len()));
        file.append(entry.local_header());
        file.append_zeros(piece_len);
    }
    let (directory_len, directory_offset) = (directory.len() as u64, file.len());
    file.append(directory);
    file.append(end_records(3, directory_len, directory_offset, true, b""));

    let mut reader = ZipReader::new(file).unwrap();
    assert_eq!(reader.entries()[2].uncompressed_size, piece_len);
    let mut writer = ZipWriter::new(io::sink());
    for index in 0..3 {
        writer.copy_entry(&mut reader, index).unwrap();
    }
    assert_eq!(
        kind(writer.copy_entry(&mut reader, 0)),
        ErrorKind::Unsupported
    );
    assert_eq!(kind(writer.add_directory("late/")), ErrorKind::Unsupported);
    assert_eq!(kind(writer.begin_file("late.txt")), ErrorKind::Unsupported);
    assert_eq!(kind(writer.finish()), ErrorKind::Unsupported);
}

#[test]
fn source_file_may_return_short_reads() {
    let sheet = rows(300_000);
    let source = write_archive(&[("sheet.xml", &sheet), ("tail.txt", b"tail")]);
    let mut reader = ZipReader::new(crate::support::Watched::trickling(
        Cursor::new(&source),
        1000,
    ))
    .unwrap();
    let mut writer = ZipWriter::new(Vec::new());
    writer.copy_entry(&mut reader, 0).unwrap();
    writer.copy_entry(&mut reader, 1).unwrap();
    let copy = writer.finish().unwrap();
    assert_eq!(read_archive(&copy).unwrap(), read_archive(&source).unwrap());
}
