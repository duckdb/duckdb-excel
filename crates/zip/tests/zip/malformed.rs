//! Broken and hostile archives: every one must end in an error, never in a panic or a hang.

use std::fs;
use std::io::{self, ErrorKind, Read};

use excel_zip::ZipWriter;

use crate::fixtures::fixture_paths;
use crate::support::{
    FLAG_ENCRYPTED, RawArchive, RawEntry, crc32, end_records, kind, noise, open, read_entry, rows,
    write_archive,
};

fn exercise(archive: &[u8]) {
    let Ok(mut reader) = open(archive) else {
        return;
    };
    let mut copies = ZipWriter::new(io::sink());
    let mut chunk = [0; 4096];
    for index in 0..reader.entries().len().min(100) {
        let stated = reader.entries()[index].uncompressed_size;
        if reader.open_entry(index).is_ok() {
            let mut total = 0;
            while let Ok(count) = reader.read(&mut chunk) {
                total += count as u64;
                assert!(total <= stated && total <= 64 << 20);
                if count == 0 {
                    break;
                }
            }
            let _ = reader.close_entry();
        }
        let _ = copies.copy_entry(&mut reader, index);
    }
}

fn small_archives() -> Vec<Vec<u8>> {
    let written = write_archive(&[
        ("xl/", b""),
        ("xl/workbook.xml", &rows(700)),
        ("xl/media/image.bin", &noise(300, 5)),
        ("empty.xml", b""),
    ]);
    let mut described = RawEntry::deflated("sheet.xml", &rows(900));
    described.descriptor = true;
    described.local_extra = vec![0; 20];
    let mut wide = RawEntry::stored("wide.bin", &noise(200, 6));
    (wide.zip64_sizes, wide.zip64_offset) = (true, true);
    let entries = vec![described, wide, RawEntry::stored("dir/", b"")];
    let classic = RawArchive {
        comment: b"comment".to_vec(),
        ..RawArchive::of(entries.clone())
    }
    .build();
    let zip64 = RawArchive {
        zip64: true,
        ..RawArchive::of(entries)
    }
    .build();
    vec![written, classic, zip64]
}

#[test]
fn not_an_archive() {
    let mut almost = vec![0; 100];
    almost.extend_from_slice(b"PK\x05\x06");
    almost.extend_from_slice(&[0; 17]);
    for bytes in [
        Vec::new(),
        b"PK".to_vec(),
        vec![0; 22],
        noise(1000, 1),
        noise(200_000, 2),
        almost,
    ] {
        assert_eq!(kind(open(&bytes)), ErrorKind::InvalidData);
    }
}

#[test]
fn truncated_at_every_length() {
    for archive in small_archives() {
        for len in 0..archive.len() {
            let truncated = &archive[..len];
            exercise(truncated);
            if let Ok(mut reader) = open(truncated) {
                let all_read = (0..reader.entries().len())
                    .all(|index| read_entry(&mut reader, index, 512).is_ok());
                assert!(
                    !all_read,
                    "a prefix of {len} bytes reads like the whole archive"
                );
            }
        }
        for len in 1..archive.len() {
            exercise(&archive[len..]);
        }
    }
}

#[test]
fn every_byte_mutated() {
    let mut archives = small_archives();
    let fixtures = fixture_paths();
    for name in [
        "gdal/with_xml_prefix.xlsx",
        "gdal/row_without_r_attribute.xlsx",
    ] {
        let path = fixtures.iter().find(|path| path.ends_with(name)).unwrap();
        archives.push(fs::read(path).unwrap());
    }
    for archive in archives {
        let mut mutated = archive.clone();
        for at in 0..archive.len() {
            for mask in [0x01, 0x80, 0xFF] {
                mutated[at] = archive[at] ^ mask;
                exercise(&mutated);
            }
            mutated[at] = archive[at];
        }
    }
}

#[test]
fn random_bytes_mutated() {
    let archives = small_archives();
    for round in 0..20_000 {
        let archive = &archives[round % archives.len()];
        let mut mutated = archive.clone();
        for draw in noise(8 * (1 + round % 6), round as u64).as_chunks::<8>().0 {
            let position = (u64::from_le_bytes(*draw) >> 8) as usize % archive.len();
            mutated[position] = draw[0];
        }
        exercise(&mutated);
    }
}

#[test]
fn wrong_checksum() {
    let content = rows(20_000);
    for mut entry in [
        RawEntry::deflated("a.xml", &content),
        RawEntry::stored("a.xml", &content),
    ] {
        entry.crc32 ^= 0x8000_0000;
        let archive = RawArchive::of(vec![entry]).build();
        let mut reader = open(&archive).unwrap();
        reader.open_entry(0).unwrap();
        let mut read = Vec::new();
        reader.read_to_end(&mut read).unwrap();
        assert!(read == content);
        assert_eq!(kind(reader.close_entry()), ErrorKind::InvalidData);
        assert!(!reader.is_entry_open());
    }
}

#[test]
fn corrupt_data() {
    let content = rows(200_000);
    let mut deflated = RawEntry::deflated("a.xml", &content);
    let middle = deflated.data.len() / 2;
    deflated.data[middle..middle + 64].fill(0xFF);
    let mut stored = RawEntry::stored("b.bin", &content);
    stored.data[100_000] ^= 1;
    let archive = RawArchive::of(vec![deflated, stored]).build();
    let mut reader = open(&archive).unwrap();
    assert_eq!(
        kind(read_entry(&mut reader, 0, 4096)),
        ErrorKind::InvalidData
    );
    reader.close_entry().unwrap();
    assert_eq!(
        kind(read_entry(&mut reader, 1, 4096)),
        ErrorKind::InvalidData
    );
}

#[test]
fn bomb_that_understates_its_size() {
    let content = vec![b'A'; 8 << 20];
    for stated in [0, 100, 70_000] {
        let mut entry = RawEntry::deflated("bomb.xml", &content);
        entry.uncompressed_size = stated;
        entry.crc32 = crc32(&content[..stated as usize]);
        let archive = RawArchive::of(vec![entry]).build();
        let mut reader = open(&archive).unwrap();
        reader.open_entry(0).unwrap();
        let mut read = Vec::new();
        reader.read_to_end(&mut read).unwrap();
        assert_eq!(read.len() as u64, stated);
        assert_eq!(kind(reader.close_entry()), ErrorKind::InvalidData);
    }
}

#[test]
fn entry_that_overstates_its_size() {
    let content = rows(1000);
    for stated in [1001, 1 << 30, u64::from(u32::MAX) - 1] {
        let mut entry = RawEntry::deflated("short.xml", &content);
        entry.uncompressed_size = stated;
        let archive = RawArchive::of(vec![entry]).build();
        let mut reader = open(&archive).unwrap();
        reader.open_entry(0).unwrap();
        assert_eq!(reader.entry_len(), stated);
        let mut read = Vec::new();
        assert_eq!(kind(reader.read_to_end(&mut read)), ErrorKind::InvalidData);
        assert!(read == content);
        assert!(!reader.is_done());
        reader.close_entry().unwrap();
    }
}

#[test]
fn deflate_stream_cut_short() {
    let content = rows(100_000);
    let mut entry = RawEntry::deflated("cut.xml", &content);
    entry.compressed_size /= 2;
    entry.data.truncate(entry.compressed_size as usize);
    let archive = RawArchive::of(vec![entry, RawEntry::stored("next.txt", b"next")]).build();
    let mut reader = open(&archive).unwrap();
    assert_eq!(
        kind(read_entry(&mut reader, 0, 4096)),
        ErrorKind::InvalidData
    );
    reader.close_entry().unwrap();
    assert_eq!(read_entry(&mut reader, 1, 16).unwrap(), b"next");
}

#[test]
fn stored_entry_with_two_sizes() {
    let mut entry = RawEntry::stored("a.bin", &noise(100, 1));
    entry.uncompressed_size = 1 << 30;
    let archive = RawArchive::of(vec![entry]).build();
    assert_eq!(
        kind(open(&archive).unwrap().open_entry(0)),
        ErrorKind::InvalidData
    );
}

#[test]
fn entry_data_outside_the_file() {
    let content = noise(100, 1);
    for compressed_size in [5000, u64::from(u32::MAX) - 1, u64::MAX, u64::MAX - 40] {
        let mut entry = RawEntry::deflated("a.bin", &content);
        entry.compressed_size = compressed_size;
        entry.zip64_sizes = compressed_size > u64::from(u32::MAX);
        let archive = RawArchive::of(vec![entry]).build();
        let mut reader = open(&archive).unwrap();
        assert_eq!(reader.entries()[0].compressed_size, compressed_size);
        assert_eq!(
            kind(reader.open_entry(0)),
            ErrorKind::InvalidData,
            "compressed size {compressed_size}"
        );
        assert!(!reader.is_entry_open());
    }
}

#[test]
fn entry_header_outside_the_file_or_elsewhere() {
    let entry = RawEntry::stored("a.txt", b"abcdefgh");
    let mut body = entry.local_header();
    body.extend_from_slice(&entry.data);
    let directory_offset = body.len() as u64;
    for header_offset in [
        1,
        20,
        directory_offset,
        directory_offset + 60,
        1 << 31,
        u64::MAX,
        u64::MAX - 29,
    ] {
        let mut lying = entry.clone();
        lying.zip64_offset = header_offset > u64::from(u32::MAX);
        let directory = lying.central_header(header_offset);
        let mut archive = body.clone();
        archive.extend_from_slice(&directory);
        archive.extend_from_slice(&end_records(
            1,
            directory.len() as u64,
            directory_offset,
            false,
            b"",
        ));
        let mut reader = open(&archive).unwrap();
        assert_eq!(
            kind(reader.open_entry(0)),
            ErrorKind::InvalidData,
            "header offset {header_offset}"
        );
    }
}

#[test]
fn directory_outside_the_file() {
    let entry = RawEntry::stored("a.txt", b"abcdefgh");
    let mut body = entry.local_header();
    body.extend_from_slice(&entry.data);
    let directory = entry.central_header(0);
    let (len, offset) = (directory.len() as u64, body.len() as u64);
    body.extend_from_slice(&directory);
    let cases = [
        (1, len, offset + 1),
        (1, len, 1 << 31),
        (1, len - 1, offset),
        (1, 0, offset),
        (2, len, offset),
        (u64::from(u16::MAX), len, offset),
    ];
    for (count, len, offset) in cases {
        let mut archive = body.clone();
        archive.extend_from_slice(&end_records(count, len, offset, false, b""));
        assert_eq!(
            kind(open(&archive)),
            ErrorKind::InvalidData,
            "count {count} len {len} offset {offset}"
        );
    }
}

#[test]
fn unsupported_features() {
    let mut encrypted = RawEntry::stored("secret.bin", &noise(100, 1));
    encrypted.flags = FLAG_ENCRYPTED;
    let mut exotic = RawEntry::stored("exotic.bin", &noise(100, 2));
    exotic.method = 12;
    let archive = RawArchive::of(vec![
        encrypted,
        exotic,
        RawEntry::stored("fine.txt", b"fine"),
    ])
    .build();
    let mut reader = open(&archive).unwrap();
    assert_eq!(kind(reader.open_entry(0)), ErrorKind::Unsupported);
    assert_eq!(kind(reader.open_entry(1)), ErrorKind::Unsupported);
    assert!(!reader.is_entry_open());
    assert_eq!(read_entry(&mut reader, 2, 16).unwrap(), b"fine");
}

#[test]
fn multi_disk_archive() {
    let good = RawArchive::of(vec![RawEntry::stored("a.txt", b"abc")]).build();
    let end = good.len() - 22;
    for (field, value) in [(4, 1u16), (6, 1), (8, 0), (8, 2)] {
        let mut archive = good.clone();
        archive[end + field..end + field + 2].copy_from_slice(&value.to_le_bytes());
        assert_eq!(
            kind(open(&archive)),
            ErrorKind::Unsupported,
            "field at {field} set to {value}"
        );
    }
}

#[test]
fn file_that_ends_while_an_entry_is_read() {
    struct Shrinking<'a> {
        bytes: &'a [u8],
        position: u64,
        reads_left: usize,
    }
    impl Read for Shrinking<'_> {
        fn read(&mut self, buffer: &mut [u8]) -> io::Result<usize> {
            if self.reads_left == 0 {
                return Ok(0);
            }
            self.reads_left -= 1;
            let rest = &self.bytes[self.position as usize..];
            let count = rest.len().min(buffer.len());
            buffer[..count].copy_from_slice(&rest[..count]);
            self.position += count as u64;
            Ok(count)
        }
    }
    impl io::Seek for Shrinking<'_> {
        fn seek(&mut self, position: io::SeekFrom) -> io::Result<u64> {
            self.position = match position {
                io::SeekFrom::Start(offset) => offset,
                _ => self.bytes.len() as u64,
            };
            Ok(self.position)
        }
    }

    let archive = RawArchive::of(vec![
        RawEntry::deflated("a.xml", &rows(2_000_000)),
        RawEntry::stored("b.bin", &noise(500_000, 3)),
    ])
    .build();
    for index in 0..2 {
        for reads_left in 1..4 {
            let file = Shrinking {
                bytes: &archive,
                position: 0,
                reads_left,
            };
            let result = excel_zip::ZipReader::new(file)
                .and_then(|mut reader| read_entry(&mut reader, index, 8192));
            assert_eq!(
                kind(result),
                ErrorKind::InvalidData,
                "entry {index} after {reads_left} reads"
            );
        }
    }
}
