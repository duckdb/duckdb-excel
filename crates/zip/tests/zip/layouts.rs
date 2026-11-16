//! Archive layouts that other producers write, built by hand.

use std::io::Cursor;

use excel_zip::ZipReader;

use crate::support::{
    FLAG_UTF8, RawArchive, RawEntry, Watched, end_records, noise, open, read_archive, read_entry,
    rows,
};

fn sample_entries() -> Vec<RawEntry> {
    vec![
        RawEntry::deflated("[Content_Types].xml", &rows(3000)),
        RawEntry::stored("xl/", b""),
        RawEntry::stored("xl/media/image.bin", &noise(5000, 1)),
        RawEntry::deflated("xl/worksheets/sheet1.xml", &rows(100_000)),
    ]
}

fn assert_sample(archive: &[u8]) {
    let files = read_archive(archive).unwrap();
    let names: Vec<&str> = files.iter().map(|(name, _)| name.as_str()).collect();
    assert_eq!(
        names,
        [
            "[Content_Types].xml",
            "xl/",
            "xl/media/image.bin",
            "xl/worksheets/sheet1.xml"
        ]
    );
    assert!(files[0].1 == rows(3000));
    assert!(files[1].1.is_empty());
    assert!(files[2].1 == noise(5000, 1));
    assert!(files[3].1 == rows(100_000));
}

#[test]
fn plain_archive() {
    assert_sample(&RawArchive::of(sample_entries()).build());
}

#[test]
fn data_descriptors_with_empty_local_sizes() {
    let mut entries = sample_entries();
    for entry in &mut entries {
        entry.descriptor = true;
    }
    assert_sample(&RawArchive::of(entries).build());
}

#[test]
fn local_extra_fields_unknown_to_the_directory() {
    for extra_len in [4, 20, 520, 5000, 65_535] {
        let mut entries = sample_entries();
        for entry in &mut entries {
            entry.local_extra = vec![0; extra_len];
        }
        assert_sample(&RawArchive::of(entries).build());
    }
}

#[test]
fn local_extra_fields_cost_one_more_read_at_most() {
    let content = rows(3000);
    let mut entry = RawEntry::deflated("sheet.xml", &content);
    entry.local_extra = vec![0; 520];
    let archive = RawArchive::of(vec![entry]).build();
    let (file, log) = Watched::new(Cursor::new(&archive));
    let mut reader = ZipReader::new(file).unwrap();
    assert!(read_entry(&mut reader, 0, 100).unwrap() == content);
    assert!(log.borrow().reads.len() <= 3);
}

#[test]
fn archive_comment() {
    for comment_len in [1, 300, 65_535] {
        let archive = RawArchive {
            comment: vec![b'c'; comment_len],
            ..RawArchive::of(sample_entries())
        };
        assert_sample(&archive.build());
    }
}

#[test]
fn archive_comment_that_looks_like_an_end_record() {
    let mut comment = b"see PK\x05\x06".to_vec();
    comment.extend_from_slice(&[0; 40]);
    let archive = RawArchive {
        comment,
        ..RawArchive::of(sample_entries())
    };
    assert_sample(&archive.build());
}

#[test]
fn data_in_front_of_the_archive() {
    for zip64 in [false, true] {
        let mut entries = sample_entries();
        entries
            .iter_mut()
            .for_each(|entry| (entry.zip64_sizes, entry.zip64_offset) = (zip64, zip64));
        let archive = RawArchive {
            prefix: noise(70_000, 2),
            ..RawArchive::of(entries)
        };
        assert_sample(&archive.build());
    }
}

#[test]
fn name_that_is_not_utf8() {
    let mut entry = RawEntry::stored("placeholder", b"latin");
    entry.name = b"caf\xE9.txt".to_vec();
    let mut flagged = RawEntry::stored("crème.txt", b"utf8");
    flagged.flags = FLAG_UTF8;
    let archive = RawArchive::of(vec![entry, flagged]).build();
    let mut reader = open(&archive).unwrap();
    assert_eq!(reader.entries()[0].name, "caf\u{FFFD}.txt");
    assert_eq!(reader.entries()[1].name, "crème.txt");
    assert_eq!(reader.find("caf\u{FFFD}.txt"), Some(0));
    assert_eq!(read_entry(&mut reader, 0, 64).unwrap(), b"latin");
    assert_eq!(read_entry(&mut reader, 1, 64).unwrap(), b"utf8");
}

#[test]
fn stated_values_are_reported_as_they_are() {
    let content = rows(10_000);
    let entry = RawEntry::deflated("sheet.xml", &content);
    let archive = RawArchive::of(vec![entry.clone()]).build();
    let reader = open(&archive).unwrap();
    let listed = &reader.entries()[0];
    assert_eq!(listed.name, "sheet.xml");
    assert_eq!(listed.compressed_size, entry.data.len() as u64);
    assert_eq!(listed.uncompressed_size, 10_000);
    assert_eq!(listed.crc32, entry.crc32);
    assert!(!listed.is_directory());
}

#[test]
fn stray_bytes_after_the_end_record() {
    let mut archive = RawArchive::of(sample_entries()).build();
    archive.extend_from_slice(&[0; 300]);
    assert_sample(&archive);
}

#[test]
fn directory_shorter_than_the_end_record_says() {
    let entry = RawEntry::stored("a.txt", b"abc");
    let mut archive = entry.local_header();
    archive.extend_from_slice(&entry.data);
    let offset = archive.len() as u64;
    archive.extend_from_slice(&entry.central_header(0));
    archive.extend_from_slice(&end_records(1, 1 << 31, offset, false, b""));
    assert_eq!(
        read_archive(&archive).unwrap(),
        [("a.txt".to_string(), b"abc".to_vec())]
    );
}

#[test]
fn records_larger_than_the_directory_window() {
    let names: Vec<String> = [65_535, 40_000, 70, 65_000, 1, 65_535]
        .iter()
        .enumerate()
        .map(|(index, len)| format!("{index}").repeat(*len))
        .collect();
    let entries: Vec<RawEntry> = names
        .iter()
        .map(|name| RawEntry::stored(name, &name.as_bytes()[..1]))
        .collect();
    for comment_len in [0, 65_535] {
        let archive = RawArchive {
            comment: vec![b'c'; comment_len],
            ..RawArchive::of(entries.clone())
        }
        .build();
        let files = read_archive(&archive).unwrap();
        assert_eq!(files.len(), names.len());
        for ((name, content), expected) in files.iter().zip(&names) {
            assert!(name == expected);
            assert_eq!(content, &expected.as_bytes()[..1]);
        }
    }
}
