//! ZIP64 archives, which the reader accepts and the writer never produces.

use std::io::ErrorKind;

use crate::support::{RawArchive, RawEntry, end_records, kind, noise, open, read_archive, rows};

fn zip64_entries() -> Vec<RawEntry> {
    let mut entries = vec![
        RawEntry::deflated("xl/worksheets/sheet1.xml", &rows(50_000)),
        RawEntry::stored("xl/media/image.bin", &noise(3000, 4)),
        RawEntry::deflated("xl/workbook.xml", b"<workbook/>"),
        RawEntry::stored("docProps/", b""),
    ];
    (entries[0].zip64_sizes, entries[0].zip64_offset) = (true, true);
    entries[1].zip64_offset = true;
    entries[2].zip64_sizes = true;
    entries[2].descriptor = true;
    entries
}

fn assert_content(archive: &[u8]) {
    let files = read_archive(archive).unwrap();
    let names: Vec<&str> = files.iter().map(|(name, _)| name.as_str()).collect();
    assert_eq!(
        names,
        [
            "xl/worksheets/sheet1.xml",
            "xl/media/image.bin",
            "xl/workbook.xml",
            "docProps/"
        ]
    );
    assert!(files[0].1 == rows(50_000));
    assert!(files[1].1 == noise(3000, 4));
    assert_eq!(files[2].1, b"<workbook/>");
    assert!(files[3].1.is_empty());
}

#[test]
fn zip64_end_record_and_extra_fields() {
    let archive = RawArchive {
        zip64: true,
        ..RawArchive::of(zip64_entries())
    }
    .build();
    let end = &archive[archive.len() - 22..];
    assert_eq!(
        end[8..20],
        [0xFF; 12],
        "the classic end record defers to the ZIP64 one"
    );
    assert_content(&archive);

    let reader = open(&archive).unwrap();
    assert_eq!(reader.entries()[0].uncompressed_size, 50_000);
    assert_eq!(reader.entries()[1].compressed_size, 3000);
}

#[test]
fn zip64_extra_fields_with_a_classic_end_record() {
    assert_content(&RawArchive::of(zip64_entries()).build());
}

#[test]
fn zip64_records_outside_the_last_64_kib() {
    let archive = RawArchive {
        zip64: true,
        comment: vec![b'-'; 65_535],
        ..RawArchive::of(zip64_entries())
    };
    assert_content(&archive.build());
}

#[test]
fn zip64_with_more_entries_than_the_classic_count() {
    let entries: Vec<RawEntry> = (0..66_000)
        .map(|index| RawEntry::stored(&format!("e{index}"), &[index as u8]))
        .collect();
    let archive = RawArchive {
        zip64: true,
        ..RawArchive::of(entries)
    }
    .build();
    let mut reader = open(&archive).unwrap();
    assert_eq!(reader.entries().len(), 66_000);
    assert_eq!(reader.find("e65999"), Some(65_999));
    assert_eq!(
        crate::support::read_entry(&mut reader, 65_999, 8).unwrap(),
        [65_999u32 as u8]
    );
}

#[test]
fn zip64_extra_field_without_the_marked_values() {
    let mut entry = RawEntry::stored("a.txt", b"abc");
    entry.uncompressed_size = u64::from(u32::MAX);
    entry.zip64_offset = true;
    let archive = RawArchive::of(vec![entry]).build();
    assert_eq!(kind(open(&archive)), ErrorKind::InvalidData);
}

#[test]
fn zip64_locator_pointing_nowhere() {
    let entries = zip64_entries();
    let good = RawArchive {
        zip64: true,
        ..RawArchive::of(entries)
    }
    .build();
    let locator = good.len() - 22 - 20;
    assert_eq!(&good[locator..locator + 4], b"PK\x06\x07");
    for offset in [
        u64::MAX,
        u64::MAX - 55,
        good.len() as u64,
        locator as u64 - 55,
        0,
    ] {
        let mut archive = good.clone();
        archive[locator + 8..locator + 16].copy_from_slice(&offset.to_le_bytes());
        assert_eq!(
            kind(open(&archive)),
            ErrorKind::InvalidData,
            "offset {offset}"
        );
    }
}

#[test]
fn zip64_entry_count_far_beyond_the_directory() {
    let entry = RawEntry::stored("a.txt", b"abc");
    let mut archive = entry.local_header();
    archive.extend_from_slice(&entry.data);
    let offset = archive.len() as u64;
    let directory = entry.central_header(0);
    archive.extend_from_slice(&directory);
    for count in [2, 1 << 40, u64::MAX] {
        let mut lying = archive.clone();
        lying.extend_from_slice(&end_records(
            count,
            directory.len() as u64,
            offset,
            true,
            b"",
        ));
        assert_eq!(kind(open(&lying)), ErrorKind::InvalidData, "count {count}");
    }
}
