//! The XLSX files of the SQL tests: written by Excel, LibreOffice, Google Sheets and GDAL.

use std::fs::{self, File};
use std::path::{Path, PathBuf};

use excel_zip::ZipReader;

use crate::support::{FLAG_DATA_DESCRIPTOR, central_directory, crc32, read_entry};

fn collect(directory: &Path, paths: &mut Vec<PathBuf>) {
    for item in fs::read_dir(directory).unwrap() {
        let path = item.unwrap().path();
        if path.is_dir() {
            collect(&path, paths);
        } else if path
            .extension()
            .is_some_and(|extension| extension == "xlsx")
        {
            paths.push(path);
        }
    }
}

pub fn fixture_paths() -> Vec<PathBuf> {
    let root = Path::new(env!("CARGO_MANIFEST_DIR")).join("../../test/data/xlsx");
    let mut paths = Vec::new();
    collect(&root, &mut paths);
    paths.sort();
    paths
}

pub fn fixture(name: &str) -> PathBuf {
    let found = fixture_paths()
        .into_iter()
        .find(|path| path.ends_with(name));
    found.unwrap_or_else(|| panic!("no fixture {name}"))
}

#[test]
fn every_fixture_reads_with_valid_checksums() {
    for path in fixture_paths() {
        let mut reader = ZipReader::new(File::open(&path).unwrap()).unwrap();
        let entries = reader.entries().to_vec();
        assert!(!entries.is_empty(), "{}", path.display());
        assert!(
            reader.find("[Content_Types].xml").is_some(),
            "{}",
            path.display()
        );
        for (index, entry) in entries.iter().enumerate() {
            let content = read_entry(&mut reader, index, 4096)
                .unwrap_or_else(|error| panic!("{}: '{}': {error}", path.display(), entry.name));
            assert_eq!(content.len() as u64, entry.uncompressed_size);
            assert_eq!(crc32(&content), entry.crc32);
        }
    }
}

#[test]
fn fixtures_cover_data_descriptors_and_plain_headers() {
    let uses_descriptors = |name: &str| -> Vec<bool> {
        central_directory(&fs::read(fixture(name)).unwrap())
            .iter()
            .map(|record| record.flags & FLAG_DATA_DESCRIPTOR != 0)
            .collect()
    };
    assert!(uses_descriptors("gdal/test.xlsx").iter().all(|used| *used));
    assert!(uses_descriptors("basic.xlsx").iter().all(|used| !*used));
}
