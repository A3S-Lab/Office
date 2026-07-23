use std::io::Write;
use std::path::Path;

use tempfile::TempDir;
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipWriter};

use super::{DocumentKind, NativeOfficeDocument, NativeOfficePackage, PackageLimits};

const CONTENT_TYPES: &[u8] = br#"<?xml version="1.0"?><Types/>"#;
const ROOT_RELATIONSHIPS: &[u8] = br#"<?xml version="1.0"?><Relationships/>"#;

#[tokio::test]
async fn native_package_detects_each_supported_document_kind() {
    for (extension, main_part, kind) in [
        ("docx", "word/document.xml", DocumentKind::Word),
        ("xlsx", "xl/workbook.xml", DocumentKind::Spreadsheet),
        ("pptx", "ppt/presentation.xml", DocumentKind::Presentation),
    ] {
        let temp = tempfile::tempdir().unwrap();
        let path = temp.path().join(format!("document.{extension}"));
        write_package(
            &path,
            &[
                ("[Content_Types].xml", CONTENT_TYPES),
                ("_rels/.rels", ROOT_RELATIONSHIPS),
                (main_part, b"<document/>"),
            ],
        );

        let package = NativeOfficePackage::open(&path).await.unwrap();

        assert_eq!(package.kind(), kind);
        assert_eq!(package.path(), path.as_path());
        assert_eq!(package.part(main_part).unwrap(), b"<document/>");
        assert!(!package.is_dirty());
    }
}

#[tokio::test]
async fn native_package_creates_readable_blank_formats_without_clobbering() {
    let temp = tempfile::tempdir().unwrap();
    for (extension, kind) in [
        ("docx", DocumentKind::Word),
        ("xlsx", DocumentKind::Spreadsheet),
        ("pptx", DocumentKind::Presentation),
    ] {
        let path = temp.path().join(format!("blank.{extension}"));

        let package = NativeOfficePackage::create(&path).await.unwrap();

        assert_eq!(package.kind(), kind);
        assert_eq!(package.path(), path.as_path());
        assert!(!package.is_dirty());
        let document = NativeOfficeDocument::open(&path).await.unwrap();
        assert_eq!(document.kind(), kind);
        let original = std::fs::read(&path).unwrap();

        let error = NativeOfficePackage::create(&path).await.unwrap_err();

        assert_eq!(error.code, "use.office.package_exists");
        assert_eq!(std::fs::read(&path).unwrap(), original);
    }
}

#[tokio::test]
async fn native_package_create_rejects_an_unsupported_extension() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("blank.odt");

    let error = NativeOfficePackage::create(&path).await.unwrap_err();

    assert_eq!(error.code, "use.office.package_extension_unsupported");
    assert!(!path.exists());
}

#[tokio::test]
async fn native_package_rejects_extension_and_content_mismatch() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("spoofed.docx");
    write_package(
        &path,
        &[
            ("[Content_Types].xml", CONTENT_TYPES),
            ("_rels/.rels", ROOT_RELATIONSHIPS),
            ("xl/workbook.xml", b"<workbook/>"),
        ],
    );

    let error = NativeOfficePackage::open(&path).await.unwrap_err();

    assert_eq!(error.code, "use.office.package_kind_mismatch");
}

#[tokio::test]
async fn native_package_rejects_ambiguous_and_unsafe_part_names() {
    let temp = tempfile::tempdir().unwrap();
    let duplicate = temp.path().join("duplicate.docx");
    write_package(
        &duplicate,
        &[
            ("[Content_Types].xml", CONTENT_TYPES),
            ("_rels/.rels", ROOT_RELATIONSHIPS),
            ("word/document.xml", b"<document/>"),
            ("Word/Document.xml", b"<duplicate/>"),
        ],
    );
    let error = NativeOfficePackage::open(&duplicate).await.unwrap_err();
    assert_eq!(error.code, "use.office.package_part_duplicate");

    let baseline = fixture(DocumentKind::Word);
    let mut package = NativeOfficePackage::open(baseline.path().join("document.docx"))
        .await
        .unwrap();
    let error = package
        .set_part("Word/Document.xml", b"<duplicate/>".to_vec())
        .unwrap_err();
    assert_eq!(error.code, "use.office.package_part_duplicate");

    let unsafe_path = temp.path().join("unsafe.docx");
    write_package(
        &unsafe_path,
        &[
            ("[Content_Types].xml", CONTENT_TYPES),
            ("_rels/.rels", ROOT_RELATIONSHIPS),
            ("word/document.xml", b"<document/>"),
            ("../payload", b"unsafe"),
        ],
    );
    let error = NativeOfficePackage::open(&unsafe_path).await.unwrap_err();
    assert_eq!(error.code, "use.office.package_part_invalid");
}

#[tokio::test]
async fn native_package_accepts_safe_zip_directory_entries() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("directories.docx");
    write_package(
        &path,
        &[
            ("word/", b""),
            ("[Content_Types].xml", CONTENT_TYPES),
            ("_rels/.rels", ROOT_RELATIONSHIPS),
            ("word/document.xml", b"<document/>"),
        ],
    );

    let package = NativeOfficePackage::open(&path).await.unwrap();

    assert_eq!(package.kind(), DocumentKind::Word);
}

#[tokio::test]
async fn native_package_enforces_bounded_archive_limits() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("large.docx");
    write_package(
        &path,
        &[
            ("[Content_Types].xml", CONTENT_TYPES),
            ("_rels/.rels", ROOT_RELATIONSHIPS),
            ("word/document.xml", &[b'x'; 64]),
        ],
    );
    let limits = PackageLimits {
        max_part_bytes: 32,
        ..PackageLimits::default()
    };

    let error = NativeOfficePackage::open_with_limits(&path, limits)
        .await
        .unwrap_err();

    assert_eq!(error.code, "use.office.package_part_too_large");
}

#[tokio::test]
async fn native_package_enforces_entry_limits_during_mutation() {
    let temp = fixture(DocumentKind::Word);
    let path = temp.path().join("document.docx");
    let limits = PackageLimits {
        max_entries: 3,
        ..PackageLimits::default()
    };
    let mut package = NativeOfficePackage::open_with_limits(&path, limits)
        .await
        .unwrap();

    let error = package
        .set_part("word/header1.xml", b"<header/>".to_vec())
        .unwrap_err();

    assert_eq!(error.code, "use.office.package_entry_limit");
    assert!(!package.contains_part("word/header1.xml"));
    assert!(!package.is_dirty());
}

#[tokio::test]
async fn native_package_round_trip_preserves_unknown_parts() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("round-trip.docx");
    write_package(
        &path,
        &[
            ("[Content_Types].xml", CONTENT_TYPES),
            ("_rels/.rels", ROOT_RELATIONSHIPS),
            ("word/document.xml", b"<document>before</document>"),
            ("customXml/item1.xml", b"<custom>preserve me</custom>"),
        ],
    );
    let mut package = NativeOfficePackage::open(&path).await.unwrap();
    let original_revision = package.source_revision().clone();
    package
        .set_part("/word/document.xml", b"<document>after</document>".to_vec())
        .unwrap();
    assert!(package.is_dirty());

    package.save().await.unwrap();

    assert!(!package.is_dirty());
    assert_ne!(package.source_revision(), &original_revision);
    let reopened = NativeOfficePackage::open(&path).await.unwrap();
    assert_eq!(
        reopened.part("word/document.xml").unwrap(),
        b"<document>after</document>"
    );
    assert_eq!(
        reopened.part("customXml/item1.xml").unwrap(),
        b"<custom>preserve me</custom>"
    );
}

#[tokio::test]
async fn native_package_refuses_to_overwrite_a_concurrent_revision() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("conflict.docx");
    write_package(
        &path,
        &[
            ("[Content_Types].xml", CONTENT_TYPES),
            ("_rels/.rels", ROOT_RELATIONSHIPS),
            ("word/document.xml", b"<document>original</document>"),
        ],
    );
    let mut package = NativeOfficePackage::open(&path).await.unwrap();
    package
        .set_part(
            "word/document.xml",
            b"<document>local change</document>".to_vec(),
        )
        .unwrap();

    write_package(
        &path,
        &[
            ("[Content_Types].xml", CONTENT_TYPES),
            ("_rels/.rels", ROOT_RELATIONSHIPS),
            ("word/document.xml", b"<document>external change</document>"),
        ],
    );
    let external_bytes = std::fs::read(&path).unwrap();

    let error = package.save().await.unwrap_err();

    assert_eq!(error.code, "use.office.save_conflict");
    assert_eq!(std::fs::read(&path).unwrap(), external_bytes);
    assert!(package.is_dirty());
}

#[tokio::test]
async fn native_package_save_as_requires_the_same_document_kind() {
    let temp = fixture(DocumentKind::Word);
    let source = temp.path().join("document.docx");
    let mut package = NativeOfficePackage::open(&source).await.unwrap();

    let error = package
        .save_as(temp.path().join("document.xlsx"))
        .await
        .unwrap_err();

    assert_eq!(error.code, "use.office.package_kind_mismatch");
}

#[test]
fn native_package_is_send_and_sync() {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<NativeOfficePackage>();
}

fn fixture(kind: DocumentKind) -> TempDir {
    let temp = tempfile::tempdir().unwrap();
    let (extension, main_part) = match kind {
        DocumentKind::Word => ("docx", "word/document.xml"),
        DocumentKind::Spreadsheet => ("xlsx", "xl/workbook.xml"),
        DocumentKind::Presentation => ("pptx", "ppt/presentation.xml"),
    };
    write_package(
        &temp.path().join(format!("document.{extension}")),
        &[
            ("[Content_Types].xml", CONTENT_TYPES),
            ("_rels/.rels", ROOT_RELATIONSHIPS),
            (main_part, b"<document/>"),
        ],
    );
    temp
}

fn write_package(path: &Path, entries: &[(&str, &[u8])]) {
    let file = std::fs::File::create(path).unwrap();
    let mut writer = ZipWriter::new(file);
    let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
    for (name, bytes) in entries {
        if name.ends_with('/') {
            writer.add_directory(*name, options).unwrap();
            continue;
        }
        writer.start_file(*name, options).unwrap();
        writer.write_all(bytes).unwrap();
    }
    writer.finish().unwrap();
}
