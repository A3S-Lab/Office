use std::io::Write;
use std::path::Path;

use tempfile::TempDir;
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipWriter};

use super::{DocumentKind, NativeOfficePackage, RelationshipSource, RelationshipTarget};

const CONTENT_TYPES_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

#[tokio::test]
async fn opc_model_reads_content_types_and_resolves_relationships() {
    let fixture = document_fixture(DocumentKind::Word, &[], None);
    let path = fixture.path().join("document.docx");
    let package = NativeOfficePackage::open(path).await.unwrap();

    let model = package.opc_model().unwrap();

    assert_eq!(
        model.content_types().content_type("word/document.xml"),
        Some("application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml")
    );
    assert_eq!(
        model.content_types().content_type("word/media/image1.png"),
        Some("image/png")
    );
    let package_relationships = model
        .relationships()
        .relationships_from(&RelationshipSource::Package);
    assert_eq!(package_relationships.len(), 1);
    assert_eq!(
        package_relationships[0].target.internal_part_name(),
        Some("word/document.xml")
    );

    let document_source = RelationshipSource::Part {
        part_name: "word/document.xml".to_string(),
    };
    let image = model
        .relationships()
        .relationship(&document_source, "rIdImage")
        .unwrap();
    assert_eq!(
        image.target,
        RelationshipTarget::Internal {
            part_name: "word/media/image1.png".to_string(),
            fragment: None,
        }
    );
    let hyperlink = model
        .relationships()
        .relationship(&document_source, "rIdLink")
        .unwrap();
    assert_eq!(
        hyperlink.target,
        RelationshipTarget::External {
            uri: "https://example.invalid/a?b=1".to_string(),
        }
    );
}

#[tokio::test]
async fn opc_model_accepts_each_supported_main_content_type() {
    for kind in [
        DocumentKind::Word,
        DocumentKind::Spreadsheet,
        DocumentKind::Presentation,
    ] {
        let fixture = document_fixture(kind, &[], None);
        let path = fixture.path().join(format!("document.{}", extension(kind)));
        let package = NativeOfficePackage::open(path).await.unwrap();

        let model = package.opc_model().unwrap();

        assert_eq!(
            model.content_types().content_type(main_part(kind)),
            Some(main_content_type(kind))
        );
    }
}

#[tokio::test]
async fn opc_model_rejects_relationship_target_escape_and_duplicate_ids() {
    let escaping = format!(
        r#"<?xml version="1.0"?><Relationships xmlns="{RELATIONSHIPS_NAMESPACE}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT_RELATIONSHIP}" Target="../../word/document.xml"/></Relationships>"#
    );
    let fixture = document_fixture(DocumentKind::Word, &[], Some(escaping.as_bytes()));
    let package = NativeOfficePackage::open(fixture.path().join("document.docx"))
        .await
        .unwrap();
    assert_eq!(
        package.opc_model().unwrap_err().code,
        "use.office.relationship_target_escape"
    );

    let duplicate = format!(
        r#"<?xml version="1.0"?><Relationships xmlns="{RELATIONSHIPS_NAMESPACE}"><Relationship Id="same" Type="{OFFICE_DOCUMENT_RELATIONSHIP}" Target="word/document.xml"/><Relationship Id="same" Type="urn:metadata" Target="docProps/core.xml"/></Relationships>"#
    );
    let fixture = document_fixture(DocumentKind::Word, &[], Some(duplicate.as_bytes()));
    let package = NativeOfficePackage::open(fixture.path().join("document.docx"))
        .await
        .unwrap();
    assert_eq!(
        package.opc_model().unwrap_err().code,
        "use.office.relationship_id_duplicate"
    );
}

#[tokio::test]
async fn opc_model_requires_a_content_type_for_every_part() {
    let fixture = document_fixture(
        DocumentKind::Word,
        &[("custom/data.unknown", b"opaque")],
        None,
    );
    let package = NativeOfficePackage::open(fixture.path().join("document.docx"))
        .await
        .unwrap();

    let error = package.opc_model().unwrap_err();

    assert_eq!(error.code, "use.office.content_type_missing");
    assert!(error.message.contains("custom/data.unknown"));
}

#[tokio::test]
async fn opc_model_rejects_dangling_internal_relationship_targets() {
    let relationships = format!(
        r#"<?xml version="1.0"?><Relationships xmlns="{RELATIONSHIPS_NAMESPACE}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT_RELATIONSHIP}" Target="word/document.xml"/><Relationship Id="rIdMissing" Type="urn:image" Target="word/media/missing.png"/></Relationships>"#
    );
    let fixture = document_fixture(DocumentKind::Word, &[], Some(relationships.as_bytes()));
    let package = NativeOfficePackage::open(fixture.path().join("document.docx"))
        .await
        .unwrap();

    assert_eq!(
        package.opc_model().unwrap_err().code,
        "use.office.relationship_target_missing"
    );
}

#[tokio::test]
async fn opc_model_rejects_case_ambiguous_content_type_overrides() {
    let overrides = format!(
        r#"<?xml version="1.0"?><Types xmlns="{CONTENT_TYPES_NAMESPACE}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="{}"/><Override PartName="/Word/Document.xml" ContentType="{}"/></Types>"#,
        main_content_type(DocumentKind::Word),
        main_content_type(DocumentKind::Word),
    );
    let fixture = document_fixture(
        DocumentKind::Word,
        &[("[Content_Types].xml", overrides.as_bytes())],
        None,
    );
    let package = NativeOfficePackage::open(fixture.path().join("document.docx"))
        .await
        .unwrap();

    assert_eq!(
        package.opc_model().unwrap_err().code,
        "use.office.content_types_duplicate"
    );
}

fn document_fixture(
    kind: DocumentKind,
    extra_parts: &[(&str, &[u8])],
    root_relationships: Option<&[u8]>,
) -> TempDir {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join(format!("document.{}", extension(kind)));
    let content_types = content_types(kind);
    let default_root_relationships = root_relationships_xml(kind);
    let document_relationships = format!(
        r#"<?xml version="1.0"?><Relationships xmlns="{RELATIONSHIPS_NAMESPACE}"><Relationship Id="rIdImage" Type="urn:image" Target="media/image1.png"/><Relationship Id="rIdLink" Type="urn:hyperlink" Target="https://example.invalid/a?b=1" TargetMode="External"/></Relationships>"#
    );
    let document_xml = main_xml(kind);
    let relationship_part = relationship_part(kind);

    let mut entries = vec![
        ("[Content_Types].xml", content_types.as_bytes()),
        (
            "_rels/.rels",
            root_relationships.unwrap_or(default_root_relationships.as_bytes()),
        ),
        (main_part(kind), document_xml.as_bytes()),
        (relationship_part, document_relationships.as_bytes()),
        (media_part(kind), b"fake png".as_slice()),
    ];
    for (name, bytes) in extra_parts {
        if *name == "[Content_Types].xml" {
            entries[0] = (*name, *bytes);
        } else {
            entries.push((*name, *bytes));
        }
    }
    write_package(&path, &entries);
    temp
}

fn content_types(kind: DocumentKind) -> String {
    format!(
        r#"<?xml version="1.0"?><Types xmlns="{CONTENT_TYPES_NAMESPACE}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/{}" ContentType="{}"/></Types>"#,
        main_part(kind),
        main_content_type(kind),
    )
}

fn root_relationships_xml(kind: DocumentKind) -> String {
    format!(
        r#"<?xml version="1.0"?><Relationships xmlns="{RELATIONSHIPS_NAMESPACE}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT_RELATIONSHIP}" Target="/{}"/></Relationships>"#,
        main_part(kind),
    )
}

fn main_xml(kind: DocumentKind) -> String {
    match kind {
        DocumentKind::Word => concat!(
            r#"<?xml version="1.0"?><w:document xmlns:w="#,
            r#""http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body/></w:document>"#
        )
        .to_string(),
        DocumentKind::Spreadsheet => concat!(
            r#"<?xml version="1.0"?><workbook xmlns="#,
            r#""http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>"#
        )
        .to_string(),
        DocumentKind::Presentation => concat!(
            r#"<?xml version="1.0"?><p:presentation xmlns:p="#,
            r#""http://schemas.openxmlformats.org/presentationml/2006/main"/>"#
        )
        .to_string(),
    }
}

fn extension(kind: DocumentKind) -> &'static str {
    match kind {
        DocumentKind::Word => "docx",
        DocumentKind::Spreadsheet => "xlsx",
        DocumentKind::Presentation => "pptx",
    }
}

fn main_part(kind: DocumentKind) -> &'static str {
    match kind {
        DocumentKind::Word => "word/document.xml",
        DocumentKind::Spreadsheet => "xl/workbook.xml",
        DocumentKind::Presentation => "ppt/presentation.xml",
    }
}

fn relationship_part(kind: DocumentKind) -> &'static str {
    match kind {
        DocumentKind::Word => "word/_rels/document.xml.rels",
        DocumentKind::Spreadsheet => "xl/_rels/workbook.xml.rels",
        DocumentKind::Presentation => "ppt/_rels/presentation.xml.rels",
    }
}

fn media_part(kind: DocumentKind) -> &'static str {
    match kind {
        DocumentKind::Word => "word/media/image1.png",
        DocumentKind::Spreadsheet => "xl/media/image1.png",
        DocumentKind::Presentation => "ppt/media/image1.png",
    }
}

fn main_content_type(kind: DocumentKind) -> &'static str {
    match kind {
        DocumentKind::Word => {
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
        }
        DocumentKind::Spreadsheet => {
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
        }
        DocumentKind::Presentation => {
            "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
        }
    }
}

fn write_package(path: &Path, entries: &[(&str, &[u8])]) {
    let file = std::fs::File::create(path).unwrap();
    let mut writer = ZipWriter::new(file);
    let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
    for (name, bytes) in entries {
        writer.start_file(*name, options).unwrap();
        writer.write_all(bytes).unwrap();
    }
    writer.finish().unwrap();
}
