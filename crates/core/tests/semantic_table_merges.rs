use std::io::Write;
use std::path::{Path, PathBuf};

use a3s_office::{NativeOfficeDocument, NativeOfficeReplayArtifact};
use tempfile::TempDir;
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipWriter};

const CONTENT_TYPES_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

#[tokio::test]
async fn word_semantics_resolve_horizontal_and_vertical_merges() {
    let (_fixture, path) = word_fixture();
    let document = NativeOfficeDocument::open(&path).await.unwrap();

    let anchor = document.get("/body/tbl[1]/tr[1]/tc[1]", 0).unwrap();
    assert_grid_cell(&anchor, 1, 1, 2, 2, true);

    let first_row_tail = document.get("/body/tbl[1]/tr[1]/tc[2]", 0).unwrap();
    assert_grid_cell(&first_row_tail, 1, 3, 1, 1, true);

    let continuation = document.get("/body/tbl[1]/tr[2]/tc[1]", 0).unwrap();
    assert_grid_cell(&continuation, 2, 1, 1, 2, false);
    assert_eq!(
        continuation
            .format
            .get("mergeAnchorPath")
            .map(String::as_str),
        Some("/body/tbl[1]/tr[1]/tc[1]")
    );

    let after_merge = document.get("/body/tbl[1]/tr[3]/tc[1]", 0).unwrap();
    assert_grid_cell(&after_merge, 3, 1, 1, 1, true);

    let replay_error = NativeOfficeReplayArtifact::dump(&document, "/").unwrap_err();
    assert_eq!(replay_error.code, "use.office.dump_unsupported");
}

#[tokio::test]
async fn presentation_semantics_resolve_rectangular_merge_anchors_and_covered_cells() {
    let (_fixture, path) = presentation_fixture();
    let document = NativeOfficeDocument::open(&path).await.unwrap();

    let anchor = document.get("/slide[1]/table[1]/tr[1]/tc[1]", 0).unwrap();
    assert_grid_cell(&anchor, 1, 1, 2, 2, true);

    for (path, row, column) in [
        ("/slide[1]/table[1]/tr[1]/tc[2]", 1, 2),
        ("/slide[1]/table[1]/tr[2]/tc[1]", 2, 1),
        ("/slide[1]/table[1]/tr[2]/tc[2]", 2, 2),
    ] {
        let covered = document.get(path, 0).unwrap();
        assert_grid_cell(&covered, row, column, 1, 1, false);
        assert_eq!(
            covered.format.get("mergeAnchorPath").map(String::as_str),
            Some("/slide[1]/table[1]/tr[1]/tc[1]")
        );
    }

    let tail = document.get("/slide[1]/table[1]/tr[2]/tc[3]", 0).unwrap();
    assert_grid_cell(&tail, 2, 3, 1, 1, true);

    let replay_error = NativeOfficeReplayArtifact::dump(&document, "/").unwrap_err();
    assert_eq!(replay_error.code, "use.office.dump_unsupported");
}

#[tokio::test]
async fn word_semantics_reject_orphan_vertical_merge_continuations() {
    let document = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tblGrid><w:gridCol/></w:tblGrid><w:tr><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>"#;
    let (_fixture, path) = word_fixture_with_document("orphan-vmerge.docx", document);

    let error = NativeOfficeDocument::open(&path).await.unwrap_err();

    assert_eq!(error.code, "use.office.word_table_grid_invalid");
}

#[tokio::test]
async fn word_semantics_reject_invalid_grid_spans() {
    for (file_name, span) in [
        ("zero-grid-span.docx", "0"),
        ("oversized-grid-span.docx", "1000001"),
    ] {
        let document = format!(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tr><w:tc><w:tcPr><w:gridSpan w:val="{span}"/></w:tcPr><w:p/></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>"#
        );
        let (_fixture, path) = word_fixture_with_document(file_name, &document);

        let error = NativeOfficeDocument::open(&path).await.unwrap_err();

        assert_eq!(error.code, "use.office.word_table_grid_invalid", "{span}");
    }
}

#[tokio::test]
async fn word_semantics_honor_grid_before_for_logical_columns() {
    let document = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tblGrid><w:gridCol/><w:gridCol/><w:gridCol/></w:tblGrid><w:tr><w:trPr><w:gridBefore w:val="1"/></w:trPr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>Offset</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>"#;
    let (_fixture, path) = word_fixture_with_document("grid-before.docx", document);

    let document = NativeOfficeDocument::open(&path).await.unwrap();
    let cell = document.get("/body/tbl[1]/tr[1]/tc[1]", 0).unwrap();

    assert_grid_cell(&cell, 1, 2, 1, 2, true);
}

#[tokio::test]
async fn presentation_semantics_reject_orphan_covered_cells() {
    let slide = presentation_slide(
        r#"<a:tblGrid><a:gridCol w="100"/></a:tblGrid><a:tr h="100"><a:tc hMerge="1"><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr/></a:tc></a:tr>"#,
    );
    let (_fixture, path) = presentation_fixture_with_slide("orphan-hmerge.pptx", &slide);

    let error = NativeOfficeDocument::open(&path).await.unwrap_err();

    assert_eq!(error.code, "use.office.presentation_table_grid_invalid");
}

#[tokio::test]
async fn presentation_semantics_reject_invalid_or_out_of_bounds_spans() {
    for (file_name, span) in [("zero-grid-span.pptx", "0"), ("overflow-grid.pptx", "2")] {
        let slide = presentation_slide(&format!(
            r#"<a:tblGrid><a:gridCol w="100"/></a:tblGrid><a:tr h="100"><a:tc gridSpan="{span}"><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr/></a:tc></a:tr>"#
        ));
        let (_fixture, path) = presentation_fixture_with_slide(file_name, &slide);

        let error = NativeOfficeDocument::open(&path).await.unwrap_err();

        assert_eq!(
            error.code, "use.office.presentation_table_grid_invalid",
            "{span}"
        );
    }
}

#[tokio::test]
async fn presentation_semantics_resolve_compact_merged_rows_without_covered_cells() {
    let slide = presentation_slide(
        r#"<a:tblGrid><a:gridCol w="100"/><a:gridCol w="100"/><a:gridCol w="100"/></a:tblGrid><a:tr h="100"><a:tc gridSpan="2" rowSpan="2"><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Merged</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>R1C3</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr><a:tr h="100"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>R2C3</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>"#,
    );
    let (_fixture, path) = presentation_fixture_with_slide("compact-merge.pptx", &slide);

    let document = NativeOfficeDocument::open(&path).await.unwrap();
    let anchor = document.get("/slide[1]/table[1]/tr[1]/tc[1]", 0).unwrap();
    let first_tail = document.get("/slide[1]/table[1]/tr[1]/tc[2]", 0).unwrap();
    let second_tail = document.get("/slide[1]/table[1]/tr[2]/tc[1]", 0).unwrap();

    assert_grid_cell(&anchor, 1, 1, 2, 2, true);
    assert_grid_cell(&first_tail, 1, 3, 1, 1, true);
    assert_grid_cell(&second_tail, 2, 3, 1, 1, true);
}

fn assert_grid_cell(
    cell: &a3s_office::DocumentNode,
    row: u32,
    column: u32,
    row_span: u32,
    column_span: u32,
    merge_anchor: bool,
) {
    assert_eq!(cell.format.get("row"), Some(&row.to_string()));
    assert_eq!(cell.format.get("column"), Some(&column.to_string()));
    assert_eq!(cell.format.get("rowSpan"), Some(&row_span.to_string()));
    assert_eq!(
        cell.format.get("columnSpan"),
        Some(&column_span.to_string())
    );
    assert_eq!(
        cell.format.get("mergeAnchor"),
        Some(&merge_anchor.to_string())
    );
}

fn word_fixture() -> (TempDir, PathBuf) {
    let document = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:tbl><w:tblGrid><w:gridCol/><w:gridCol/><w:gridCol/></w:tblGrid><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>Merged</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>R1C3</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/><w:vMerge/></w:tcPr><w:p/></w:tc><w:tc><w:p><w:r><w:t>R2C3</w:t></w:r></w:p></w:tc></w:tr><w:tr><w:tc><w:p><w:r><w:t>R3C1</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>R3C2</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>R3C3</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>"#;
    word_fixture_with_document("merged.docx", document)
}

fn word_fixture_with_document(file_name: &str, document: &str) -> (TempDir, PathBuf) {
    let content_types = format!(
        r#"<Types xmlns="{CONTENT_TYPES_NAMESPACE}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>"#
    );
    let relationships = root_relationships("word/document.xml");
    fixture(
        file_name,
        &[
            ("[Content_Types].xml", content_types.as_bytes()),
            ("_rels/.rels", relationships.as_bytes()),
            ("word/document.xml", document.as_bytes()),
        ],
    )
}

fn presentation_fixture() -> (TempDir, PathBuf) {
    let slide = presentation_slide(
        r#"<a:tblGrid><a:gridCol w="100"/><a:gridCol w="100"/><a:gridCol w="100"/></a:tblGrid><a:tr h="100"><a:tc gridSpan="2" rowSpan="2"><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Merged</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc><a:tc hMerge="1"><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr/></a:tc><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>R1C3</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr><a:tr h="100"><a:tc vMerge="1"><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr/></a:tc><a:tc hMerge="1" vMerge="1"><a:txBody><a:bodyPr/><a:lstStyle/><a:p/></a:txBody><a:tcPr/></a:tc><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>R2C3</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr>"#,
    );
    presentation_fixture_with_slide("merged.pptx", &slide)
}

fn presentation_slide(table_contents: &str) -> String {
    format!(
        r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="2" name="Table"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>{table_contents}</a:tbl></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>"#
    )
}

fn presentation_fixture_with_slide(file_name: &str, slide: &str) -> (TempDir, PathBuf) {
    let content_types = format!(
        r#"<Types xmlns="{CONTENT_TYPES_NAMESPACE}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>"#
    );
    let relationships = root_relationships("ppt/presentation.xml");
    let presentation = r#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>"#;
    let presentation_relationships = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NAMESPACE}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>"#
    );
    fixture(
        file_name,
        &[
            ("[Content_Types].xml", content_types.as_bytes()),
            ("_rels/.rels", relationships.as_bytes()),
            ("ppt/presentation.xml", presentation.as_bytes()),
            (
                "ppt/_rels/presentation.xml.rels",
                presentation_relationships.as_bytes(),
            ),
            ("ppt/slides/slide1.xml", slide.as_bytes()),
        ],
    )
}

fn root_relationships(main_part: &str) -> String {
    format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NAMESPACE}"><Relationship Id="rId1" Type="{OFFICE_DOCUMENT_RELATIONSHIP}" Target="/{main_part}"/></Relationships>"#
    )
}

fn fixture(file_name: &str, entries: &[(&str, &[u8])]) -> (TempDir, PathBuf) {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join(file_name);
    write_package(&path, entries);
    (temp, path)
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
