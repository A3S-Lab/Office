use std::collections::BTreeMap;

use super::{NativeOfficeEditor, NativeOfficeMutation, NativeOfficePartType, NativeRawXmlPart};
use crate::{NativeOfficePackage, RelationshipSource, RelationshipTarget};

const TRANSITIONAL_WORD: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_WORD: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const TRANSITIONAL_SPREADSHEET: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SPREADSHEET: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
const TRANSITIONAL_RELATIONSHIPS: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const STRICT_RELATIONSHIPS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";

#[tokio::test]
async fn creates_word_parts_with_typed_receipts_and_owner_relationships() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("parts.docx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();

    let header = editor.add_part("/", NativeOfficePartType::Header).unwrap();
    assert_eq!(header.path, "/header[1]");
    assert_eq!(header.part, "/word/header1.xml");
    assert_eq!(header.owner_part, "/word/document.xml");
    assert_eq!(header.part_type, NativeOfficePartType::Header);
    assert_eq!(
        editor.raw_xml_part(&header.part).unwrap().root.local_name,
        "hdr"
    );
    assert_eq!(
        editor
            .snapshot()
            .unwrap()
            .get("/header[1]", 0)
            .unwrap()
            .path,
        "/header[1]"
    );

    let footer = editor.add_part("/", NativeOfficePartType::Footer).unwrap();
    assert_eq!(footer.path, "/footer[1]");
    assert_eq!(footer.part, "/word/footer1.xml");

    let chart = editor.add_part("/", NativeOfficePartType::Chart).unwrap();
    assert_eq!(chart.path, "/chart[1]");
    assert_eq!(chart.part, "/word/charts/chart1.xml");
    assert_eq!(
        editor.raw_xml_part(&chart.part).unwrap().root.local_name,
        "chartSpace"
    );
    assert_relationship(
        &editor,
        "word/document.xml",
        &chart.relationship_id,
        "word/charts/chart1.xml",
    );
}

#[tokio::test]
async fn creates_presentation_and_spreadsheet_chart_carriers() {
    let temp = tempfile::tempdir().unwrap();
    let presentation_path = temp.path().join("parts.pptx");
    let mut presentation = NativeOfficeEditor::create(&presentation_path)
        .await
        .unwrap();
    presentation.add_slide("/", "Chart carrier").unwrap();
    let chart = presentation
        .add_part("/slide[1]", NativeOfficePartType::Chart)
        .unwrap();
    assert_eq!(chart.path, "/slide[1]/chart[1]");
    assert_eq!(chart.part, "/ppt/charts/chart1.xml");
    assert_eq!(chart.owner_part, "/ppt/slides/slide1.xml");
    assert_relationship(
        &presentation,
        "ppt/slides/slide1.xml",
        &chart.relationship_id,
        "ppt/charts/chart1.xml",
    );

    let spreadsheet_path = temp.path().join("parts.xlsx");
    let mut spreadsheet = NativeOfficeEditor::create(&spreadsheet_path).await.unwrap();
    let first = spreadsheet
        .add_part("/Sheet1", NativeOfficePartType::Chart)
        .unwrap();
    assert_eq!(first.path, "/Sheet1/chart[1]");
    assert_eq!(first.part, "/xl/charts/chart1.xml");
    assert_eq!(first.owner_part, "/xl/drawings/drawing1.xml");
    assert!(part_text(&spreadsheet, "xl/worksheets/sheet1.xml").contains("<drawing"));
    assert_relationship(
        &spreadsheet,
        "xl/drawings/drawing1.xml",
        &first.relationship_id,
        "xl/charts/chart1.xml",
    );

    let second = spreadsheet
        .add_part("/Sheet1", NativeOfficePartType::Chart)
        .unwrap();
    assert_eq!(second.path, "/Sheet1/chart[2]");
    assert_eq!(second.part, "/xl/charts/chart2.xml");
    assert_eq!(second.owner_part, first.owner_part);
    assert!(!spreadsheet
        .package()
        .contains_part("xl/drawings/drawing2.xml"));
    spreadsheet.snapshot().unwrap();
}

#[tokio::test]
async fn typed_part_creation_reports_batch_receipts_and_rolls_back_atomically() {
    let temp = tempfile::tempdir().unwrap();
    let word_path = temp.path().join("batch.docx");
    let mut word = NativeOfficeEditor::create(&word_path).await.unwrap();
    let result = word
        .apply_batch(&[
            NativeOfficeMutation::AddPart {
                parent: "/".to_string(),
                part_type: NativeOfficePartType::Header,
            },
            NativeOfficeMutation::AddPart {
                parent: "/".to_string(),
                part_type: NativeOfficePartType::Chart,
            },
        ])
        .unwrap();
    assert_eq!(result.paths, ["/header[1]", "/chart[1]"]);
    assert_eq!(result.created_parts.len(), 2);
    assert_eq!(result.created_parts[1].relationship_id, "rId3");

    let spreadsheet_path = temp.path().join("rollback.xlsx");
    let mut spreadsheet = NativeOfficeEditor::create(&spreadsheet_path).await.unwrap();
    let original = package_parts(&spreadsheet);
    let error = spreadsheet
        .apply_batch(&[
            NativeOfficeMutation::AddPart {
                parent: "/Sheet1".to_string(),
                part_type: NativeOfficePartType::Chart,
            },
            NativeOfficeMutation::AddPart {
                parent: "/Sheet1".to_string(),
                part_type: NativeOfficePartType::Header,
            },
        ])
        .unwrap_err();
    assert_eq!(error.code, "use.office.part_type_unsupported");
    assert_eq!(package_parts(&spreadsheet), original);
    assert!(!spreadsheet.is_dirty());
}

#[tokio::test]
async fn typed_parts_follow_the_strict_ooxml_dialect() {
    let temp = tempfile::tempdir().unwrap();
    let word_path = temp.path().join("strict.docx");
    let mut word_package = NativeOfficePackage::create(&word_path).await.unwrap();
    let strict_document =
        String::from_utf8(word_package.part("word/document.xml").unwrap().to_vec())
            .unwrap()
            .replace(TRANSITIONAL_WORD, STRICT_WORD);
    word_package
        .set_part("word/document.xml", strict_document.into_bytes())
        .unwrap();
    let mut word = NativeOfficeEditor::from_package(word_package).unwrap();
    let header = word.add_part("/", NativeOfficePartType::Header).unwrap();
    assert_eq!(
        word.raw_xml_part(&header.part)
            .unwrap()
            .root
            .namespace
            .as_deref(),
        Some(STRICT_WORD)
    );
    assert!(
        relationship_type(&word, "word/document.xml", &header.relationship_id)
            .starts_with(STRICT_RELATIONSHIPS)
    );

    let spreadsheet_path = temp.path().join("strict.xlsx");
    let mut spreadsheet_package = NativeOfficePackage::create(&spreadsheet_path)
        .await
        .unwrap();
    for part in ["xl/workbook.xml", "xl/worksheets/sheet1.xml"] {
        let strict = String::from_utf8(spreadsheet_package.part(part).unwrap().to_vec())
            .unwrap()
            .replace(TRANSITIONAL_SPREADSHEET, STRICT_SPREADSHEET)
            .replace(TRANSITIONAL_RELATIONSHIPS, STRICT_RELATIONSHIPS);
        spreadsheet_package
            .set_part(part, strict.into_bytes())
            .unwrap();
    }
    let mut spreadsheet = NativeOfficeEditor::from_package(spreadsheet_package).unwrap();
    let chart = spreadsheet
        .add_part("/Sheet1", NativeOfficePartType::Chart)
        .unwrap();
    assert!(spreadsheet
        .raw_xml_part(&chart.owner_part)
        .unwrap()
        .root
        .namespace
        .as_deref()
        .unwrap()
        .contains("purl.oclc.org"));
    assert!(relationship_type(
        &spreadsheet,
        chart.owner_part.trim_start_matches('/'),
        &chart.relationship_id
    )
    .starts_with(STRICT_RELATIONSHIPS));
}

#[test]
fn created_part_contracts_are_send_and_sync() {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<super::NativeCreatedPart>();
    assert_send_sync::<NativeOfficePartType>();
    assert_send_sync::<NativeRawXmlPart>();
}

fn assert_relationship(editor: &NativeOfficeEditor, owner: &str, id: &str, expected_target: &str) {
    let source = RelationshipSource::Part {
        part_name: owner.to_string(),
    };
    let model = editor.package().opc_model().unwrap();
    let relationship = model.relationships().relationship(&source, id).unwrap();
    let RelationshipTarget::Internal { part_name, .. } = &relationship.target else {
        panic!("expected internal relationship target");
    };
    assert_eq!(part_name, expected_target);
}

fn relationship_type(editor: &NativeOfficeEditor, owner: &str, id: &str) -> String {
    let source = RelationshipSource::Part {
        part_name: owner.to_string(),
    };
    editor
        .package()
        .opc_model()
        .unwrap()
        .relationships()
        .relationship(&source, id)
        .unwrap()
        .relationship_type
        .clone()
}

fn part_text(editor: &NativeOfficeEditor, part: &str) -> String {
    String::from_utf8(editor.package().part(part).unwrap().to_vec()).unwrap()
}

fn package_parts(editor: &NativeOfficeEditor) -> BTreeMap<String, Vec<u8>> {
    editor
        .package()
        .part_names()
        .map(|part| {
            (
                part.to_string(),
                editor.package().part(part).unwrap().to_vec(),
            )
        })
        .collect()
}
