use std::path::Path;

use crate::{
    NativeOfficeDocument, NativeOfficeEditor, NativeOfficeReplayArtifact, NativeOfficeRgbColor,
    NativeSpreadsheetConditionalFormat, NativeSpreadsheetConditionalFormatOperator,
    NativeSpreadsheetConditionalFormatRule, NativeSpreadsheetConditionalFormatThreshold,
    NativeSpreadsheetDataValidation, NativeSpreadsheetDataValidationType,
    NativeSpreadsheetDifferentialFormat, NativeSpreadsheetNamedRange,
    NativeSpreadsheetNamedRangeScope, NativeSpreadsheetSort, NativeSpreadsheetSortKey,
    SpreadsheetCellValue, NATIVE_OFFICE_REPLAY_FORMAT, NATIVE_OFFICE_REPLAY_SCHEMA_VERSION,
};

async fn assert_exact_replay(
    source: &NativeOfficeEditor,
    artifact: &NativeOfficeReplayArtifact,
    target_path: &Path,
) {
    let expected_sha256 = source.package().content_sha256();
    let expected_root = source.snapshot().unwrap().root().clone();
    let mut target = NativeOfficeEditor::create(target_path).await.unwrap();
    let result = target.apply_replay(artifact).unwrap();
    assert_eq!(result.applied, artifact.mutations.len());
    assert_eq!(target.package().content_sha256(), expected_sha256);
    target.save().await.unwrap();

    let reopened = NativeOfficeDocument::open(target_path).await.unwrap();
    assert_eq!(reopened.package().content_sha256(), expected_sha256);
    assert_eq!(reopened.root(), &expected_root);
}

#[tokio::test]
async fn word_dump_replays_plain_paragraphs_and_tables_exactly() {
    let temp = tempfile::tempdir().unwrap();
    let source_path = temp.path().join("source.docx");
    let target_path = temp.path().join("target.docx");
    let mut source = NativeOfficeEditor::create(&source_path).await.unwrap();
    source.set_text("/body/p[1]", "Overview").unwrap();
    source.add_paragraph("/body", "Details").unwrap();
    source.add_table("/body", 2, 2).unwrap();
    source.set_text("/body/tbl[1]/tr[1]/tc[1]", "A").unwrap();
    source.set_text("/body/tbl[1]/tr[2]/tc[2]", "D").unwrap();

    let artifact = NativeOfficeReplayArtifact::dump(&source.snapshot().unwrap(), "/").unwrap();
    assert_eq!(artifact.format, NATIVE_OFFICE_REPLAY_FORMAT);
    assert_eq!(artifact.schema_version, NATIVE_OFFICE_REPLAY_SCHEMA_VERSION);
    assert_eq!(artifact.result_sha256, source.package().content_sha256());
    let decoded: NativeOfficeReplayArtifact =
        serde_json::from_slice(&serde_json::to_vec(&artifact).unwrap()).unwrap();
    assert_eq!(decoded, artifact);

    assert_exact_replay(&source, &artifact, &target_path).await;
}

#[tokio::test]
async fn spreadsheet_dump_replays_sheets_and_typed_cells_exactly() {
    let temp = tempfile::tempdir().unwrap();
    let source_path = temp.path().join("source.xlsx");
    let target_path = temp.path().join("target.xlsx");
    let mut source = NativeOfficeEditor::create(&source_path).await.unwrap();
    source.rename_worksheet("/Sheet1", "Data").unwrap();
    source.add_worksheet("Summary").unwrap();
    source
        .set_cell_value(
            "/Data/A1",
            SpreadsheetCellValue::Text {
                value: "Revenue".into(),
            },
        )
        .unwrap();
    source
        .set_cell_value(
            "/Data/B2",
            SpreadsheetCellValue::Number {
                value: "42.5".into(),
            },
        )
        .unwrap();
    source
        .set_cell_value("/Summary/A1", SpreadsheetCellValue::Boolean { value: true })
        .unwrap();
    source
        .set_cell_value(
            "/Summary/B1",
            SpreadsheetCellValue::Formula {
                expression: "Data!B2*2".into(),
            },
        )
        .unwrap();
    source.merge_cells("/Data/A1:B2").unwrap();
    source
        .add_data_validation(
            "/Summary",
            NativeSpreadsheetDataValidation::new(
                NativeSpreadsheetDataValidationType::List,
                "C1:C10",
                "Ready,Blocked",
            ),
        )
        .unwrap();
    source
        .add_conditional_format(
            "/Summary",
            NativeSpreadsheetConditionalFormat::new(
                "D1:D10",
                NativeSpreadsheetConditionalFormatRule::CellIs {
                    operator: NativeSpreadsheetConditionalFormatOperator::GreaterThan,
                    formula1: "50".into(),
                    formula2: None,
                    format: NativeSpreadsheetDifferentialFormat::default()
                        .with_fill(NativeOfficeRgbColor::new(198, 239, 206)),
                },
            ),
        )
        .unwrap();
    source
        .add_conditional_format(
            "/Summary",
            NativeSpreadsheetConditionalFormat::new(
                "E1:E10",
                NativeSpreadsheetConditionalFormatRule::DataBar {
                    color: NativeOfficeRgbColor::new(99, 142, 198),
                    min: NativeSpreadsheetConditionalFormatThreshold::min(),
                    max: NativeSpreadsheetConditionalFormatThreshold::max(),
                    show_value: true,
                    min_length: None,
                    max_length: None,
                },
            ),
        )
        .unwrap();
    source
        .add_named_range(
            NativeSpreadsheetNamedRange::new("Revenue", "'Data'!$B$2")
                .with_comment("Workbook total"),
        )
        .unwrap();
    source
        .add_named_range(
            NativeSpreadsheetNamedRange::new("Status", "C1:C10")
                .with_scope(NativeSpreadsheetNamedRangeScope::worksheet("Summary"))
                .with_volatile(true),
        )
        .unwrap();

    let artifact = NativeOfficeReplayArtifact::dump(&source.snapshot().unwrap(), "/").unwrap();
    assert!(artifact.mutations.iter().any(|mutation| matches!(
        mutation,
        crate::NativeOfficeMutation::MergeCells { path } if path == "/Data/A1:B2"
    )));
    assert!(artifact.mutations.iter().any(|mutation| matches!(
        mutation,
        crate::NativeOfficeMutation::AddDataValidation { sheet, .. } if sheet == "/Summary"
    )));
    assert!(artifact.mutations.iter().any(|mutation| matches!(
        mutation,
        crate::NativeOfficeMutation::AddConditionalFormat { sheet, .. } if sheet == "/Summary"
    )));
    assert!(artifact.mutations.iter().any(|mutation| matches!(
        mutation,
        crate::NativeOfficeMutation::AddNamedRange { named_range }
            if named_range.name == "Status"
                && named_range.scope
                    == NativeSpreadsheetNamedRangeScope::worksheet("Summary")
    )));
    assert_exact_replay(&source, &artifact, &target_path).await;
}

#[tokio::test]
async fn spreadsheet_dump_replays_physical_sort_and_sort_state_exactly() {
    let temp = tempfile::tempdir().unwrap();
    let source_path = temp.path().join("sorted-source.xlsx");
    let target_path = temp.path().join("sorted-target.xlsx");
    let mut source = NativeOfficeEditor::create(&source_path).await.unwrap();
    for (path, value) in [
        ("/Sheet1/A1", "Name"),
        ("/Sheet1/B1", "Rank"),
        ("/Sheet1/A2", "Beta"),
        ("/Sheet1/A3", "Alpha"),
    ] {
        source
            .set_cell_value(
                path,
                SpreadsheetCellValue::Text {
                    value: value.into(),
                },
            )
            .unwrap();
    }
    source
        .set_cell_value(
            "/Sheet1/B2",
            SpreadsheetCellValue::Number { value: "2".into() },
        )
        .unwrap();
    source
        .set_cell_value(
            "/Sheet1/B3",
            SpreadsheetCellValue::Number { value: "1".into() },
        )
        .unwrap();
    source
        .sort_spreadsheet_range(
            "/Sheet1/A1:B3",
            NativeSpreadsheetSort::new(vec![NativeSpreadsheetSortKey::ascending("B")])
                .with_header(true),
        )
        .unwrap();

    let artifact = NativeOfficeReplayArtifact::dump(&source.snapshot().unwrap(), "/").unwrap();
    assert!(artifact.mutations.iter().any(|mutation| matches!(
        mutation,
        crate::NativeOfficeMutation::SortSpreadsheetRange { path, sort }
            if path == "/Sheet1/A1:B3" && sort.header
    )));
    assert_exact_replay(&source, &artifact, &target_path).await;
}

#[tokio::test]
async fn presentation_dump_replays_slides_and_text_shapes_exactly() {
    let temp = tempfile::tempdir().unwrap();
    let source_path = temp.path().join("source.pptx");
    let target_path = temp.path().join("target.pptx");
    let mut source = NativeOfficeEditor::create(&source_path).await.unwrap();
    source.add_slide("/", "Quarterly review").unwrap();
    source.add_shape("/slide[1]", "Revenue increased").unwrap();
    source.add_table("/slide[1]", 2, 2).unwrap();
    source
        .set_text("/slide[1]/table[1]/tr[1]/tc[1]", "Metric")
        .unwrap();
    source
        .set_text("/slide[1]/table[1]/tr[1]/tc[2]", "Value")
        .unwrap();
    source
        .set_text("/slide[1]/table[1]/tr[2]/tc[1]", "Revenue")
        .unwrap();
    source
        .set_text("/slide[1]/table[1]/tr[2]/tc[2]", "42")
        .unwrap();
    source.add_slide("/", "").unwrap();
    source.add_shape("/slide[2]", "Appendix").unwrap();

    let artifact = NativeOfficeReplayArtifact::dump(&source.snapshot().unwrap(), "/").unwrap();
    assert_exact_replay(&source, &artifact, &target_path).await;
}

#[tokio::test]
async fn dump_fails_closed_and_replay_rolls_back_failed_postconditions() {
    let temp = tempfile::tempdir().unwrap();
    let rich_path = temp.path().join("rich.docx");
    let mut rich = NativeOfficeEditor::create(&rich_path).await.unwrap();
    rich.replace_xml_part(
        "/word/document.xml",
        r#"<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Rich</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#,
    )
    .unwrap();
    let error = NativeOfficeReplayArtifact::dump(&rich.snapshot().unwrap(), "/").unwrap_err();
    assert_eq!(error.code, "use.office.dump_unsupported");

    let source_path = temp.path().join("source.docx");
    let mut source = NativeOfficeEditor::create(&source_path).await.unwrap();
    source.set_text("/body/p[1]", "Expected").unwrap();
    let artifact = NativeOfficeReplayArtifact::dump(&source.snapshot().unwrap(), "/").unwrap();

    let target_path = temp.path().join("target.docx");
    let mut target = NativeOfficeEditor::create(&target_path).await.unwrap();
    let before = target.package().content_sha256();
    let mut tampered = artifact.clone();
    tampered.result_sha256 = "0".repeat(64);
    let error = target.apply_replay(&tampered).unwrap_err();
    assert_eq!(error.code, "use.office.replay_result_mismatch");
    assert_eq!(target.package().content_sha256(), before);
    assert!(!target.is_dirty());

    target.set_text("/body/p[1]", "Occupied").unwrap();
    let occupied = target.package().content_sha256();
    let error = target.apply_replay(&artifact).unwrap_err();
    assert_eq!(error.code, "use.office.replay_base_mismatch");
    assert_eq!(target.package().content_sha256(), occupied);
}
