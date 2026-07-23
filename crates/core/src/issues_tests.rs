use crate::{
    NativeOfficeDocument, NativeOfficeEditor, NativeOfficeImage, NativeOfficeIssueCategory,
    NativeOfficeIssueFilter, NativeOfficeIssueOptions, NativeOfficeIssueReport,
    NativeOfficeIssueSeverity, NativeOfficeIssueSubtype, NativeOfficePackage, SpreadsheetCellValue,
};

const PNG_1X1: &[u8] = &[
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
    0x89, 0x00, 0x00, 0x00, 0x0a, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
    0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
    0x42, 0x60, 0x82,
];

#[tokio::test]
async fn native_issue_view_detects_accessibility_and_formula_issues() {
    let temp = tempfile::tempdir().unwrap();

    let mut word = NativeOfficeEditor::create(temp.path().join("report.docx"))
        .await
        .unwrap();
    let picture = word
        .add_image(
            "/body",
            NativeOfficeImage::from_bytes(PNG_1X1)
                .unwrap()
                .with_name("Quarterly chart"),
        )
        .unwrap();
    let report = word.snapshot().unwrap().issue_view().unwrap();
    assert_eq!(report.count, 1);
    assert_eq!(report.issues[0].path, picture.path);
    assert_eq!(
        report.issues[0].subtype,
        NativeOfficeIssueSubtype::MissingAltText
    );
    assert_eq!(
        report.issues[0].severity,
        NativeOfficeIssueSeverity::Warning
    );

    let mut spreadsheet = NativeOfficeEditor::create(temp.path().join("book.xlsx"))
        .await
        .unwrap();
    spreadsheet
        .set_cell_value(
            "/Sheet1/A1",
            SpreadsheetCellValue::Formula {
                expression: "SUM(B1:B2)".to_string(),
            },
        )
        .unwrap();
    spreadsheet
        .set_cell_value(
            "/Sheet1/A2",
            SpreadsheetCellValue::Formula {
                expression: "Missing!A1".to_string(),
            },
        )
        .unwrap();
    spreadsheet
        .set_cell_value(
            "/Sheet1/A3",
            SpreadsheetCellValue::Formula {
                expression: "#REF!+1".to_string(),
            },
        )
        .unwrap();
    let report = spreadsheet.snapshot().unwrap().issue_view().unwrap();
    assert_eq!(report.count, 3);
    assert_eq!(
        report.issues[0].subtype,
        NativeOfficeIssueSubtype::FormulaNotEvaluated
    );
    assert_eq!(
        report.issues[1].subtype,
        NativeOfficeIssueSubtype::FormulaRefMissingSheet
    );
    assert_eq!(
        report.issues[2].subtype,
        NativeOfficeIssueSubtype::FormulaEvalError
    );

    let filtered = spreadsheet
        .snapshot()
        .unwrap()
        .issues(NativeOfficeIssueOptions {
            filter: Some(NativeOfficeIssueFilter::FormulaRefMissingSheet),
            limit: 10,
        })
        .unwrap();
    assert_eq!(filtered.count, 1);
    assert_eq!(filtered.returned, 1);
    assert_eq!(filtered.issues[0].path, "/Sheet1/A2");
}

#[tokio::test]
async fn native_issue_view_detects_formula_errors_from_cached_values() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("errors.xlsx");
    let mut package = NativeOfficePackage::create(&path).await.unwrap();
    package
        .set_part(
            "xl/worksheets/sheet1.xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="e"><f>1/0</f><v>#DIV/0!</v></c></row></sheetData></worksheet>"#
                .to_vec(),
        )
        .unwrap();
    let document = NativeOfficeDocument::from_package(package).unwrap();
    let report = document.issue_view().unwrap();

    assert_eq!(report.count, 1);
    assert_eq!(
        report.issues[0].subtype,
        NativeOfficeIssueSubtype::FormulaEvalError
    );
    assert_eq!(report.issues[0].severity, NativeOfficeIssueSeverity::Error);
}

#[tokio::test]
async fn native_issue_view_detects_broken_references_and_explicit_low_contrast() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("deck.pptx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor.add_slide("/", "Fixture").unwrap();
    let mut package = editor.package().clone();
    package
        .set_part(
            "ppt/slides/slide1.xml",
            br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Low contrast"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:solidFill><a:srgbClr val="111111"/></a:solidFill></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr><a:solidFill><a:srgbClr val="222222"/></a:solidFill></a:rPr><a:t>Unreadable text</a:t></a:r></a:p></p:txBody></p:sp><p:pic><p:nvPicPr><p:cNvPr id="3" name="Broken picture"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId999"/></p:blipFill><p:spPr/></p:pic></p:spTree></p:cSld></p:sld>"#
                .to_vec(),
        )
        .unwrap();
    let document = NativeOfficeDocument::from_package(package).unwrap();
    let report = document.issue_view().unwrap();

    assert_eq!(report.count, 3);
    assert!(report
        .issues
        .iter()
        .any(|issue| issue.subtype == NativeOfficeIssueSubtype::LowContrast));
    assert!(report
        .issues
        .iter()
        .any(|issue| issue.subtype == NativeOfficeIssueSubtype::MissingAltText));
    assert!(report
        .issues
        .iter()
        .any(|issue| issue.subtype == NativeOfficeIssueSubtype::BrokenPartRef));

    let structure = document
        .issues(NativeOfficeIssueOptions {
            filter: Some(NativeOfficeIssueFilter::Structure),
            limit: 10,
        })
        .unwrap();
    assert_eq!(structure.count, 1);
    assert_eq!(structure.returned, 1);
    assert_eq!(
        structure.issues[0].subtype,
        NativeOfficeIssueSubtype::BrokenPartRef
    );
}

#[tokio::test]
async fn native_issue_view_parses_sheet_references_without_scanning_string_literals() {
    let temp = tempfile::tempdir().unwrap();
    let mut spreadsheet = NativeOfficeEditor::create(temp.path().join("references.xlsx"))
        .await
        .unwrap();
    spreadsheet.add_worksheet("Q1 Data").unwrap();
    for (path, formula) in [
        ("/Sheet1/A1", "'Q1 Data'!A1"),
        ("/Sheet1/A2", "SUM(Sheet1!B1:B2)"),
        ("/Sheet1/A3", "=\"Missing!A1\""),
        ("/Sheet1/A4", "=\"#REF!\""),
        ("/Sheet1/A5", "'Missing Data'!A1"),
        ("/Sheet1/A6", "#DIV/0!+1"),
        ("/Sheet1/A7", "[Book.xlsx]Missing!A1"),
        ("/Sheet1/A8", "'[Book.xlsx]Missing Data'!A1"),
    ] {
        spreadsheet
            .set_cell_value(
                path,
                SpreadsheetCellValue::Formula {
                    expression: formula.to_string(),
                },
            )
            .unwrap();
    }

    let report = spreadsheet.snapshot().unwrap().issue_view().unwrap();
    assert_eq!(report.count, 8);
    assert_eq!(
        report
            .issues
            .iter()
            .filter(|issue| issue.subtype == NativeOfficeIssueSubtype::FormulaRefMissingSheet)
            .count(),
        1
    );
    assert_eq!(
        report
            .issues
            .iter()
            .find(|issue| issue.subtype == NativeOfficeIssueSubtype::FormulaRefMissingSheet)
            .unwrap()
            .path,
        "/Sheet1/A5"
    );
    assert_eq!(
        report
            .issues
            .iter()
            .find(|issue| issue.path == "/Sheet1/A6")
            .unwrap()
            .subtype,
        NativeOfficeIssueSubtype::FormulaEvalError
    );
    for path in ["/Sheet1/A7", "/Sheet1/A8"] {
        assert_eq!(
            report
                .issues
                .iter()
                .find(|issue| issue.path == path)
                .unwrap()
                .subtype,
            NativeOfficeIssueSubtype::FormulaNotEvaluated
        );
    }
}

#[tokio::test]
async fn native_issue_view_is_bounded_and_blank_documents_are_clean() {
    let temp = tempfile::tempdir().unwrap();
    for extension in ["docx", "xlsx", "pptx"] {
        let document = NativeOfficeEditor::create(temp.path().join(format!("blank.{extension}")))
            .await
            .unwrap()
            .snapshot()
            .unwrap();
        let report = document.issue_view().unwrap();
        assert_eq!(report.count, 0);
        assert!(!report.truncated);
    }

    let document = NativeOfficeEditor::create(temp.path().join("bounded.docx"))
        .await
        .unwrap()
        .snapshot()
        .unwrap();
    assert_eq!(
        document
            .issues(NativeOfficeIssueOptions {
                filter: None,
                limit: 0,
            })
            .unwrap_err()
            .code,
        "use.office.issue_limit_invalid"
    );
    assert_eq!(
        NativeOfficeIssueFilter::parse("not-real").unwrap_err().code,
        "use.office.issue_filter_invalid"
    );
}

#[test]
fn public_issue_contracts_are_send_and_sync() {
    fn assert_send_sync<T: Send + Sync>() {}

    assert_send_sync::<NativeOfficeIssueCategory>();
    assert_send_sync::<NativeOfficeIssueSeverity>();
    assert_send_sync::<NativeOfficeIssueSubtype>();
    assert_send_sync::<NativeOfficeIssueFilter>();
    assert_send_sync::<NativeOfficeIssueOptions>();
    assert_send_sync::<NativeOfficeIssueReport>();
}
