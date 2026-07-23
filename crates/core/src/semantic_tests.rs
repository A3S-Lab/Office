use std::io::Write;
use std::path::{Path, PathBuf};

use tempfile::TempDir;
use zip::write::SimpleFileOptions;
use zip::{CompressionMethod, ZipWriter};

use super::{
    NativeOfficeDocument, NativeOfficeEditor, NativeOfficeMutation, OfficeNodeType,
    SpreadsheetCellValue,
};

const CONTENT_TYPES_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/package/2006/content-types";
const RELATIONSHIPS_NAMESPACE: &str =
    "http://schemas.openxmlformats.org/package/2006/relationships";
const OFFICE_DOCUMENT_RELATIONSHIP: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

#[tokio::test]
async fn native_word_read_supports_text_get_query_outline_and_statistics() {
    let (fixture, path) = word_fixture();
    let document = NativeOfficeDocument::open(&path).await.unwrap();

    assert_eq!(document.text_view().blocks[0].text, "Quarterly Results");
    assert!(document.text_view().text.contains("North\t100"));
    let heading = document.get("/body/p[1]", 1).unwrap();
    assert_eq!(heading.style.as_deref(), Some("Heading1"));
    assert_eq!(heading.children[0].format["bold"], "true");
    assert_eq!(
        document
            .get("/body/tbl[1]/tr[1]/tc[2]/p[1]", 0)
            .unwrap()
            .text,
        "100"
    );
    assert_eq!(
        document
            .query("p[style=Heading1] > r[bold=true]")
            .unwrap()
            .len(),
        1
    );
    assert!(document
        .outline()
        .iter()
        .any(|entry| entry.path == "/body/tbl[1]"));
    let statistics = document.statistics();
    assert_eq!(statistics.table_count, 1);
    assert_eq!(statistics.row_count, 1);
    assert_eq!(statistics.cell_count, 2);
    assert!(statistics.paragraph_count >= 4);
    drop(fixture);
}

#[tokio::test]
async fn native_spreadsheet_read_resolves_sheets_shared_strings_formulas_and_styles() {
    let (fixture, path) = spreadsheet_fixture();
    let document = NativeOfficeDocument::open(&path).await.unwrap();

    assert_eq!(document.get("/Sheet1/A1", 0).unwrap().text, "Revenue");
    let formula = document.get("/sheet1/b1", 0).unwrap();
    assert_eq!(formula.text, "42");
    assert_eq!(formula.format["formula"], "SUM(B2:B3)");
    assert_eq!(formula.format["numberFormat"], "0.00");
    assert_eq!(
        document
            .query("Sheet1!cell[formula][value>=40]")
            .unwrap()
            .len(),
        1
    );
    assert_eq!(document.query("cell[type=String]").unwrap().len(), 2);
    assert_eq!(document.query("B[value>=40]").unwrap().len(), 1);
    assert_eq!(
        document.query("/Sheet1/cell[type=String]").unwrap().len(),
        2
    );
    let range = document.get("/Sheet1/B2:A1", 1).unwrap();
    assert_eq!(range.format["normalizedRef"], "A1:B2");
    assert_eq!(range.child_count, 3);
    assert_eq!(document.get("/Sheet1/col[C]", 1).unwrap().child_count, 1);
    assert!(document.text_view().text.contains("/Sheet1/C2=North"));
    let statistics = document.statistics();
    assert_eq!(statistics.sheet_count, 1);
    assert_eq!(statistics.formula_count, 1);
    assert_eq!(statistics.cell_count, 4);
    drop(fixture);
}

#[tokio::test]
async fn native_presentation_read_resolves_slides_shapes_tables_pictures_and_notes() {
    let (fixture, path) = presentation_fixture();
    let document = NativeOfficeDocument::open(&path).await.unwrap();

    let title = document.get("/slide[1]/shape[1]", 2).unwrap();
    assert_eq!(title.node_type, OfficeNodeType::Placeholder);
    assert_eq!(title.text, "Native Office");
    assert_eq!(title.format["title"], "true");
    assert_eq!(
        document
            .get("/slide[1]/placeholder[title]", 0)
            .unwrap()
            .path,
        "/slide[1]/shape[1]"
    );
    assert_eq!(document.query("title").unwrap().len(), 1);
    assert_eq!(document.query("picture:no-alt").unwrap().len(), 1);
    assert_eq!(
        document
            .get("/slide[1]/table[1]/tr[1]/tc[2]", 1)
            .unwrap()
            .text,
        "Value"
    );
    assert_eq!(
        document.get("/slide[1]/notes", 0).unwrap().text,
        "Speaker note"
    );
    let statistics = document.statistics();
    assert_eq!(statistics.slide_count, 1);
    assert_eq!(statistics.picture_count, 1);
    assert_eq!(statistics.table_count, 1);
    assert_eq!(
        document.text_view().blocks[0].text,
        "Native Office\nBody text\nName\tValue"
    );
    drop(fixture);
}

#[tokio::test]
async fn native_editor_sets_text_in_all_formats_and_preserves_untouched_bytes() {
    let (word_temp, word_path) = word_fixture();
    let mut word = NativeOfficeEditor::open(&word_path).await.unwrap();
    let unknown = word.package().part("customXml/item1.xml").unwrap().to_vec();
    let before =
        String::from_utf8(word.package().part("word/document.xml").unwrap().to_vec()).unwrap();
    let untouched =
        r#"<w:p><w:r><w:t xml:space="preserve">Summary &amp; outlook</w:t></w:r></w:p>"#;
    assert!(before.contains(untouched));

    word.set_text("/body/p[1]/r[1]", " Native & < ").unwrap();

    assert_eq!(
        word.snapshot()
            .unwrap()
            .get("/body/p[1]/r[1]", 0)
            .unwrap()
            .text,
        " Native & < "
    );
    assert_eq!(word.package().part("customXml/item1.xml").unwrap(), unknown);
    let after =
        String::from_utf8(word.package().part("word/document.xml").unwrap().to_vec()).unwrap();
    assert!(after.contains(untouched));
    word.save().await.unwrap();
    assert_eq!(
        NativeOfficeDocument::open(&word_path)
            .await
            .unwrap()
            .get("/body/p[1]", 0)
            .unwrap()
            .text,
        " Native & < "
    );

    let (spreadsheet_temp, spreadsheet_path) = spreadsheet_fixture();
    let mut spreadsheet = NativeOfficeEditor::open(&spreadsheet_path).await.unwrap();
    spreadsheet.set_text("/Sheet1/B1", "forty two").unwrap();
    let cell = spreadsheet
        .snapshot()
        .unwrap()
        .get("/Sheet1/B1", 0)
        .unwrap();
    assert_eq!(cell.text, "forty two");
    assert_eq!(cell.format["valueType"], "String");
    assert_eq!(cell.format["styleIndex"], "1");
    assert!(!cell.format.contains_key("formula"));

    let (presentation_temp, presentation_path) = presentation_fixture();
    let mut presentation = NativeOfficeEditor::open(&presentation_path).await.unwrap();
    presentation
        .set_text("/slide[1]/shape[1]", "A3S & Office")
        .unwrap();
    assert_eq!(
        presentation
            .snapshot()
            .unwrap()
            .get("/slide[1]/shape[1]", 0)
            .unwrap()
            .text,
        "A3S & Office"
    );

    drop((word_temp, spreadsheet_temp, presentation_temp));
}

#[tokio::test]
async fn native_editor_sets_text_in_a_created_empty_word_paragraph() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("blank.docx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();

    editor.set_text("/body/p[1]", "First paragraph").unwrap();
    editor.save().await.unwrap();

    assert_eq!(
        NativeOfficeDocument::open(&path)
            .await
            .unwrap()
            .get("/body/p[1]", 1)
            .unwrap()
            .text,
        "First paragraph"
    );
}

#[tokio::test]
async fn native_editor_upserts_cells_in_a_created_empty_spreadsheet() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("blank.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();

    editor.set_text("/Sheet1/B2", "second").unwrap();
    editor.set_text("/Sheet1/A1", "first").unwrap();
    editor.set_text("/Sheet1/C2", "third").unwrap();
    editor.save().await.unwrap();

    let document = NativeOfficeDocument::open(&path).await.unwrap();
    assert_eq!(document.get("/Sheet1/A1", 0).unwrap().text, "first");
    assert_eq!(document.get("/Sheet1/B2", 0).unwrap().text, "second");
    assert_eq!(document.get("/Sheet1/C2", 0).unwrap().text, "third");
    let worksheet = String::from_utf8(
        document
            .package()
            .part("xl/worksheets/sheet1.xml")
            .unwrap()
            .to_vec(),
    )
    .unwrap();
    assert!(worksheet.contains("<dimension ref=\"A1:C2\"/>"));
    assert!(worksheet.find("r=\"A1\"").unwrap() < worksheet.find("r=\"B2\"").unwrap());
    assert!(worksheet.find("r=\"B2\"").unwrap() < worksheet.find("r=\"C2\"").unwrap());
}

#[tokio::test]
async fn native_editor_sets_and_removes_bounded_spreadsheet_ranges_atomically() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("ranges.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();

    editor.set_text("/Sheet1/C3:A2", "filled").unwrap();
    let range = editor.snapshot().unwrap().get("/Sheet1/A2:C3", 1).unwrap();
    assert_eq!(range.child_count, 6);
    assert!(range.children.iter().all(|cell| cell.text == "filled"));
    assert_eq!(range.format["normalizedRef"], "A2:C3");

    editor.remove("/Sheet1/B2:C3").unwrap();
    let remaining = editor.snapshot().unwrap().get("/Sheet1/A2:C3", 1).unwrap();
    assert_eq!(remaining.child_count, 2);
    assert_eq!(remaining.children[0].path, "/Sheet1/A2");
    assert_eq!(remaining.children[1].path, "/Sheet1/A3");

    let before = editor
        .package()
        .part("xl/worksheets/sheet1.xml")
        .unwrap()
        .to_vec();
    let error = editor
        .set_text("/Sheet1/A1:XFD1048576", "too large")
        .unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_range_too_large");
    assert_eq!(
        editor.package().part("xl/worksheets/sheet1.xml").unwrap(),
        before
    );

    editor.save().await.unwrap();
    let document = NativeOfficeDocument::open(&path).await.unwrap();
    let worksheet = String::from_utf8(
        document
            .package()
            .part("xl/worksheets/sheet1.xml")
            .unwrap()
            .to_vec(),
    )
    .unwrap();
    assert!(worksheet.contains("<dimension ref=\"A2:A3\"/>"));
    assert!(worksheet.find("r=\"A2\"").unwrap() < worksheet.find("r=\"A3\"").unwrap());
}

#[tokio::test]
async fn native_editor_writes_typed_spreadsheet_values_and_marks_formulas_for_recalculation() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("typed.xlsx");
    let editor = NativeOfficeEditor::create(&path).await.unwrap();
    let mut package = editor.package().clone();
    package
        .set_part(
            "xl/calcChain.xml",
            br#"<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A1" i="1"/></calcChain>"#
                .to_vec(),
        )
        .unwrap();
    crate::opc_edit::add_content_type_override(
        &mut package,
        "xl/calcChain.xml",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml",
    )
    .unwrap();
    crate::opc_edit::add_relationship(
        &mut package,
        "xl/_rels/workbook.xml.rels",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain",
        "calcChain.xml",
    )
    .unwrap();
    let mut editor = NativeOfficeEditor::from_package(package).unwrap();
    let original_workbook = editor.package().part("xl/workbook.xml").unwrap().to_vec();
    let original_worksheet = editor
        .package()
        .part("xl/worksheets/sheet1.xml")
        .unwrap()
        .to_vec();

    let error = editor
        .apply_batch(&[
            NativeOfficeMutation::SetCellValue {
                path: "/Sheet1/A1".into(),
                value: SpreadsheetCellValue::Formula {
                    expression: "=1+1".into(),
                },
            },
            NativeOfficeMutation::SetCellValue {
                path: "/Sheet1/B1".into(),
                value: SpreadsheetCellValue::Number {
                    value: "NaN".into(),
                },
            },
        ])
        .unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_number_invalid");
    assert_eq!(
        editor.package().part("xl/workbook.xml").unwrap(),
        original_workbook
    );
    assert_eq!(
        editor.package().part("xl/worksheets/sheet1.xml").unwrap(),
        original_worksheet
    );
    assert!(editor.package().contains_part("xl/calcChain.xml"));

    editor
        .set_cell_value(
            "/Sheet1/A1",
            SpreadsheetCellValue::Number {
                value: "42.5".into(),
            },
        )
        .unwrap();
    editor
        .set_cell_value("/Sheet1/B1", SpreadsheetCellValue::Boolean { value: true })
        .unwrap();
    editor
        .set_cell_value(
            "/Sheet1/C1",
            SpreadsheetCellValue::Formula {
                expression: "=A1*2".into(),
            },
        )
        .unwrap();
    assert_eq!(
        editor
            .set_cell_value(
                "/Sheet1/D1",
                SpreadsheetCellValue::Formula {
                    expression: "=".into(),
                },
            )
            .unwrap_err()
            .code,
        "use.office.spreadsheet_formula_invalid"
    );
    editor.save().await.unwrap();

    let document = NativeOfficeDocument::open(&path).await.unwrap();
    let number = document.get("/Sheet1/A1", 0).unwrap();
    assert_eq!(number.text, "42.5");
    assert_eq!(number.format["valueType"], "Number");
    let boolean = document.get("/Sheet1/B1", 0).unwrap();
    assert_eq!(boolean.text, "true");
    assert_eq!(boolean.format["valueType"], "Boolean");
    let formula = document.get("/Sheet1/C1", 0).unwrap();
    assert_eq!(formula.text, "");
    assert_eq!(formula.format["formula"], "A1*2");
    assert_eq!(formula.format["valueType"], "Number");

    let workbook =
        String::from_utf8(document.package().part("xl/workbook.xml").unwrap().to_vec()).unwrap();
    assert!(workbook.contains("calcMode=\"auto\""));
    assert!(workbook.contains("fullCalcOnLoad=\"1\""));
    assert!(workbook.contains("forceFullCalc=\"1\""));
    assert!(!document.package().contains_part("xl/calcChain.xml"));
    assert!(!String::from_utf8(
        document
            .package()
            .part("xl/_rels/workbook.xml.rels")
            .unwrap()
            .to_vec()
    )
    .unwrap()
    .contains("calcChain"));
}

#[tokio::test]
async fn native_editor_rejects_syntactically_invalid_formulas_atomically() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("invalid-formulas.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    let original = editor.package().content_sha256();

    for expression in ["1+", "SUM(A1", "A1::B2", "\"unterminated"] {
        let error = editor
            .set_cell_value(
                "/Sheet1/A1",
                SpreadsheetCellValue::Formula {
                    expression: expression.into(),
                },
            )
            .unwrap_err();
        assert_eq!(
            error.code, "use.office.spreadsheet_formula_invalid",
            "{expression}"
        );
        assert!(
            error.details.contains_key("characterOffset"),
            "{expression}"
        );
        assert!(error.details.contains_key("byteOffset"), "{expression}");
        assert_eq!(editor.package().content_sha256(), original, "{expression}");
    }
}

#[tokio::test]
async fn native_editor_adds_a_worksheet_and_populates_it() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("sheets.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();

    let sheet = editor.add_worksheet("Data").unwrap();
    editor.set_text("/Data/C3", "native sheet").unwrap();
    editor.save().await.unwrap();

    assert_eq!(sheet, "/Data");
    let document = NativeOfficeDocument::open(&path).await.unwrap();
    assert_eq!(document.statistics().sheet_count, 2);
    assert_eq!(document.get("/Data/C3", 0).unwrap().text, "native sheet");
    assert!(document.package().contains_part("xl/worksheets/sheet2.xml"));

    let error = editor.add_worksheet("data").unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_sheet_exists");
    assert_eq!(editor.remove("/Data").unwrap(), "/Data");
    let document = editor.snapshot().unwrap();
    assert_eq!(document.statistics().sheet_count, 1);
    assert!(!document.package().contains_part("xl/worksheets/sheet2.xml"));
    assert_eq!(
        editor.remove("/Sheet1").unwrap_err().code,
        "use.office.spreadsheet_last_sheet"
    );
}

#[tokio::test]
async fn native_editor_structurally_edits_spreadsheets_and_rewrites_cross_sheet_formulas() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("structure.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor.add_worksheet("Data").unwrap();
    editor
        .set_cell_value(
            "/Sheet1/A1",
            SpreadsheetCellValue::Number { value: "1".into() },
        )
        .unwrap();
    editor.set_text("/Sheet1/A3", "tail").unwrap();
    editor
        .set_cell_value(
            "/Sheet1/B2",
            SpreadsheetCellValue::Formula {
                expression: "A1+$A$3+'Data'!C4".into(),
            },
        )
        .unwrap();
    editor
        .set_cell_value(
            "/Data/C4",
            SpreadsheetCellValue::Formula {
                expression: "Sheet1!B2+C4".into(),
            },
        )
        .unwrap();

    assert_eq!(
        editor.insert_rows("/Sheet1", 2, 2).unwrap(),
        "/Sheet1/row[2:3]"
    );
    let snapshot = editor.snapshot().unwrap();
    assert_eq!(
        snapshot.get("/Sheet1/B4", 0).unwrap().format["formula"],
        "A1+$A$5+'Data'!C4"
    );
    assert_eq!(
        snapshot.get("/Data/C4", 0).unwrap().format["formula"],
        "Sheet1!B4+C4"
    );
    assert_eq!(snapshot.get("/Sheet1/A5", 0).unwrap().text, "tail");

    assert_eq!(
        editor.delete_columns("/Sheet1", "A", 1).unwrap(),
        "/Sheet1/col[1:1]"
    );
    let snapshot = editor.snapshot().unwrap();
    assert_eq!(
        snapshot.get("/Sheet1/A4", 0).unwrap().format["formula"],
        "#REF!+#REF!+'Data'!C4"
    );
    assert_eq!(
        snapshot.get("/Data/C4", 0).unwrap().format["formula"],
        "Sheet1!A4+C4"
    );
    assert_eq!(
        snapshot.get("/Sheet1/A1", 0).unwrap_err().code,
        "use.office.node_not_found"
    );

    assert_eq!(
        editor.rename_worksheet("/Data", "Q1 Data").unwrap(),
        "/Q1 Data"
    );
    assert_eq!(
        editor
            .snapshot()
            .unwrap()
            .get("/Sheet1/A4", 0)
            .unwrap()
            .format["formula"],
        "#REF!+#REF!+'Q1 Data'!C4"
    );
    assert_eq!(editor.move_worksheet("/Q1 Data", 1).unwrap(), "/Q1 Data");
    let snapshot = editor.snapshot().unwrap();
    assert_eq!(snapshot.root().children[0].path, "/Q1 Data");
    assert_eq!(snapshot.root().children[1].path, "/Sheet1");

    let before = editor
        .package()
        .part("xl/worksheets/sheet1.xml")
        .unwrap()
        .to_vec();
    editor.set_text("/Sheet1/XFD1", "edge").unwrap();
    let with_edge = editor
        .package()
        .part("xl/worksheets/sheet1.xml")
        .unwrap()
        .to_vec();
    let error = editor.insert_columns("/Sheet1", "XFD", 1).unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_structure_overflow");
    assert_eq!(
        editor.package().part("xl/worksheets/sheet1.xml").unwrap(),
        with_edge
    );
    assert_ne!(before, with_edge);

    editor.save().await.unwrap();
    NativeOfficeDocument::open(&path).await.unwrap();
}

#[tokio::test]
async fn native_spreadsheet_structure_preserves_and_rewrites_area_metadata() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("areas.xlsx");
    let editor = NativeOfficeEditor::create(&path).await.unwrap();
    let mut package = editor.package().clone();
    let worksheet = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <dimension ref="A1:B2"/>
  <sheetViews><sheetView workbookViewId="0"><selection activeCell="B2" sqref="B2"/></sheetView></sheetViews>
  <sheetFormatPr defaultRowHeight="15"/>
  <cols><col min="2" max="3" width="20" customWidth="1"/></cols>
  <sheetData><row r="1" spans="1:2"><c r="A1" t="inlineStr"><is><t>a</t></is></c><c r="B1" t="inlineStr"><is><t>b</t></is></c></row></sheetData>
  <autoFilter ref="A1:C10"/>
  <mergeCells count="2"><mergeCell ref="B2:C3"/><mergeCell ref="D5:E6"/></mergeCells>
  <dataValidations count="2"><dataValidation type="whole" sqref="B2:C3 D5"/><dataValidation type="list" sqref="F1"/></dataValidations>
  <pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>
</worksheet>"#;
    package
        .set_part("xl/worksheets/sheet1.xml", worksheet.as_bytes().to_vec())
        .unwrap();
    let mut editor = NativeOfficeEditor::from_package(package).unwrap();

    editor.insert_columns("/Sheet1", "B", 1).unwrap();
    let inserted = String::from_utf8(
        editor
            .package()
            .part("xl/worksheets/sheet1.xml")
            .unwrap()
            .to_vec(),
    )
    .unwrap();
    assert!(inserted.contains("min=\"3\""));
    assert!(inserted.contains("max=\"4\""));
    assert!(inserted.contains("ref=\"A1:D10\""));
    assert!(inserted.contains("ref=\"C2:D3\""));
    assert!(inserted.contains("ref=\"E5:F6\""));
    assert!(inserted.contains("sqref=\"C2:D3 E5\""));
    assert!(!inserted.contains("spans="));

    editor.delete_columns("/Sheet1", "C", 2).unwrap();
    let deleted = String::from_utf8(
        editor
            .package()
            .part("xl/worksheets/sheet1.xml")
            .unwrap()
            .to_vec(),
    )
    .unwrap();
    assert!(!deleted.contains("<col "));
    assert!(deleted.contains("ref=\"A1:B10\""));
    assert!(!deleted.contains("ref=\"C2:D3\""));
    assert!(deleted.contains("ref=\"C5:D6\""));
    assert!(deleted.contains("count=\"1\""));
    assert!(deleted.contains("sqref=\"C5\""));
    NativeOfficeDocument::from_package(editor.package().clone()).unwrap();
}

#[tokio::test]
async fn native_editor_adds_word_paragraphs_and_rolls_them_back_with_a_failed_batch() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("blank.docx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();

    let added = editor.add_paragraph("/body", "Second paragraph").unwrap();

    assert_eq!(added, "/body/p[2]");
    assert_eq!(
        editor.snapshot().unwrap().get(&added, 1).unwrap().text,
        "Second paragraph"
    );
    let error = editor
        .apply_batch(&[
            NativeOfficeMutation::AddParagraph {
                parent: "/body".into(),
                text: "must roll back".into(),
            },
            NativeOfficeMutation::SetText {
                path: "/body/p[999]".into(),
                text: "missing".into(),
            },
        ])
        .unwrap_err();
    assert_eq!(error.code, "use.office.node_not_found");
    assert_eq!(
        editor
            .snapshot()
            .unwrap()
            .get("/body", 1)
            .unwrap()
            .child_count,
        2
    );
}

#[tokio::test]
async fn native_editor_structurally_edits_word_tables_and_rolls_back_failed_batches() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("tables.docx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();

    assert_eq!(
        editor.add_table("/body", 1, 64).unwrap_err().code,
        "use.office.word_table_limit"
    );
    assert_eq!(
        editor.add_table("/body", 10_000, 63).unwrap_err().code,
        "use.office.word_table_limit"
    );

    let result = editor
        .apply_batch(&[
            NativeOfficeMutation::AddTable {
                parent: "/body".into(),
                rows: 2,
                columns: 2,
            },
            NativeOfficeMutation::SetText {
                path: "/body/tbl[1]/tr[1]/tc[1]".into(),
                text: "Name".into(),
            },
            NativeOfficeMutation::SetText {
                path: "/body/tbl[1]/tr[1]/tc[2]".into(),
                text: "Value".into(),
            },
        ])
        .unwrap();
    assert_eq!(result.applied, 3);
    assert_eq!(result.paths[0], "/body/tbl[1]");

    let row = editor.add_table_row("/body/tbl[1]", None).unwrap();
    let cell = editor.add_table_cell(&row, "extra").unwrap();
    assert_eq!(row, "/body/tbl[1]/tr[3]");
    assert_eq!(cell, "/body/tbl[1]/tr[3]/tc[3]");
    assert_eq!(
        editor.snapshot().unwrap().get(&cell, 1).unwrap().text,
        "extra"
    );

    let before = editor.snapshot().unwrap().root().clone();
    let error = editor
        .apply_batch(&[
            NativeOfficeMutation::AddTableRow {
                parent: "/body/tbl[1]".into(),
                columns: None,
            },
            NativeOfficeMutation::SetText {
                path: "/body/tbl[1]/tr[999]/tc[1]".into(),
                text: "missing".into(),
            },
        ])
        .unwrap_err();
    assert_eq!(error.code, "use.office.node_not_found");
    assert_eq!(editor.snapshot().unwrap().root(), &before);

    editor.remove(&cell).unwrap();
    editor.remove(&row).unwrap();
    let snapshot = editor.snapshot().unwrap();
    assert_eq!(snapshot.statistics().table_count, 1);
    assert_eq!(snapshot.statistics().row_count, 2);
    assert_eq!(snapshot.statistics().cell_count, 4);
    let document_xml = String::from_utf8(
        snapshot
            .package()
            .part("word/document.xml")
            .unwrap()
            .to_vec(),
    )
    .unwrap();
    assert_eq!(document_xml.matches("<w:gridCol").count(), 2);

    editor.remove("/body/tbl[1]/tr[2]").unwrap();
    assert_eq!(
        editor.remove("/body/tbl[1]/tr[1]").unwrap_err().code,
        "use.office.word_last_table_row"
    );
    editor.remove("/body/tbl[1]/tr[1]/tc[2]").unwrap();
    assert_eq!(
        editor.remove("/body/tbl[1]/tr[1]/tc[1]").unwrap_err().code,
        "use.office.word_last_table_cell"
    );
    assert_eq!(
        editor
            .remove("/body/tbl[1]/tr[1]/tc[1]/p[1]")
            .unwrap_err()
            .code,
        "use.office.word_last_cell_paragraph"
    );

    editor.save().await.unwrap();
    let reopened = NativeOfficeDocument::open(&path).await.unwrap();
    assert_eq!(
        reopened.get("/body/tbl[1]/tr[1]/tc[1]", 1).unwrap().text,
        "Name"
    );
}

#[tokio::test]
async fn native_editor_adds_and_edits_slides_in_a_created_presentation() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("blank.pptx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();

    let first = editor.add_slide("/", "Native title").unwrap();
    let second = editor.add_slide("/", "").unwrap();
    let shape = editor.add_shape(&second, "Native body").unwrap();
    editor
        .set_text("/slide[1]/shape[1]", "Updated title")
        .unwrap();
    editor.save().await.unwrap();

    assert_eq!(first, "/slide[1]");
    assert_eq!(second, "/slide[2]");
    assert_eq!(shape, "/slide[2]/shape[1]");
    let document = NativeOfficeDocument::open(&path).await.unwrap();
    assert_eq!(document.statistics().slide_count, 2);
    assert_eq!(
        document.get("/slide[1]/shape[1]", 1).unwrap().text,
        "Updated title"
    );
    assert_eq!(document.get("/slide[2]", 1).unwrap().child_count, 1);
    assert_eq!(document.get(&shape, 1).unwrap().text, "Native body");
    assert!(document
        .package()
        .contains_part("ppt/slides/_rels/slide1.xml.rels"));
    let presentation_xml = String::from_utf8(
        document
            .package()
            .part("ppt/presentation.xml")
            .unwrap()
            .to_vec(),
    )
    .unwrap();
    assert!(presentation_xml.contains("id=\"256\""));
    assert!(presentation_xml.contains("id=\"257\""));
}

#[tokio::test]
async fn native_editor_removes_core_nodes_across_all_formats() {
    let temp = tempfile::tempdir().unwrap();

    let word_path = temp.path().join("remove.docx");
    let mut word = NativeOfficeEditor::create(&word_path).await.unwrap();
    let paragraph = word.add_paragraph("/body", "remove me").unwrap();
    assert_eq!(word.remove(&paragraph).unwrap(), paragraph);
    assert_eq!(
        word.snapshot()
            .unwrap()
            .get("/body", 1)
            .unwrap()
            .child_count,
        1
    );

    let spreadsheet_path = temp.path().join("remove.xlsx");
    let mut spreadsheet = NativeOfficeEditor::create(&spreadsheet_path).await.unwrap();
    spreadsheet.set_text("/Sheet1/A1", "remove me").unwrap();
    spreadsheet.set_text("/Sheet1/B1", "keep me").unwrap();
    assert_eq!(spreadsheet.remove("/Sheet1/A1").unwrap(), "/Sheet1/A1");
    let spreadsheet = spreadsheet.snapshot().unwrap();
    assert_eq!(spreadsheet.get("/Sheet1/B1", 0).unwrap().text, "keep me");
    assert_eq!(
        spreadsheet.get("/Sheet1/A1", 0).unwrap_err().code,
        "use.office.node_not_found"
    );
    assert!(String::from_utf8(
        spreadsheet
            .package()
            .part("xl/worksheets/sheet1.xml")
            .unwrap()
            .to_vec()
    )
    .unwrap()
    .contains("<dimension ref=\"B1\"/>"));

    let presentation_path = temp.path().join("remove.pptx");
    let mut presentation = NativeOfficeEditor::create(&presentation_path)
        .await
        .unwrap();
    let slide = presentation.add_slide("/", "keep me").unwrap();
    let shape = presentation.add_shape(&slide, "remove me").unwrap();
    assert_eq!(presentation.remove(&shape).unwrap(), shape);
    let snapshot = presentation.snapshot().unwrap();
    assert_eq!(
        snapshot.get("/slide[1]/shape[1]", 0).unwrap().text,
        "keep me"
    );
    assert_eq!(
        snapshot.get("/slide[1]/shape[2]", 0).unwrap_err().code,
        "use.office.node_not_found"
    );
    drop(snapshot);
    assert_eq!(presentation.remove(&slide).unwrap(), slide);
    let snapshot = presentation.snapshot().unwrap();
    assert_eq!(snapshot.statistics().slide_count, 0);
    assert!(!snapshot.package().contains_part("ppt/slides/slide1.xml"));
    assert!(!snapshot
        .package()
        .contains_part("ppt/slides/_rels/slide1.xml.rels"));
}

#[tokio::test]
async fn native_editor_rolls_back_an_entire_failed_batch() {
    let (fixture, path) = word_fixture();
    let mut editor = NativeOfficeEditor::open(path).await.unwrap();

    let error = editor
        .apply_batch(&[
            NativeOfficeMutation::SetText {
                path: "/body/p[1]".into(),
                text: "would be applied".into(),
            },
            NativeOfficeMutation::SetText {
                path: "/body/p[999]".into(),
                text: "missing".into(),
            },
        ])
        .unwrap_err();

    assert_eq!(error.code, "use.office.node_not_found");
    assert_eq!(
        editor
            .snapshot()
            .unwrap()
            .get("/body/p[1]", 0)
            .unwrap()
            .text,
        "Quarterly Results"
    );
    assert!(!editor.is_dirty());
    drop(fixture);
}

fn word_fixture() -> (TempDir, PathBuf) {
    let content_types = format!(
        r#"<Types xmlns="{CONTENT_TYPES_NAMESPACE}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>"#
    );
    let root_relationships = root_relationships("word/document.xml");
    let document = r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:pStyle w:val="Heading1"/><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="28"/></w:rPr><w:t>Quarterly Results</w:t></w:r></w:p><w:p><w:r><w:t xml:space="preserve">Summary &amp; outlook</w:t></w:r></w:p><w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr><w:tr><w:tc><w:p><w:r><w:t>North</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>100</w:t></w:r></w:p></w:tc></w:tr></w:tbl><w:sectPr/></w:body></w:document>"#;
    let styles = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/></w:style><w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/></w:style></w:styles>"#;
    fixture(
        "document.docx",
        &[
            ("[Content_Types].xml", content_types.as_bytes()),
            ("_rels/.rels", root_relationships.as_bytes()),
            ("word/document.xml", document.as_bytes()),
            ("word/styles.xml", styles.as_bytes()),
            ("customXml/item1.xml", b"<keep exact='yes' />"),
        ],
    )
}

fn spreadsheet_fixture() -> (TempDir, PathBuf) {
    let content_types = format!(
        r#"<Types xmlns="{CONTENT_TYPES_NAMESPACE}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>"#
    );
    let root_relationships = root_relationships("xl/workbook.xml");
    let workbook = r#"<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets></workbook>"#;
    let workbook_relationships = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NAMESPACE}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>"#
    );
    let worksheet = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" s="1"><f>SUM(B2:B3)</f><v>42</v></c></row><row r="2"><c r="A2" t="b"><v>1</v></c><c r="C2" t="inlineStr"><is><t>North</t></is></c></row></sheetData></worksheet>"#;
    let shared_strings = r#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1" uniqueCount="1"><si><t>Revenue</t></si></sst>"#;
    let styles = r#"<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="165" formatCode="0.00"/></numFmts><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="165" fontId="0" fillId="0" borderId="0" applyNumberFormat="1"/></cellXfs></styleSheet>"#;
    fixture(
        "workbook.xlsx",
        &[
            ("[Content_Types].xml", content_types.as_bytes()),
            ("_rels/.rels", root_relationships.as_bytes()),
            ("xl/workbook.xml", workbook.as_bytes()),
            (
                "xl/_rels/workbook.xml.rels",
                workbook_relationships.as_bytes(),
            ),
            ("xl/worksheets/sheet1.xml", worksheet.as_bytes()),
            ("xl/sharedStrings.xml", shared_strings.as_bytes()),
            ("xl/styles.xml", styles.as_bytes()),
        ],
    )
}

fn presentation_fixture() -> (TempDir, PathBuf) {
    let content_types = format!(
        r#"<Types xmlns="{CONTENT_TYPES_NAMESPACE}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/><Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/><Override PartName="/ppt/notesSlides/notesSlide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/></Types>"#
    );
    let root_relationships = root_relationships("ppt/presentation.xml");
    let presentation = r#"<p:presentation xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:sldIdLst><p:sldId id="256" r:id="rId1"/></p:sldIdLst><p:sldSz cx="12192000" cy="6858000"/></p:presentation>"#;
    let presentation_relationships = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NAMESPACE}"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/></Relationships>"#
    );
    let slide = r#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5000000" cy="1000000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" sz="3200" b="1"/><a:t>Native Office</a:t></a:r></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="TextBox 2"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Body text</a:t></a:r></a:p></p:txBody></p:sp><p:pic><p:nvPicPr><p:cNvPr id="4" name="Picture 3"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdImage"/></p:blipFill><p:spPr/></p:pic><p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="Table 4"/><p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblGrid/><a:tr h="300000"><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Name</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc><a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Value</a:t></a:r></a:p></a:txBody><a:tcPr/></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame></p:spTree></p:cSld></p:sld>"#;
    let slide_relationships = format!(
        r#"<Relationships xmlns="{RELATIONSHIPS_NAMESPACE}"><Relationship Id="rIdImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/><Relationship Id="rIdNotes" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide1.xml"/></Relationships>"#
    );
    let notes = r#"<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Notes Placeholder"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Speaker note</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>"#;
    fixture(
        "deck.pptx",
        &[
            ("[Content_Types].xml", content_types.as_bytes()),
            ("_rels/.rels", root_relationships.as_bytes()),
            ("ppt/presentation.xml", presentation.as_bytes()),
            (
                "ppt/_rels/presentation.xml.rels",
                presentation_relationships.as_bytes(),
            ),
            ("ppt/slides/slide1.xml", slide.as_bytes()),
            (
                "ppt/slides/_rels/slide1.xml.rels",
                slide_relationships.as_bytes(),
            ),
            ("ppt/notesSlides/notesSlide1.xml", notes.as_bytes()),
            ("ppt/media/image1.png", b"fake png"),
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
