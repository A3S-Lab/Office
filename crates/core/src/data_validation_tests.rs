use crate::{
    NativeOfficeEditor, NativeOfficeMutation, NativeSpreadsheetDataValidation,
    NativeSpreadsheetDataValidationErrorStyle, NativeSpreadsheetDataValidationOperator,
    NativeSpreadsheetDataValidationType, OfficeNodeType,
};

const STRICT_SPREADSHEET_NAMESPACE: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";

fn list_validation(range: &str) -> NativeSpreadsheetDataValidation {
    NativeSpreadsheetDataValidation::new(
        NativeSpreadsheetDataValidationType::List,
        range,
        "Draft,Review,Approved",
    )
}

#[test]
fn data_validation_mutations_have_closed_typed_json_and_are_send_sync() {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<NativeSpreadsheetDataValidation>();
    assert_send_sync::<NativeSpreadsheetDataValidationType>();
    assert_send_sync::<NativeSpreadsheetDataValidationOperator>();
    assert_send_sync::<NativeSpreadsheetDataValidationErrorStyle>();

    let mutation = NativeOfficeMutation::AddDataValidation {
        sheet: "/Sheet1".into(),
        validation: list_validation("A2:A20")
            .with_range("C2:C20")
            .with_in_cell_dropdown(false),
    };
    assert_eq!(
        serde_json::to_value(&mutation).unwrap(),
        serde_json::json!({
            "operation": "add-data-validation",
            "sheet": "/Sheet1",
            "validation": {
                "type": "list",
                "ranges": ["A2:A20", "C2:C20"],
                "formula1": "Draft,Review,Approved",
                "inCellDropdown": false
            }
        })
    );

    let decoded: NativeOfficeMutation = serde_json::from_value(serde_json::json!({
        "operation": "set-data-validation",
        "path": "/Sheet1/dataValidation[1]",
        "validation": {
            "type": "whole",
            "ranges": ["B2:B50"],
            "operator": "between",
            "formula1": "1",
            "formula2": "100",
            "allowBlank": false,
            "errorStyle": "warning"
        }
    }))
    .unwrap();
    assert!(matches!(
        decoded,
        NativeOfficeMutation::SetDataValidation {
            path,
            validation: NativeSpreadsheetDataValidation {
                validation_type: NativeSpreadsheetDataValidationType::Whole,
                operator: Some(NativeSpreadsheetDataValidationOperator::Between),
                allow_blank: false,
                error_style: NativeSpreadsheetDataValidationErrorStyle::Warning,
                ..
            }
        } if path == "/Sheet1/dataValidation[1]"
    ));

    assert!(
        serde_json::from_value::<NativeOfficeMutation>(serde_json::json!({
            "operation": "add-data-validation",
            "sheet": "/Sheet1",
            "validation": {
                "type": "script",
                "ranges": ["A1"]
            }
        }))
        .is_err()
    );
}

#[tokio::test]
async fn data_validation_add_set_query_remove_and_reopen_are_native() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("validations.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor.set_text("/Sheet1/A2", "Draft").unwrap();

    let validation = list_validation("A20:A2")
        .with_range("C2:C20")
        .with_input_message("Status", "Choose a workflow state")
        .with_error_message(
            NativeSpreadsheetDataValidationErrorStyle::Stop,
            "Invalid status",
            "Choose one of the listed states",
        )
        .with_in_cell_dropdown(false);
    assert_eq!(
        editor.add_data_validation("/sheet1", validation).unwrap(),
        "/Sheet1/dataValidation[1]"
    );

    let worksheet = part_text(&editor, "xl/worksheets/sheet1.xml");
    assert!(worksheet.contains("<dataValidations count=\"1\">"));
    assert!(worksheet.contains("type=\"list\""));
    assert!(worksheet.contains("allowBlank=\"1\""));
    assert!(worksheet.contains("showInputMessage=\"1\""));
    assert!(worksheet.contains("showErrorMessage=\"1\""));
    assert!(worksheet.contains("showDropDown=\"1\""));
    assert!(worksheet.contains("sqref=\"A2:A20 C2:C20\""));
    assert!(worksheet.contains("<formula1>&quot;Draft,Review,Approved&quot;</formula1>"));

    let snapshot = editor.snapshot().unwrap();
    let node = snapshot.get("/Sheet1/dataValidation[1]", 0).unwrap();
    assert_eq!(node.node_type, OfficeNodeType::DataValidation);
    assert_eq!(node.text, "A2:A20 C2:C20");
    assert_eq!(node.format["type"], "list");
    assert_eq!(node.format["ref"], "A2:A20 C2:C20");
    assert_eq!(node.format["formula1"], "\"Draft,Review,Approved\"");
    assert_eq!(node.format["allowBlank"], "true");
    assert_eq!(node.format["showInput"], "true");
    assert_eq!(node.format["showError"], "true");
    assert_eq!(node.format["inCellDropdown"], "false");
    let observed = snapshot.get("/Sheet1/A2", 0).unwrap();
    assert_eq!(
        observed.format["dataValidation"],
        "/Sheet1/dataValidation[1]"
    );
    assert_eq!(observed.format["validationType"], "list");
    let virtual_blank = snapshot.get("/Sheet1/C3", 0).unwrap();
    assert_eq!(virtual_blank.format["empty"], "true");
    assert_eq!(
        virtual_blank.format["dataValidation"],
        "/Sheet1/dataValidation[1]"
    );
    let html = snapshot.html_view().unwrap();
    assert!(html
        .content
        .contains("data-validation=\"/Sheet1/dataValidation[1]\""));
    assert!(html.content.contains("data-validation-type=\"list\""));
    let svg = snapshot.svg_view().unwrap();
    assert!(svg
        .content
        .contains("data-validation=\"/Sheet1/dataValidation[1]\""));
    assert!(svg.content.contains("data-validation-type=\"list\""));
    let queried = snapshot.query("dataValidation[type=list]").unwrap();
    assert_eq!(queried.len(), 1);
    assert_eq!(queried[0].path, node.path);

    let updated = NativeSpreadsheetDataValidation::new(
        NativeSpreadsheetDataValidationType::Whole,
        "B2:B50",
        "18",
    )
    .with_operator(NativeSpreadsheetDataValidationOperator::Between)
    .with_formula2("120")
    .with_allow_blank(false)
    .with_error_message(
        NativeSpreadsheetDataValidationErrorStyle::Warning,
        "Age outside range",
        "Enter an age from 18 through 120",
    );
    assert_eq!(
        editor
            .set_data_validation("/sheet1/validation[1]", updated)
            .unwrap(),
        "/Sheet1/dataValidation[1]"
    );
    let updated = editor
        .snapshot()
        .unwrap()
        .get("/Sheet1/dataValidation[1]", 0)
        .unwrap();
    assert_eq!(updated.format["type"], "whole");
    assert_eq!(updated.format["ref"], "B2:B50");
    assert_eq!(updated.format["operator"], "between");
    assert_eq!(updated.format["formula1"], "18");
    assert_eq!(updated.format["formula2"], "120");
    assert_eq!(updated.format["allowBlank"], "false");
    assert_eq!(updated.format["errorStyle"], "warning");

    editor.save().await.unwrap();
    let mut reopened = NativeOfficeEditor::open(&path).await.unwrap();
    assert_eq!(
        reopened.remove("/Sheet1/dataValidation[1]").unwrap(),
        "/Sheet1/dataValidation[1]"
    );
    assert!(reopened
        .snapshot()
        .unwrap()
        .query("dataValidation")
        .unwrap()
        .is_empty());
    assert!(!part_text(&reopened, "xl/worksheets/sheet1.xml").contains("dataValidations"));
}

#[tokio::test]
async fn data_validation_normalizes_date_time_custom_and_list_formulas() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("formula-normalization.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();

    editor
        .add_data_validation(
            "/Sheet1",
            NativeSpreadsheetDataValidation::new(
                NativeSpreadsheetDataValidationType::Date,
                "A1:A10",
                "2024-01-01",
            )
            .with_operator(NativeSpreadsheetDataValidationOperator::Between)
            .with_formula2("2024-12-31"),
        )
        .unwrap();
    editor
        .add_data_validation(
            "/Sheet1",
            NativeSpreadsheetDataValidation::new(
                NativeSpreadsheetDataValidationType::List,
                "E1:E10",
                "=INDIRECT(\"Statuses\")",
            ),
        )
        .unwrap();
    editor
        .add_data_validation(
            "/Sheet1",
            NativeSpreadsheetDataValidation::new(
                NativeSpreadsheetDataValidationType::List,
                "F1:F10",
                "=$H$2#",
            ),
        )
        .unwrap();
    editor
        .add_data_validation(
            "/Sheet1",
            NativeSpreadsheetDataValidation::new(
                NativeSpreadsheetDataValidationType::Time,
                "B1:B10",
                "09:00:00",
            )
            .with_operator(NativeSpreadsheetDataValidationOperator::Between)
            .with_formula2("17:00:00"),
        )
        .unwrap();
    editor
        .add_data_validation(
            "/Sheet1",
            NativeSpreadsheetDataValidation::new(
                NativeSpreadsheetDataValidationType::Custom,
                "C1:C10",
                "=ISNUMBER(C1)",
            ),
        )
        .unwrap();
    editor
        .add_data_validation(
            "/Sheet1",
            NativeSpreadsheetDataValidation::new(
                NativeSpreadsheetDataValidationType::List,
                "D1:D10",
                "=$H$2:$H$5",
            ),
        )
        .unwrap();

    let snapshot = editor.snapshot().unwrap();
    assert_eq!(
        snapshot.get("/Sheet1/dataValidation[1]", 0).unwrap().format["formula1"],
        "45292"
    );
    assert_eq!(
        snapshot.get("/Sheet1/dataValidation[1]", 0).unwrap().format["formula2"],
        "45657"
    );
    assert_eq!(
        snapshot.get("/Sheet1/dataValidation[2]", 0).unwrap().format["formula1"],
        "INDIRECT(\"Statuses\")"
    );
    assert_eq!(
        snapshot.get("/Sheet1/dataValidation[3]", 0).unwrap().format["formula1"],
        "$H$2#"
    );
    assert_eq!(
        snapshot.get("/Sheet1/dataValidation[4]", 0).unwrap().format["formula1"],
        "0.375"
    );
    assert_eq!(
        snapshot.get("/Sheet1/dataValidation[5]", 0).unwrap().format["formula1"],
        "ISNUMBER(C1)"
    );
    assert_eq!(
        snapshot.get("/Sheet1/dataValidation[6]", 0).unwrap().format["formula1"],
        "$H$2:$H$5"
    );
}

#[tokio::test]
async fn data_validation_honors_1904_dates_and_ooxml_operator_defaults() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("date-system.xlsx");
    let editor = NativeOfficeEditor::create(&path).await.unwrap();
    let mut package = editor.package().clone();
    let workbook = String::from_utf8(package.part("xl/workbook.xml").unwrap().to_vec())
        .unwrap()
        .replacen("<bookViews>", "<workbookPr date1904=\"1\"/><bookViews>", 1);
    package
        .set_part("xl/workbook.xml", workbook.into_bytes())
        .unwrap();
    let mut editor = NativeOfficeEditor::from_package(package).unwrap();
    editor
        .add_data_validation(
            "/Sheet1",
            NativeSpreadsheetDataValidation::new(
                NativeSpreadsheetDataValidationType::Date,
                "A1:A10",
                "2024-01-01",
            )
            .with_operator(NativeSpreadsheetDataValidationOperator::Between)
            .with_formula2("2024-12-31"),
        )
        .unwrap();
    let snapshot = editor.snapshot().unwrap();
    let date = snapshot.get("/Sheet1/dataValidation[1]", 0).unwrap();
    assert_eq!(date.format["formula1"], "43830");
    assert_eq!(date.format["formula2"], "44195");

    let original = part_text(&editor, "xl/worksheets/sheet1.xml");
    let worksheet = original.replacen(" operator=\"between\"", "", 1);
    assert_ne!(worksheet, original);
    let mut package = editor.package().clone();
    package
        .set_part("xl/worksheets/sheet1.xml", worksheet.into_bytes())
        .unwrap();
    let editor = NativeOfficeEditor::from_package(package).unwrap();
    assert_eq!(
        editor
            .snapshot()
            .unwrap()
            .get("/Sheet1/dataValidation[1]", 0)
            .unwrap()
            .format["operator"],
        "between"
    );
}

#[tokio::test]
async fn data_validation_rejects_invalid_and_overlapping_rules_atomically() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("invalid.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor
        .add_data_validation("/Sheet1", list_validation("A1:B10"))
        .unwrap();
    let before = editor.package().content_sha256();

    let error = editor
        .add_data_validation("/Sheet1", list_validation("B10:C20"))
        .unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_validation_overlap");
    assert_eq!(error.details["requested"], "B10:C20");
    assert_eq!(error.details["existing"], "A1:B10");
    assert_eq!(editor.package().content_sha256(), before);

    let error = editor
        .add_data_validation(
            "/Sheet1",
            NativeSpreadsheetDataValidation::new(
                NativeSpreadsheetDataValidationType::Whole,
                "D1:D10",
                "1",
            )
            .with_operator(NativeSpreadsheetDataValidationOperator::Between),
        )
        .unwrap_err();
    assert_eq!(
        error.code,
        "use.office.spreadsheet_validation_formula2_required"
    );
    assert_eq!(editor.package().content_sha256(), before);

    let error = editor
        .add_data_validation(
            "/Sheet1",
            NativeSpreadsheetDataValidation::new(
                NativeSpreadsheetDataValidationType::List,
                "D1:D10",
                "",
            ),
        )
        .unwrap_err();
    assert_eq!(
        error.code,
        "use.office.spreadsheet_validation_formula_required"
    );
    assert_eq!(editor.package().content_sha256(), before);

    let quoted_over_limit = format!("{},x", "a".repeat(253));
    let error = editor
        .add_data_validation(
            "/Sheet1",
            NativeSpreadsheetDataValidation::new(
                NativeSpreadsheetDataValidationType::List,
                "D1:D10",
                quoted_over_limit,
            ),
        )
        .unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_validation_text_invalid");
    assert_eq!(editor.package().content_sha256(), before);

    let error = editor
        .apply_batch(&[
            NativeOfficeMutation::AddDataValidation {
                sheet: "/Sheet1".into(),
                validation: list_validation("D1:D10"),
            },
            NativeOfficeMutation::AddDataValidation {
                sheet: "/Sheet1".into(),
                validation: list_validation("D10:E20"),
            },
        ])
        .unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_validation_overlap");
    assert_eq!(editor.package().content_sha256(), before);

    let word_path = temp.path().join("word.docx");
    let mut word = NativeOfficeEditor::create(&word_path).await.unwrap();
    let error = word
        .add_data_validation("/Sheet1", list_validation("A1:A2"))
        .unwrap_err();
    assert_eq!(error.code, "use.office.mutation_type_unsupported");
}

#[tokio::test]
async fn data_validation_preserves_strict_namespaces_order_and_unknown_attributes() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("strict-validation.xlsx");
    let editor = NativeOfficeEditor::create(&path).await.unwrap();
    let mut package = editor.package().clone();
    let worksheet = format!(
        "<worksheet xmlns=\"{STRICT_SPREADSHEET_NAMESPACE}\" xmlns:v=\"urn:vendor\"><dimension ref=\"A1\"/><sheetViews><sheetView workbookViewId=\"0\"/></sheetViews><sheetFormatPr defaultRowHeight=\"15\"/><sheetData/><conditionalFormatting sqref=\"A1\"><cfRule type=\"expression\" priority=\"1\"><formula>TRUE</formula></cfRule></conditionalFormatting><dataValidations count=\"1\" v:collection=\"keep\"><dataValidation type=\"list\" sqref=\"A1:A5\" allowBlank=\"1\" showInputMessage=\"1\" showErrorMessage=\"1\" v:id=\"keep\"><formula1>\"A,B\"</formula1></dataValidation><extLst><v:payload/></extLst></dataValidations><hyperlinks/><pageMargins left=\"0.7\" right=\"0.7\" top=\"0.75\" bottom=\"0.75\" header=\"0.3\" footer=\"0.3\"/></worksheet>"
    );
    package
        .set_part("xl/worksheets/sheet1.xml", worksheet.into_bytes())
        .unwrap();
    let mut editor = NativeOfficeEditor::from_package(package).unwrap();

    editor
        .set_data_validation(
            "/Sheet1/dataValidation[1]",
            NativeSpreadsheetDataValidation::new(
                NativeSpreadsheetDataValidationType::Custom,
                "B1:B5",
                "=B1<>\"\"",
            ),
        )
        .unwrap();
    editor
        .add_data_validation("/Sheet1", list_validation("D1:D5"))
        .unwrap();
    let edited = part_text(&editor, "xl/worksheets/sheet1.xml");
    assert!(edited.contains(STRICT_SPREADSHEET_NAMESPACE));
    assert!(edited.contains("v:collection=\"keep\""));
    assert!(edited.contains("v:id=\"keep\""));
    assert!(edited.contains("<v:payload/>"));
    assert!(edited.contains("count=\"2\""));
    assert!(
        edited.find("<conditionalFormatting").unwrap() < edited.find("<dataValidations").unwrap()
    );
    assert!(edited.find("<dataValidations").unwrap() < edited.find("<hyperlinks").unwrap());
    assert!(edited.find("sqref=\"D1:D5\"").unwrap() < edited.find("<extLst>").unwrap());

    editor.remove("/Sheet1/dataValidation[2]").unwrap();
    let preserved = part_text(&editor, "xl/worksheets/sheet1.xml");
    assert!(preserved.contains("v:collection=\"keep\""));
    assert!(preserved.contains("v:id=\"keep\""));
    assert!(preserved.contains("count=\"1\""));

    let before = editor.package().content_sha256();
    let error = editor.remove("/Sheet1/dataValidation[1]").unwrap_err();
    assert_eq!(
        error.code,
        "use.office.spreadsheet_validation_unknown_content"
    );
    assert_eq!(editor.package().content_sha256(), before);

    let worksheet = part_text(&editor, "xl/worksheets/sheet1.xml").replacen(
        "<formula1>",
        "<!--vendor-comment--><formula1>",
        1,
    );
    let mut package = editor.package().clone();
    package
        .set_part("xl/worksheets/sheet1.xml", worksheet.into_bytes())
        .unwrap();
    let mut editor = NativeOfficeEditor::from_package(package).unwrap();
    let before = editor.package().content_sha256();
    let error = editor
        .set_data_validation(
            "/Sheet1/dataValidation[1]",
            NativeSpreadsheetDataValidation::new(
                NativeSpreadsheetDataValidationType::List,
                "A1:A5",
                "One,Two",
            ),
        )
        .unwrap_err();
    assert_eq!(
        error.code,
        "use.office.spreadsheet_validation_unknown_content"
    );
    assert_eq!(editor.package().content_sha256(), before);
}

fn part_text(editor: &NativeOfficeEditor, part: &str) -> String {
    String::from_utf8(editor.package().part(part).unwrap().to_vec()).unwrap()
}
