use crate::{
    NativeOfficeEditor, NativeOfficeMutation, NativeSpreadsheetNamedRange,
    NativeSpreadsheetNamedRangeScope,
};

const SPREADSHEET_NAMESPACE: &str = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const STRICT_SPREADSHEET_NAMESPACE: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";

#[test]
fn named_range_mutations_have_stable_typed_json_and_are_send_sync() {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<NativeSpreadsheetNamedRange>();
    assert_send_sync::<NativeOfficeMutation>();

    let mutation = NativeOfficeMutation::AddNamedRange {
        named_range: NativeSpreadsheetNamedRange::new("Revenue", "A2:A20")
            .with_scope(NativeSpreadsheetNamedRangeScope::worksheet("Sheet1"))
            .with_comment("Workflow revenue")
            .with_volatile(true),
    };
    assert_eq!(
        serde_json::to_value(&mutation).unwrap(),
        serde_json::json!({
            "operation": "add-named-range",
            "namedRange": {
                "name": "Revenue",
                "ref": "A2:A20",
                "scope": "Sheet1",
                "comment": "Workflow revenue",
                "volatile": true
            }
        })
    );
    let decoded: NativeOfficeMutation = serde_json::from_value(serde_json::json!({
        "operation": "set-named-range",
        "path": "/namedrange[@name=Revenue][@scope=workbook]",
        "namedRange": {
            "name": "Revenue",
            "ref": "'Sheet1'!$A$2:$A$20"
        }
    }))
    .unwrap();
    assert!(matches!(
        decoded,
        NativeOfficeMutation::SetNamedRange {
            named_range: NativeSpreadsheetNamedRange {
                scope: NativeSpreadsheetNamedRangeScope::Workbook,
                ..
            },
            ..
        }
    ));

    let escaped = NativeSpreadsheetNamedRangeScope::worksheet("workbook");
    assert_eq!(
        serde_json::to_value(&escaped).unwrap(),
        serde_json::json!("worksheet:workbook")
    );
    assert_eq!(
        serde_json::from_value::<NativeSpreadsheetNamedRangeScope>(serde_json::json!(
            "worksheet:workbook"
        ))
        .unwrap(),
        escaped
    );
}

#[tokio::test]
async fn named_ranges_have_typed_scoped_lifecycle_and_stable_selectors() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("names.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();

    let workbook_path = editor
        .add_named_range(
            NativeSpreadsheetNamedRange::new("Revenue", "'Sheet1'!$A$2:$A$20")
                .with_comment("Workbook revenue")
                .with_volatile(true),
        )
        .unwrap();
    assert_eq!(workbook_path, "/namedrange[@name=Revenue][@scope=workbook]");

    let local_path = editor
        .add_named_range(
            NativeSpreadsheetNamedRange::new("Revenue", "A2:A20")
                .with_scope(NativeSpreadsheetNamedRangeScope::worksheet("Sheet1")),
        )
        .unwrap();
    assert_eq!(local_path, "/namedrange[@name=Revenue][@scope=Sheet1]");

    let snapshot = editor.snapshot().unwrap();
    let collection = snapshot.get("/namedrange", 1).unwrap();
    assert_eq!(collection.child_count, 2);
    assert_eq!(collection.children[0].path, workbook_path);
    assert_eq!(collection.children[0].format["name"], "Revenue");
    assert_eq!(collection.children[0].format["scope"], "workbook");
    assert_eq!(collection.children[0].format["volatile"], "true");
    assert_eq!(collection.children[0].format["comment"], "Workbook revenue");
    assert_eq!(
        snapshot.get("/namedrange[1]", 0).unwrap().path,
        workbook_path
    );
    assert_eq!(
        snapshot
            .get("/namedrange[@name=Revenue][@scope=Sheet1]", 0)
            .unwrap()
            .path,
        local_path
    );
    let ambiguous = snapshot.get("/namedrange[Revenue]", 0).unwrap_err();
    assert_eq!(
        ambiguous.code,
        "use.office.spreadsheet_named_range_ambiguous"
    );
    assert_eq!(
        snapshot
            .query("namedrange[name=Revenue][scope=Sheet1]")
            .unwrap()
            .len(),
        1
    );

    let updated_path = editor
        .set_named_range(
            &local_path,
            NativeSpreadsheetNamedRange::new("LocalRevenue", "'Sheet1'!$B$2:$B$20")
                .with_scope(NativeSpreadsheetNamedRangeScope::worksheet("Sheet1"))
                .with_comment("Local revenue"),
        )
        .unwrap();
    assert_eq!(
        updated_path,
        "/namedrange[@name=LocalRevenue][@scope=Sheet1]"
    );
    let updated = editor.snapshot().unwrap().get(&updated_path, 0).unwrap();
    assert_eq!(updated.format["ref"], "'Sheet1'!$B$2:$B$20");
    assert_eq!(updated.format["comment"], "Local revenue");
    assert_eq!(updated.format["volatile"], "false");

    assert_eq!(editor.remove(&workbook_path).unwrap(), workbook_path);
    assert_eq!(
        editor
            .snapshot()
            .unwrap()
            .query("namedrange")
            .unwrap()
            .len(),
        1
    );
    editor.save().await.unwrap();
    let reopened = NativeOfficeEditor::open(&path).await.unwrap();
    assert_eq!(
        reopened
            .snapshot()
            .unwrap()
            .get("/namedrange[LocalRevenue]", 0)
            .unwrap()
            .path,
        updated_path
    );
}

#[tokio::test]
async fn worksheet_named_workbook_has_an_unambiguous_scope_identity() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("workbook-scope.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor.rename_worksheet("/Sheet1", "workbook").unwrap();

    let global = editor
        .add_named_range(NativeSpreadsheetNamedRange::new(
            "Revenue",
            "'workbook'!$A$1",
        ))
        .unwrap();
    let local = editor
        .add_named_range(
            NativeSpreadsheetNamedRange::new("Revenue", "B1")
                .with_scope(NativeSpreadsheetNamedRangeScope::worksheet("workbook")),
        )
        .unwrap();
    assert_eq!(global, "/namedrange[@name=Revenue][@scope=workbook]");
    assert_eq!(
        local,
        "/namedrange[@name=Revenue][@scope=worksheet%3Aworkbook]"
    );

    let snapshot = editor.snapshot().unwrap();
    assert_eq!(
        snapshot.get(&global, 0).unwrap().format["scope"],
        "workbook"
    );
    assert_eq!(
        snapshot.get(&local, 0).unwrap().format["scope"],
        "worksheet:workbook"
    );
    let ambiguous = snapshot.get("/namedrange[Revenue]", 0).unwrap_err();
    assert_eq!(
        ambiguous.code,
        "use.office.spreadsheet_named_range_ambiguous"
    );
}

#[tokio::test]
async fn named_range_mutations_validate_identity_scope_and_atomic_uniqueness() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("validation.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();

    for (name, reference, code) in [
        (
            "A1",
            "'Sheet1'!$A$1",
            "use.office.spreadsheet_named_range_name_invalid",
        ),
        (
            "R",
            "'Sheet1'!$A$1",
            "use.office.spreadsheet_named_range_name_invalid",
        ),
        (
            "Bad Name",
            "'Sheet1'!$A$1",
            "use.office.spreadsheet_named_range_name_invalid",
        ),
        (
            "Valid",
            "='Sheet1'!$A$1",
            "use.office.spreadsheet_named_range_ref_invalid",
        ),
        (
            "Valid",
            "'[Legacy.xls]Sheet1'!$A$1",
            "use.office.spreadsheet_named_range_ref_invalid",
        ),
        (
            "_xlnm.Print_Area",
            "'Sheet1'!$A$1",
            "use.office.spreadsheet_named_range_reserved",
        ),
        (
            "Slicer_Cache",
            "#N/A",
            "use.office.spreadsheet_named_range_reserved",
        ),
    ] {
        let error = editor
            .add_named_range(NativeSpreadsheetNamedRange::new(name, reference))
            .unwrap_err();
        assert_eq!(error.code, code, "{name}");
    }

    let before = editor.package().content_sha256();
    let missing_scope = editor
        .add_named_range(
            NativeSpreadsheetNamedRange::new("MissingSheet", "A1")
                .with_scope(NativeSpreadsheetNamedRangeScope::worksheet("Missing")),
        )
        .unwrap_err();
    assert_eq!(
        missing_scope.code,
        "use.office.spreadsheet_named_range_scope_invalid"
    );
    assert_eq!(editor.package().content_sha256(), before);

    editor
        .add_named_range(NativeSpreadsheetNamedRange::new("Revenue", "'Sheet1'!$A$1"))
        .unwrap();
    let duplicate = editor
        .add_named_range(NativeSpreadsheetNamedRange::new("revenue", "'Sheet1'!$B$1"))
        .unwrap_err();
    assert_eq!(
        duplicate.code,
        "use.office.spreadsheet_named_range_duplicate"
    );

    editor
        .add_named_range(
            NativeSpreadsheetNamedRange::new("Revenue", "A1")
                .with_scope(NativeSpreadsheetNamedRangeScope::worksheet("Sheet1")),
        )
        .unwrap();
    let before = editor.package().content_sha256();
    let collision = editor
        .set_named_range(
            "/namedrange[@name=Revenue][@scope=Sheet1]",
            NativeSpreadsheetNamedRange::new("Revenue", "'Sheet1'!$C$1"),
        )
        .unwrap_err();
    assert_eq!(
        collision.code,
        "use.office.spreadsheet_named_range_duplicate"
    );
    assert_eq!(editor.package().content_sha256(), before);
}

#[tokio::test]
async fn named_range_edits_preserve_extensions_and_protect_builtin_names() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("preservation.xlsx");
    let editor = NativeOfficeEditor::create(&path).await.unwrap();
    let mut package = editor.package().clone();
    let workbook = String::from_utf8(package.part("xl/workbook.xml").unwrap().to_vec()).unwrap();
    let workbook = workbook.replace(
        "  <calcPr",
        concat!(
            "  <definedNames vendor=\"keep\">",
            "<definedName name=\"Editable\" vendor:token=\"42\" xmlns:vendor=\"urn:vendor\">'Sheet1'!$A$1</definedName>",
            "<definedName name=\"_xlnm.Print_Area\" localSheetId=\"0\">'Sheet1'!$A$1:$B$2</definedName>",
            "<definedName name=\"Slicer_Cache\">#N/A</definedName>",
            "</definedNames>\n  <calcPr"
        ),
    );
    package
        .set_part("xl/workbook.xml", workbook.into_bytes())
        .unwrap();
    let mut editor = NativeOfficeEditor::from_package(package).unwrap();

    let path = editor
        .set_named_range(
            "/namedrange[Editable]",
            NativeSpreadsheetNamedRange::new("Editable", "'Sheet1'!$C$1").with_comment("updated"),
        )
        .unwrap();
    assert_eq!(path, "/namedrange[@name=Editable][@scope=workbook]");
    let xml = workbook_xml(&editor);
    assert!(xml.contains("vendor=\"keep\""));
    assert!(xml.contains("vendor:token=\"42\""));
    assert!(xml.contains("comment=\"updated\""));

    for protected in ["/namedrange[_xlnm.Print_Area]", "/namedrange[Slicer_Cache]"] {
        let before = editor.package().content_sha256();
        let error = editor.remove(protected).unwrap_err();
        assert_eq!(error.code, "use.office.spreadsheet_named_range_reserved");
        assert_eq!(editor.package().content_sha256(), before);
    }

    let mut unknown = editor.package().clone();
    let xml = workbook_xml(&editor).replace(
        "</definedNames>",
        "<vendor:future xmlns:vendor=\"urn:vendor\"/></definedNames>",
    );
    unknown
        .set_part("xl/workbook.xml", xml.into_bytes())
        .unwrap();
    let mut unknown = NativeOfficeEditor::from_package(unknown).unwrap();
    let before = unknown.package().content_sha256();
    let error = unknown
        .add_named_range(NativeSpreadsheetNamedRange::new("Blocked", "'Sheet1'!$D$1"))
        .unwrap_err();
    assert_eq!(
        error.code,
        "use.office.spreadsheet_named_range_unknown_content"
    );
    assert_eq!(unknown.package().content_sha256(), before);
}

#[tokio::test]
async fn named_ranges_preserve_strict_spreadsheetml_and_workbook_child_order() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("strict-names.xlsx");
    let editor = NativeOfficeEditor::create(&path).await.unwrap();
    let mut package = editor.package().clone();
    let workbook =
        workbook_xml(&editor).replace(SPREADSHEET_NAMESPACE, STRICT_SPREADSHEET_NAMESPACE);
    package
        .set_part("xl/workbook.xml", workbook.into_bytes())
        .unwrap();
    let mut editor = NativeOfficeEditor::from_package(package).unwrap();

    let created = editor
        .add_named_range(
            NativeSpreadsheetNamedRange::new("StrictName", "A1:A5")
                .with_scope(NativeSpreadsheetNamedRangeScope::worksheet("Sheet1")),
        )
        .unwrap();
    let xml = workbook_xml(&editor);
    assert!(xml.contains(STRICT_SPREADSHEET_NAMESPACE));
    assert!(!xml.contains(SPREADSHEET_NAMESPACE));
    assert!(xml.find("<definedNames>").unwrap() < xml.find("<calcPr").unwrap());
    assert!(xml.contains("<definedName"));
    assert!(xml.contains("name=\"StrictName\""));
    assert!(xml.contains("localSheetId=\"0\""));
    assert!(xml.contains(">&apos;Sheet1&apos;!A1:A5</definedName>"));

    editor
        .set_named_range(
            &created,
            NativeSpreadsheetNamedRange::new("StrictName", "'Sheet1'!$B$1:$B$5")
                .with_scope(NativeSpreadsheetNamedRangeScope::worksheet("Sheet1"))
                .with_comment("strict"),
        )
        .unwrap();
    let snapshot = editor.snapshot().unwrap();
    assert_eq!(
        snapshot.get(&created, 0).unwrap().format["comment"],
        "strict"
    );
    assert!(workbook_xml(&editor).contains(STRICT_SPREADSHEET_NAMESPACE));
}

#[tokio::test]
async fn named_ranges_reject_listobject_name_collisions_atomically() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("table-name.xlsx");
    let editor = NativeOfficeEditor::create(&path).await.unwrap();
    let mut package = editor.package().clone();
    let table_part = "xl/tables/table1.xml";
    package
        .set_part(
            table_part,
            format!(
                "<table xmlns=\"{SPREADSHEET_NAMESPACE}\" id=\"1\" name=\"RevenueTable\" displayName=\"RevenueTable\" ref=\"A1:B2\"><autoFilter ref=\"A1:B2\"/><tableColumns count=\"2\"><tableColumn id=\"1\" name=\"Name\"/><tableColumn id=\"2\" name=\"Value\"/></tableColumns></table>"
            )
            .into_bytes(),
        )
        .unwrap();
    crate::opc_edit::add_content_type_override(
        &mut package,
        table_part,
        "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
    )
    .unwrap();
    crate::opc_edit::add_relationship(
        &mut package,
        "xl/worksheets/_rels/sheet1.xml.rels",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table",
        "../tables/table1.xml",
    )
    .unwrap();
    let mut editor = NativeOfficeEditor::from_package(package).unwrap();

    let before = editor.package().content_sha256();
    let error = editor
        .add_named_range(NativeSpreadsheetNamedRange::new(
            "revenuetable",
            "'Sheet1'!$A$1:$B$2",
        ))
        .unwrap_err();
    assert_eq!(
        error.code,
        "use.office.spreadsheet_named_range_name_collision"
    );
    assert_eq!(error.details["part"], table_part);
    assert_eq!(editor.package().content_sha256(), before);

    let existing = editor
        .add_named_range(NativeSpreadsheetNamedRange::new(
            "Metrics",
            "'Sheet1'!$B$1:$B$2",
        ))
        .unwrap();
    let before = editor.package().content_sha256();
    let error = editor
        .set_named_range(
            &existing,
            NativeSpreadsheetNamedRange::new("RevenueTable", "'Sheet1'!$B$1:$B$2"),
        )
        .unwrap_err();
    assert_eq!(
        error.code,
        "use.office.spreadsheet_named_range_name_collision"
    );
    assert_eq!(editor.package().content_sha256(), before);
}

#[tokio::test]
async fn named_range_removal_refuses_to_discard_unknown_name_content() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("unknown-name-content.xlsx");
    let editor = NativeOfficeEditor::create(&path).await.unwrap();
    let mut package = editor.package().clone();
    let workbook = workbook_xml(&editor).replace(
        "  <calcPr",
        concat!(
            "  <definedNames>",
            "<definedName name=\"CData\"><![CDATA['Sheet1'!$A$1]]></definedName>",
            "</definedNames>\n  <calcPr"
        ),
    );
    package
        .set_part("xl/workbook.xml", workbook.into_bytes())
        .unwrap();
    let mut editor = NativeOfficeEditor::from_package(package).unwrap();

    let before = editor.package().content_sha256();
    let error = editor.remove("/namedrange[CData]").unwrap_err();
    assert_eq!(
        error.code,
        "use.office.spreadsheet_named_range_unknown_content"
    );
    assert_eq!(editor.package().content_sha256(), before);
}

fn workbook_xml(editor: &NativeOfficeEditor) -> String {
    String::from_utf8(editor.package().part("xl/workbook.xml").unwrap().to_vec()).unwrap()
}
