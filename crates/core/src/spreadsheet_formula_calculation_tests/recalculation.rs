#[tokio::test]
async fn calculation_treats_an_explicit_empty_string_cell_as_a_spill_obstruction() {
    let temp = tempfile::tempdir().unwrap();
    let mut editor = NativeOfficeEditor::create(temp.path().join("empty-obstruction.xlsx"))
        .await
        .unwrap();
    editor
        .set_cell_value("/Sheet1/A1", formula("SEQUENCE(1,2,1,1)"))
        .unwrap();
    editor.set_cell_value("/Sheet1/B1", text("")).unwrap();

    let calculation = editor
        .snapshot()
        .unwrap()
        .calculate_spreadsheet_formulas()
        .unwrap();
    assert_eq!(
        calculation.cells[0].value,
        SpreadsheetFormulaValue::error(SpreadsheetFormulaErrorLiteral::Spill)
    );
    assert_eq!(calculation.cells[0].spill_range, None);
}

#[tokio::test]
async fn calculation_rejects_cycles_and_unregistered_functions_without_mutation() {
    let temp = tempfile::tempdir().unwrap();
    let mut cycle = NativeOfficeEditor::create(temp.path().join("cycle.xlsx"))
        .await
        .unwrap();
    cycle.set_cell_value("/Sheet1/A1", formula("B1+1")).unwrap();
    cycle.set_cell_value("/Sheet1/B1", formula("A1+1")).unwrap();
    assert_eq!(
        cycle
            .snapshot()
            .unwrap()
            .calculate_spreadsheet_formulas()
            .unwrap_err()
            .code,
        "use.office.spreadsheet_formula_cycle"
    );

    let mut unsupported = NativeOfficeEditor::create(temp.path().join("unsupported.xlsx"))
        .await
        .unwrap();
    unsupported
        .set_cell_value("/Sheet1/A1", formula("SHELL(\"echo unsafe\")"))
        .unwrap();
    assert_eq!(
        unsupported
            .snapshot()
            .unwrap()
            .calculate_spreadsheet_formulas()
            .unwrap_err()
            .code,
        "use.office.spreadsheet_formula_function_unsupported"
    );
}

#[tokio::test]
async fn calculation_rejects_array_broadcasts_before_oversized_allocation() {
    let temp = tempfile::tempdir().unwrap();
    let mut editor = NativeOfficeEditor::create(temp.path().join("array-limit.xlsx"))
        .await
        .unwrap();
    editor
        .set_cell_value(
            "/Sheet1/A1",
            formula("SEQUENCE(100000,1)+TRANSPOSE(SEQUENCE(100000,1))"),
        )
        .unwrap();
    let error = editor
        .snapshot()
        .unwrap()
        .calculate_spreadsheet_formulas()
        .unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_formula_spill_limit");
    assert_eq!(error.details["cells"], 10_000_000_000_u64);
}

#[tokio::test]
async fn calculation_bounds_cumulative_function_argument_arrays() {
    let temp = tempfile::tempdir().unwrap();
    let mut editor = NativeOfficeEditor::create(temp.path().join("function-argument-limit.xlsx"))
        .await
        .unwrap();
    editor
        .set_cell_value(
            "/Sheet1/A1",
            formula("SUM(SEQUENCE(60000,1),SEQUENCE(60000,1))"),
        )
        .unwrap();
    let error = editor
        .snapshot()
        .unwrap()
        .calculate_spreadsheet_formulas()
        .unwrap_err();
    assert_eq!(
        error.code,
        "use.office.spreadsheet_formula_function_array_limit"
    );
    assert_eq!(error.details["cells"], 120_000);
}

#[tokio::test]
async fn calculation_bounds_cumulative_spill_cells() {
    let temp = tempfile::tempdir().unwrap();
    let mut editor = NativeOfficeEditor::create(temp.path().join("spill-total-limit.xlsx"))
        .await
        .unwrap();
    editor
        .set_cell_value("/Sheet1/A1", formula("SEQUENCE(60000,1)"))
        .unwrap();
    editor
        .set_cell_value("/Sheet1/C1", formula("SEQUENCE(60000,1)"))
        .unwrap();
    let error = editor
        .snapshot()
        .unwrap()
        .calculate_spreadsheet_formulas()
        .unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_formula_spill_limit");
    assert_eq!(error.details["cells"], 119_998);
}

#[tokio::test]
async fn calculation_bounds_concatenated_text_results() {
    let temp = tempfile::tempdir().unwrap();
    let mut editor = NativeOfficeEditor::create(temp.path().join("text-limit.xlsx"))
        .await
        .unwrap();
    let chunk = "x".repeat(600_000);
    editor.set_cell_value("/Sheet1/A1", text(&chunk)).unwrap();
    editor.set_cell_value("/Sheet1/A2", text(&chunk)).unwrap();
    editor
        .set_cell_value("/Sheet1/A3", formula("A1&A2"))
        .unwrap();
    let error = editor
        .snapshot()
        .unwrap()
        .calculate_spreadsheet_formulas()
        .unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_formula_text_limit");
    assert_eq!(error.details["bytes"], 1_200_000);
}

#[tokio::test]
async fn calculation_bounds_passthrough_text_results() {
    let temp = tempfile::tempdir().unwrap();
    let mut editor = NativeOfficeEditor::create(temp.path().join("passthrough-text-limit.xlsx"))
        .await
        .unwrap();
    let oversized = "x".repeat(MAX_SPREADSHEET_FORMULA_TEXT_BYTES + 1);
    editor
        .set_cell_value("/Sheet1/A1", text(&oversized))
        .unwrap();
    editor.set_cell_value("/Sheet1/A2", formula("A1")).unwrap();
    let error = editor
        .snapshot()
        .unwrap()
        .calculate_spreadsheet_formulas()
        .unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_formula_text_limit");
    assert_eq!(
        error.details["bytes"],
        MAX_SPREADSHEET_FORMULA_TEXT_BYTES + 1
    );
}

#[tokio::test]
async fn calculation_bounds_cumulative_text_results() {
    let temp = tempfile::tempdir().unwrap();
    let mut editor = NativeOfficeEditor::create(temp.path().join("text-total-limit.xlsx"))
        .await
        .unwrap();
    let text_result = "x".repeat(MAX_SPREADSHEET_FORMULA_TEXT_BYTES);
    editor
        .set_cell_value("/Sheet1/A1", text(&text_result))
        .unwrap();
    for row in 1..=9 {
        editor
            .set_cell_value(format!("/Sheet1/B{row}"), formula("A1"))
            .unwrap();
    }
    let error = editor
        .snapshot()
        .unwrap()
        .calculate_spreadsheet_formulas()
        .unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_formula_text_limit");
    assert_eq!(
        error.details["bytes"],
        MAX_SPREADSHEET_FORMULA_CALCULATION_TEXT_BYTES + MAX_SPREADSHEET_FORMULA_TEXT_BYTES
    );
}

#[tokio::test]
async fn calculation_bounds_broadcast_text_before_oversized_array_allocation() {
    let temp = tempfile::tempdir().unwrap();
    let mut editor = NativeOfficeEditor::create(temp.path().join("text-array-limit.xlsx"))
        .await
        .unwrap();
    let text_result = "x".repeat(MAX_SPREADSHEET_FORMULA_TEXT_BYTES);
    editor
        .set_cell_value("/Sheet1/A1", text(&text_result))
        .unwrap();
    editor
        .set_cell_value("/Sheet1/B1", formula("IF(SEQUENCE(9,1)>0,A1,\"\")"))
        .unwrap();
    let error = editor
        .snapshot()
        .unwrap()
        .calculate_spreadsheet_formulas()
        .unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_formula_text_limit");
    assert_eq!(
        error.details["bytes"],
        MAX_SPREADSHEET_FORMULA_CALCULATION_TEXT_BYTES + MAX_SPREADSHEET_FORMULA_TEXT_BYTES
    );
}

#[tokio::test]
async fn recalculation_atomically_writes_cached_values_spills_and_calculation_metadata() {
    let temp = tempfile::tempdir().unwrap();
    let mut editor = NativeOfficeEditor::create(temp.path().join("recalculate.xlsx"))
        .await
        .unwrap();
    editor.set_cell_value("/Sheet1/A1", number("2")).unwrap();
    editor
        .set_cell_value("/Sheet1/B1", formula("A1*3"))
        .unwrap();
    editor
        .set_cell_value("/Sheet1/C1", formula("IF(B1=6,\"yes\",\"no\")"))
        .unwrap();
    editor.set_cell_value("/Sheet1/D1", formula("1/0")).unwrap();
    editor
        .set_cell_value("/Sheet1/E1", formula("SEQUENCE(2,2,1,1)"))
        .unwrap();

    let calculation = editor.recalculate_spreadsheet_formulas().unwrap();
    assert_eq!(calculation.formula_count, 4);
    assert_eq!(calculation.spill_cell_count, 3);

    let document = editor.snapshot().unwrap();
    let b1 = document.get("/Sheet1/B1", 0).unwrap();
    assert_eq!(b1.text, "6");
    assert_eq!(b1.format.get("formula").map(String::as_str), Some("A1*3"));
    let c1 = document.get("/Sheet1/C1", 0).unwrap();
    assert_eq!(c1.text, "yes");
    assert_eq!(
        c1.format.get("valueType").map(String::as_str),
        Some("String")
    );
    let d1 = document.get("/Sheet1/D1", 0).unwrap();
    assert_eq!(d1.text, "#DIV/0!");
    assert_eq!(
        d1.format.get("valueType").map(String::as_str),
        Some("Error")
    );
    let e1 = document.get("/Sheet1/E1", 0).unwrap();
    assert_eq!(e1.text, "1");
    assert_eq!(
        e1.format.get("formulaType").map(String::as_str),
        Some("array")
    );
    assert_eq!(
        e1.format.get("formulaRef").map(String::as_str),
        Some("E1:F2")
    );
    for (path, expected) in [
        ("/Sheet1/F1", "2"),
        ("/Sheet1/E2", "3"),
        ("/Sheet1/F2", "4"),
    ] {
        let cell = document.get(path, 0).unwrap();
        assert_eq!(cell.text, expected);
        assert!(!cell.format.contains_key("formula"));
    }
    let workbook =
        String::from_utf8(editor.package().part("xl/workbook.xml").unwrap().to_vec()).unwrap();
    assert!(workbook.contains("calcCompleted=\"1\""), "{workbook}");
    assert!(workbook.contains("forceFullCalc=\"0\""), "{workbook}");
    assert!(workbook.contains("fullCalcOnLoad=\"0\""), "{workbook}");
    assert!(!editor.package().contains_part("xl/calcChain.xml"));
}

#[tokio::test]
async fn recalculation_preserves_explicit_normal_formula_storage() {
    let temp = tempfile::tempdir().unwrap();
    let mut editor = NativeOfficeEditor::create(temp.path().join("normal-formula.xlsx"))
        .await
        .unwrap();
    editor.set_cell_value("/Sheet1/A1", number("2")).unwrap();
    editor
        .set_cell_value("/Sheet1/B1", formula("A1*3"))
        .unwrap();

    let mut package = editor.package().clone();
    let worksheet = String::from_utf8(package.part("xl/worksheets/sheet1.xml").unwrap().to_vec())
        .unwrap()
        .replace("<f>A1*3</f>", "<f t=\"normal\">A1*3</f>");
    assert!(worksheet.contains("<f t=\"normal\">A1*3</f>"));
    package
        .set_part("xl/worksheets/sheet1.xml", worksheet.into_bytes())
        .unwrap();
    let mut editor = NativeOfficeEditor::from_package(package).unwrap();

    editor.recalculate_spreadsheet_formulas().unwrap();
    let worksheet = String::from_utf8(
        editor
            .package()
            .part("xl/worksheets/sheet1.xml")
            .unwrap()
            .to_vec(),
    )
    .unwrap();
    assert!(worksheet.contains("<f t=\"normal\">A1*3</f><v>6</v>"));
}

#[tokio::test]
async fn recalculation_rejects_malformed_formula_storage_without_mutation() {
    let temp = tempfile::tempdir().unwrap();
    let mut source = NativeOfficeEditor::create(temp.path().join("malformed-storage.xlsx"))
        .await
        .unwrap();
    source.set_cell_value("/Sheet1/A1", formula("1+1")).unwrap();

    for replacement in [
        "<f t=\"normal\" ref=\"A1:A2\">1+1</f>",
        "<f t=\"array\">1+1</f>",
    ] {
        let mut package = source.package().clone();
        let worksheet =
            String::from_utf8(package.part("xl/worksheets/sheet1.xml").unwrap().to_vec())
                .unwrap()
                .replace("<f>1+1</f>", replacement);
        assert!(worksheet.contains(replacement));
        package
            .set_part("xl/worksheets/sheet1.xml", worksheet.into_bytes())
            .unwrap();
        let mut editor = NativeOfficeEditor::from_package(package).unwrap();
        let before = editor.package().content_sha256();

        let error = editor.recalculate_spreadsheet_formulas().unwrap_err();
        assert_eq!(error.code, "use.office.spreadsheet_formula_storage_invalid");
        assert_eq!(editor.package().content_sha256(), before);
    }
}

#[tokio::test]
async fn recalculation_clears_cells_from_a_previous_larger_spill() {
    let temp = tempfile::tempdir().unwrap();
    let mut editor = NativeOfficeEditor::create(temp.path().join("shrink-spill.xlsx"))
        .await
        .unwrap();
    editor.set_cell_value("/Sheet1/D1", number("2")).unwrap();
    editor
        .set_cell_value("/Sheet1/A1", formula("SEQUENCE(D1,2,1,1)"))
        .unwrap();
    editor.recalculate_spreadsheet_formulas().unwrap();
    assert!(editor.snapshot().unwrap().get("/Sheet1/B2", 0).is_ok());

    editor.set_cell_value("/Sheet1/D1", number("1")).unwrap();
    let calculation = editor.recalculate_spreadsheet_formulas().unwrap();
    assert_eq!(calculation.cells[0].spill_range.as_deref(), Some("A1:B1"));
    let document = editor.snapshot().unwrap();
    assert!(document.get("/Sheet1/A2", 0).is_err());
    assert!(document.get("/Sheet1/B2", 0).is_err());
    let worksheet = String::from_utf8(
        editor
            .package()
            .part("xl/worksheets/sheet1.xml")
            .unwrap()
            .to_vec(),
    )
    .unwrap();
    assert!(!worksheet.contains("r=\"A2\""), "{worksheet}");
    assert!(!worksheet.contains("r=\"B2\""), "{worksheet}");
}

#[tokio::test]
async fn spilled_cells_are_read_only_and_replacing_the_anchor_clears_the_spill() {
    let temp = tempfile::tempdir().unwrap();
    let mut editor = NativeOfficeEditor::create(temp.path().join("spill-edit.xlsx"))
        .await
        .unwrap();
    editor
        .set_cell_value("/Sheet1/A1", formula("SEQUENCE(2,2,1,1)"))
        .unwrap();
    editor.recalculate_spreadsheet_formulas().unwrap();

    let before = editor.package().content_sha256();
    let error = editor
        .set_cell_value("/Sheet1/B2", text("not allowed"))
        .unwrap_err();
    assert_eq!(
        error.code,
        "use.office.spreadsheet_formula_spill_cell_read_only"
    );
    assert_eq!(editor.package().content_sha256(), before);

    editor.set_cell_value("/Sheet1/A1", number("9")).unwrap();
    let document = editor.snapshot().unwrap();
    assert_eq!(document.get("/Sheet1/A1", 0).unwrap().text, "9");
    for path in ["/Sheet1/B1", "/Sheet1/A2", "/Sheet1/B2"] {
        assert!(document.get(path, 0).is_err(), "{path}");
    }
}

#[tokio::test]
async fn removing_a_formula_anchor_removes_its_spilled_cells() {
    let temp = tempfile::tempdir().unwrap();
    let mut editor = NativeOfficeEditor::create(temp.path().join("spill-remove.xlsx"))
        .await
        .unwrap();
    editor
        .set_cell_value("/Sheet1/C3", formula("SEQUENCE(2,2,1,1)"))
        .unwrap();
    editor.recalculate_spreadsheet_formulas().unwrap();
    editor.remove("/Sheet1/C3").unwrap();

    let document = editor.snapshot().unwrap();
    for path in ["/Sheet1/C3", "/Sheet1/D3", "/Sheet1/C4", "/Sheet1/D4"] {
        assert!(document.get(path, 0).is_err(), "{path}");
    }
}

#[tokio::test]
async fn recalculation_mutation_rolls_back_the_entire_batch_on_failure() {
    let temp = tempfile::tempdir().unwrap();
    let mut editor = NativeOfficeEditor::create(temp.path().join("rollback.xlsx"))
        .await
        .unwrap();
    editor
        .set_cell_value("/Sheet1/A1", formula("SHELL(\"unsafe\")"))
        .unwrap();
    let before = editor.package().content_sha256();
    let error = editor
        .apply_batch(&[
            NativeOfficeMutation::SetCellValue {
                path: "/Sheet1/B1".into(),
                value: text("must roll back"),
            },
            NativeOfficeMutation::RecalculateSpreadsheetFormulas,
        ])
        .unwrap_err();
    assert_eq!(
        error.code,
        "use.office.spreadsheet_formula_function_unsupported"
    );
    assert_eq!(editor.package().content_sha256(), before);
    assert!(editor.snapshot().unwrap().get("/Sheet1/B1", 0).is_err());
}

#[tokio::test]
async fn recalculated_formulas_and_spills_are_exactly_replayable() {
    let temp = tempfile::tempdir().unwrap();
    let mut source = NativeOfficeEditor::create(temp.path().join("source.xlsx"))
        .await
        .unwrap();
    source.set_cell_value("/Sheet1/A1", number("4")).unwrap();
    source
        .set_cell_value("/Sheet1/B1", formula("A1/2"))
        .unwrap();
    source
        .set_cell_value("/Sheet1/C1", formula("SEQUENCE(2,2,1,1)"))
        .unwrap();
    source.recalculate_spreadsheet_formulas().unwrap();

    let artifact = NativeOfficeReplayArtifact::dump(&source.snapshot().unwrap(), "/").unwrap();
    assert!(matches!(
        artifact.mutations.last(),
        Some(NativeOfficeMutation::RecalculateSpreadsheetFormulas)
    ));
    let mut restored = NativeOfficeEditor::create(temp.path().join("restored.xlsx"))
        .await
        .unwrap();
    restored.apply_replay(&artifact).unwrap();
    assert_eq!(
        restored.package().content_sha256(),
        source.package().content_sha256()
    );
}

#[tokio::test]
async fn exact_replay_rejects_an_uncached_array_formula() {
    let temp = tempfile::tempdir().unwrap();
    let mut editor = NativeOfficeEditor::create(temp.path().join("uncached-array.xlsx"))
        .await
        .unwrap();
    editor
        .set_cell_value("/Sheet1/A1", formula("SEQUENCE(2,1)"))
        .unwrap();

    let mut package = editor.package().clone();
    let worksheet = String::from_utf8(package.part("xl/worksheets/sheet1.xml").unwrap().to_vec())
        .unwrap()
        .replace(
            "<f>SEQUENCE(2,1)</f>",
            "<f t=\"array\" ref=\"A1:A2\">SEQUENCE(2,1)</f>",
        );
    assert!(worksheet.contains("<f t=\"array\" ref=\"A1:A2\">SEQUENCE(2,1)</f>"));
    package
        .set_part("xl/worksheets/sheet1.xml", worksheet.into_bytes())
        .unwrap();
    let editor = NativeOfficeEditor::from_package(package).unwrap();

    let error = NativeOfficeReplayArtifact::dump(&editor.snapshot().unwrap(), "/").unwrap_err();
    assert_eq!(error.code, "use.office.dump_unsupported");
    assert!(error.message.contains("no cached native result"));
}

#[tokio::test]
async fn exact_replay_rejects_explicit_normal_formula_storage() {
    let temp = tempfile::tempdir().unwrap();
    let mut editor = NativeOfficeEditor::create(temp.path().join("explicit-normal.xlsx"))
        .await
        .unwrap();
    editor.set_cell_value("/Sheet1/A1", formula("1+1")).unwrap();

    let mut package = editor.package().clone();
    let worksheet = String::from_utf8(package.part("xl/worksheets/sheet1.xml").unwrap().to_vec())
        .unwrap()
        .replace("<f>1+1</f>", "<f t=\"normal\">1+1</f>");
    assert!(worksheet.contains("<f t=\"normal\">1+1</f>"));
    package
        .set_part("xl/worksheets/sheet1.xml", worksheet.into_bytes())
        .unwrap();
    let editor = NativeOfficeEditor::from_package(package).unwrap();

    let error = NativeOfficeReplayArtifact::dump(&editor.snapshot().unwrap(), "/").unwrap_err();
    assert_eq!(error.code, "use.office.dump_unsupported");
    assert!(error.message.contains("not canonical replay input"));
}

#[test]
fn typed_function_registry_is_closed_and_send_sync() {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<SpreadsheetFormulaFunctionRegistry>();
    let registry = SpreadsheetFormulaFunctionRegistry::default();
    assert!(registry.contains("SUM"));
    assert!(registry.contains("_xlfn.SEQUENCE"));
    assert!(!registry.contains("SHELL"));
    assert_eq!(
        serde_json::to_value(NativeOfficeMutation::RecalculateSpreadsheetFormulas).unwrap(),
        serde_json::json!({"operation": "recalculate-spreadsheet-formulas"})
    );
}
