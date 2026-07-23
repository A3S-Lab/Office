use crate::{NativeOfficeDocument, NativeOfficeEditor, SpreadsheetCellValue};

#[tokio::test]
async fn structural_edits_rewrite_tables_comments_drawings_vml_and_chart_formulas() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("related.xlsx");
    let editor = NativeOfficeEditor::create(&path).await.unwrap();
    let mut package = editor.package().clone();

    let parts = [
        (
            "xl/tables/table1.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
            br#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:B3"><autoFilter ref="A1:B3"/><tableColumns count="2"><tableColumn id="1" name="A"/><tableColumn id="2" name="B"/></tableColumns></table>"#.as_slice(),
        ),
        (
            "xl/comments1.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
            br#"<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>A3S</author></authors><commentList><comment ref="B2" authorId="0"><text><t>note</t></text></comment></commentList></comments>"#.as_slice(),
        ),
        (
            "xl/drawings/vmlDrawing1.vml",
            "application/vnd.openxmlformats-officedocument.vmlDrawing",
            br#"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel"><v:shape id="note"><x:ClientData ObjectType="Note"><x:Row>1</x:Row><x:Column>1</x:Column></x:ClientData></v:shape></xml>"#.as_slice(),
        ),
        (
            "xl/drawings/drawing1.xml",
            "application/vnd.openxmlformats-officedocument.drawing+xml",
            br#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:row>1</xdr:row></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:row>2</xdr:row></xdr:to><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#.as_slice(),
        ),
        (
            "xl/charts/chart1.xml",
            "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numRef><c:f>'Sheet1'!$A$1:$B$3</c:f><c:numCache><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#.as_slice(),
        ),
    ];
    for (part, content_type, bytes) in parts {
        package.set_part(part, bytes.to_vec()).unwrap();
        crate::opc_edit::add_content_type_override(&mut package, part, content_type).unwrap();
    }
    for (relationship_type, target) in [
        (
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table",
            "../tables/table1.xml",
        ),
        (
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
            "../comments1.xml",
        ),
        (
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
            "../drawings/vmlDrawing1.vml",
        ),
        (
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
            "../drawings/drawing1.xml",
        ),
    ] {
        crate::opc_edit::add_relationship(
            &mut package,
            "xl/worksheets/_rels/sheet1.xml.rels",
            relationship_type,
            target,
        )
        .unwrap();
    }
    crate::opc_edit::add_relationship(
        &mut package,
        "xl/drawings/_rels/drawing1.xml.rels",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
        "../charts/chart1.xml",
    )
    .unwrap();

    let mut editor = NativeOfficeEditor::from_package(package).unwrap();
    editor.insert_rows("/Sheet1", 2, 1).unwrap();

    let table = part_text(&editor, "xl/tables/table1.xml");
    assert!(table.contains("ref=\"A1:B4\""));
    let comments = part_text(&editor, "xl/comments1.xml");
    assert!(comments.contains("ref=\"B3\""));
    let vml = part_text(&editor, "xl/drawings/vmlDrawing1.vml");
    assert!(vml.contains("<x:Row>2</x:Row>"));
    let drawing = part_text(&editor, "xl/drawings/drawing1.xml");
    assert!(drawing.contains("<xdr:row>2</xdr:row>"));
    assert!(drawing.contains("<xdr:row>3</xdr:row>"));
    let chart = part_text(&editor, "xl/charts/chart1.xml");
    assert!(chart.contains("&apos;Sheet1&apos;!$A$1:$B$4"), "{chart}");
    assert!(!chart.contains("numCache"));

    editor.insert_columns("/Sheet1", "B", 1).unwrap();
    let expanded_table = part_text(&editor, "xl/tables/table1.xml");
    assert!(expanded_table.contains("ref=\"A1:C4\""));
    assert!(expanded_table.contains("tableColumns count=\"3\""));
    assert!(expanded_table.contains("name=\"Column3\""));
    assert!(part_text(&editor, "xl/comments1.xml").contains("ref=\"C3\""));
    assert!(part_text(&editor, "xl/drawings/vmlDrawing1.vml").contains("<x:Column>2</x:Column>"));

    editor.delete_columns("/Sheet1", "B", 2).unwrap();
    let table = part_text(&editor, "xl/tables/table1.xml");
    assert!(table.contains("ref=\"A1:A4\""));
    assert!(table.contains("tableColumns count=\"1\""));
    assert!(!table.contains("name=\"B\""));
    assert!(!part_text(&editor, "xl/comments1.xml").contains("<comment "));
    assert!(!part_text(&editor, "xl/drawings/vmlDrawing1.vml").contains("<v:shape"));
    assert!(part_text(&editor, "xl/charts/chart1.xml").contains("&apos;Sheet1&apos;!$A$1:$A$4"));
    NativeOfficeDocument::from_package(editor.package().clone()).unwrap();
}

#[tokio::test]
async fn worksheet_copy_clones_related_parts_names_and_local_definitions_atomically() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("copy.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor
        .set_cell_value(
            "/Sheet1/A1",
            SpreadsheetCellValue::Formula {
                expression: "'Sheet1'!A2".into(),
            },
        )
        .unwrap();
    editor
        .set_cell_value(
            "/Sheet1/A2",
            SpreadsheetCellValue::Text {
                value: "source".into(),
            },
        )
        .unwrap();
    let mut package = editor.package().clone();

    let parts = [
        (
            "xl/tables/table1.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml",
            br#"<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:A2"><autoFilter ref="A1:A2"/><tableColumns count="1"><tableColumn id="1" name="Value"/></tableColumns></table>"#.as_slice(),
        ),
        (
            "xl/comments1.xml",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml",
            br#"<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors><author>A3S</author></authors><commentList><comment ref="A2" authorId="0"><text><t>note</t></text></comment></commentList></comments>"#.as_slice(),
        ),
        (
            "xl/drawings/vmlDrawing1.vml",
            "application/vnd.openxmlformats-officedocument.vmlDrawing",
            br#"<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:x="urn:schemas-microsoft-com:office:excel"><v:shape id="note"><x:ClientData ObjectType="Note"><x:Row>1</x:Row><x:Column>0</x:Column></x:ClientData></v:shape></xml>"#.as_slice(),
        ),
        (
            "xl/drawings/drawing1.xml",
            "application/vnd.openxmlformats-officedocument.drawing+xml",
            br#"<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><xdr:twoCellAnchor><xdr:from><xdr:col>0</xdr:col><xdr:row>0</xdr:row></xdr:from><xdr:to><xdr:col>2</xdr:col><xdr:row>4</xdr:row></xdr:to><xdr:graphicFrame><xdr:nvGraphicFramePr><xdr:cNvPr id="2" name="Chart 1"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm/><a:graphic><a:graphicData><c:chart r:id="rId1"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>"#.as_slice(),
        ),
        (
            "xl/charts/chart1.xml",
            "application/vnd.openxmlformats-officedocument.drawingml.chart+xml",
            br#"<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:barChart><c:ser><c:val><c:numRef><c:f>'Sheet1'!$A$1:$A$2</c:f></c:numRef></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>"#.as_slice(),
        ),
        (
            "xl/media/image1.png",
            "image/png",
            b"synthetic-png".as_slice(),
        ),
        (
            "xl/theme/theme1.xml",
            "application/vnd.openxmlformats-officedocument.theme+xml",
            br#"<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Shared"><a:themeElements/></a:theme>"#.as_slice(),
        ),
    ];
    for (part, content_type, bytes) in parts {
        package.set_part(part, bytes.to_vec()).unwrap();
        crate::opc_edit::add_content_type_override(&mut package, part, content_type).unwrap();
    }
    for (relationship_type, target) in [
        (
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table",
            "../tables/table1.xml",
        ),
        (
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments",
            "../comments1.xml",
        ),
        (
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
            "../drawings/vmlDrawing1.vml",
        ),
        (
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
            "../drawings/drawing1.xml",
        ),
        (
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme",
            "../theme/theme1.xml",
        ),
    ] {
        crate::opc_edit::add_relationship(
            &mut package,
            "xl/worksheets/_rels/sheet1.xml.rels",
            relationship_type,
            target,
        )
        .unwrap();
    }
    crate::opc_edit::add_relationship(
        &mut package,
        "xl/drawings/_rels/drawing1.xml.rels",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
        "../charts/chart1.xml",
    )
    .unwrap();
    crate::opc_edit::add_relationship(
        &mut package,
        "xl/drawings/_rels/drawing1.xml.rels",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        "../media/image1.png",
    )
    .unwrap();

    let workbook = String::from_utf8(package.part("xl/workbook.xml").unwrap().to_vec()).unwrap();
    package
        .set_part(
            "xl/workbook.xml",
            workbook
                .replace(
                    "  <calcPr",
                    "  <definedNames><definedName name=\"LocalValue\" localSheetId=\"0\">'Sheet1'!$A$1</definedName></definedNames>\n  <calcPr",
                )
                .into_bytes(),
        )
        .unwrap();

    let source_parts = [
        "xl/worksheets/sheet1.xml",
        "xl/worksheets/_rels/sheet1.xml.rels",
        "xl/tables/table1.xml",
        "xl/comments1.xml",
        "xl/drawings/vmlDrawing1.vml",
        "xl/drawings/drawing1.xml",
        "xl/drawings/_rels/drawing1.xml.rels",
        "xl/charts/chart1.xml",
        "xl/media/image1.png",
        "xl/theme/theme1.xml",
    ]
    .map(|part| (part, package.part(part).unwrap().to_vec()));

    let mut editor = NativeOfficeEditor::from_package(package).unwrap();
    assert_eq!(
        editor.copy_worksheet("/Sheet1", "Copy", Some(2)).unwrap(),
        "/Copy"
    );
    for (part, bytes) in &source_parts {
        assert_eq!(editor.package().part(part).unwrap(), bytes, "{part}");
    }

    for part in [
        "xl/worksheets/sheet2.xml",
        "xl/worksheets/_rels/sheet2.xml.rels",
        "xl/tables/table2.xml",
        "xl/comments2.xml",
        "xl/drawings/vmlDrawing2.vml",
        "xl/drawings/drawing2.xml",
        "xl/drawings/_rels/drawing2.xml.rels",
        "xl/charts/chart2.xml",
        "xl/media/image2.png",
    ] {
        assert!(editor.package().contains_part(part), "{part}");
    }
    assert!(!editor.package().contains_part("xl/theme/theme2.xml"));
    assert!(part_text(&editor, "xl/tables/table2.xml").contains("name=\"Table2\""));
    assert!(part_text(&editor, "xl/tables/table2.xml").contains("displayName=\"Table2\""));
    assert!(part_text(&editor, "xl/charts/chart2.xml").contains("&apos;Copy&apos;!$A$1:$A$2"));
    assert!(part_text(&editor, "xl/worksheets/_rels/sheet2.xml.rels")
        .contains("../drawings/drawing2.xml"));
    assert!(
        part_text(&editor, "xl/drawings/_rels/drawing2.xml.rels").contains("../media/image2.png")
    );
    let workbook = part_text(&editor, "xl/workbook.xml");
    assert!(workbook.contains("name=\"Copy\""));
    assert_eq!(workbook.matches("name=\"LocalValue\"").count(), 2);
    assert!(workbook.contains("localSheetId=\"1\""));
    assert!(workbook.contains("&apos;Copy&apos;!$A$1"));

    let snapshot = editor.snapshot().unwrap();
    assert_eq!(snapshot.get("/Copy/A2", 0).unwrap().text, "source");
    assert_eq!(
        snapshot
            .get("/Copy/A1", 0)
            .unwrap()
            .format
            .get("formula")
            .map(String::as_str),
        Some("'Copy'!A2")
    );
    editor.package().opc_model().unwrap();

    let parts_before_failure = editor.package().part_names().count();
    assert!(editor.copy_worksheet("/Sheet1", "Copy", None).is_err());
    assert_eq!(editor.package().part_names().count(), parts_before_failure);

    let mut shared_package = editor.package().clone();
    let drawing_relationships = part_text(&editor, "xl/drawings/_rels/drawing2.xml.rels")
        .replace("../media/image2.png", "../media/image1.png");
    shared_package
        .set_part(
            "xl/drawings/_rels/drawing2.xml.rels",
            drawing_relationships.into_bytes(),
        )
        .unwrap();
    crate::opc_edit::remove_content_type_override(&mut shared_package, "xl/media/image2.png")
        .unwrap();
    shared_package.remove_part("xl/media/image2.png").unwrap();
    let mut editor = NativeOfficeEditor::from_package(shared_package).unwrap();
    editor
        .set_cell_value(
            "/Sheet1/B1",
            SpreadsheetCellValue::Formula {
                expression: "'Copy'!A2".into(),
            },
        )
        .unwrap();
    editor.remove("/Copy").unwrap();
    for part in [
        "xl/worksheets/sheet2.xml",
        "xl/worksheets/_rels/sheet2.xml.rels",
        "xl/tables/table2.xml",
        "xl/comments2.xml",
        "xl/drawings/vmlDrawing2.vml",
        "xl/drawings/drawing2.xml",
        "xl/drawings/_rels/drawing2.xml.rels",
        "xl/charts/chart2.xml",
        "xl/media/image2.png",
    ] {
        assert!(!editor.package().contains_part(part), "{part}");
    }
    for (part, bytes) in &source_parts {
        if *part == "xl/worksheets/sheet1.xml" {
            continue;
        }
        assert_eq!(editor.package().part(part).unwrap(), bytes, "{part}");
    }
    assert!(!editor.package().contains_part("xl/theme/theme2.xml"));
    let snapshot = editor.snapshot().unwrap();
    assert_eq!(
        snapshot
            .get("/Sheet1/B1", 0)
            .unwrap()
            .format
            .get("formula")
            .map(String::as_str),
        Some("#REF!A2")
    );
    let workbook = part_text(&editor, "xl/workbook.xml");
    assert_eq!(workbook.matches("name=\"LocalValue\"").count(), 1);
    assert!(!workbook.contains("name=\"Copy\""));
    editor.package().opc_model().unwrap();
}

fn part_text(editor: &NativeOfficeEditor, part: &str) -> String {
    String::from_utf8(editor.package().part(part).unwrap().to_vec()).unwrap()
}
