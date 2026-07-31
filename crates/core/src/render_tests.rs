use crate::render::render_with_limit;
use crate::{
    NativeOfficeDocument, NativeOfficeEditor, NativeOfficeImage, NativeOfficeRenderFormat,
    NativeOfficeRenderedView, NativeOfficeUnit, NativeOfficeUnitInventoryOptions,
    NativeOfficeUnitLocator, NativeOfficeUnitRenderOptions, SpreadsheetCellValue,
    MAX_NATIVE_OFFICE_RENDER_BYTES, MAX_NATIVE_OFFICE_UNIT_INVENTORY_LIMIT,
};

const PNG_1X1: &[u8] = &[
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
    0x89, 0x00, 0x00, 0x00, 0x0a, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
    0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
    0x42, 0x60, 0x82,
];

#[tokio::test]
async fn word_html_and_svg_are_deterministic_escaped_and_path_free() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("secret-source-name.docx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor
        .set_text(
            "/body/p[1]",
            "</style><script>alert(\"x\")</script>& https://evil.invalid/a\u{1}",
        )
        .unwrap();
    let document = editor.snapshot().unwrap();

    let first = document.html_view().unwrap();
    let second = document.html_view().unwrap();
    let svg = document.svg_view().unwrap();
    let second_svg = document.svg_view().unwrap();

    assert_eq!(first, second);
    assert_eq!(svg, second_svg);
    assert_eq!(first.byte_length, first.content.len());
    assert_eq!(first.unit_count, 1);
    assert_eq!(first.media_type, "text/html; charset=utf-8");
    assert!(first.content.starts_with("<!doctype html>"));
    assert!(first.content.contains("Content-Security-Policy"));
    assert!(first.content.contains("data-path=\"/body/p[1]\""));
    assert!(first
        .content
        .contains("&lt;/style&gt;&lt;script&gt;alert(\"x\")&lt;/script&gt;&amp;"));
    assert!(first.content.contains('\u{fffd}'));
    assert!(!first.content.contains("<script"));
    assert!(!first.content.contains("src=\"http"));
    assert!(!first.content.contains("href=\"http"));
    assert!(!first.content.contains("secret-source-name.docx"));
    assert_eq!(svg.kind, crate::DocumentKind::Word);
    assert_eq!(svg.media_type, "image/svg+xml");
    assert!(svg.content.starts_with("<?xml version=\"1.0\""));
    assert!(svg.content.contains("data-document-kind=\"word\""));
    assert!(svg.content.contains("data-path=\"/body/p[1]\""));
    assert!(svg.content.contains("&lt;/style&gt;&lt;script&gt;"));
    assert!(svg.content.contains("&amp; https://evil.invalid/a"));
    assert!(!svg.content.contains("<script"));
    assert!(!svg.content.contains("href=\"http"));
    assert!(!svg.content.contains("secret-source-name.docx"));
    assert_well_formed_xml(&svg.content);
}

#[tokio::test]
async fn spreadsheet_html_and_svg_keep_observed_cells_sparse() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("sparse.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor.set_text("/Sheet1/A1", "first").unwrap();
    editor
        .set_cell_value(
            "/Sheet1/B2",
            SpreadsheetCellValue::Formula {
                expression: "SUM(A1:A2)".into(),
            },
        )
        .unwrap();
    editor.set_text("/Sheet1/XFD1048576", "edge").unwrap();

    let document = editor.snapshot().unwrap();
    let rendered = document.html_view().unwrap();
    let svg = document.svg_view().unwrap();

    assert_eq!(rendered.unit_count, 1);
    assert!(rendered.content.contains("A1"));
    assert!(rendered.content.contains("XFD1048576"));
    assert_eq!(rendered.content.matches("class=\"cell\"").count(), 3);
    assert!(rendered.byte_length < 32 * 1024);
    assert_eq!(svg.unit_count, 1);
    assert!(svg.content.contains("data-document-kind=\"spreadsheet\""));
    assert!(svg.content.contains("data-path=\"/Sheet1/A1\""));
    assert!(svg.content.contains("data-path=\"/Sheet1/XFD1048576\""));
    assert_eq!(svg.content.matches("class=\"cell\"").count(), 3);
    assert!(svg.content.contains("data-formula=\"SUM(A1:A2)\""));
    assert!(svg.content.contains("=SUM(A1:A2)"));
    assert!(svg.byte_length < 32 * 1024);
    assert_well_formed_xml(&svg.content);
}

#[tokio::test]
async fn presentation_html_and_svg_render_semantics_and_validated_images() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("deck.pptx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    let slide = editor.add_slide("/", "A3S <Native>").unwrap();
    editor.add_shape(&slide, "Body & details").unwrap();
    editor
        .add_image(
            &slide,
            NativeOfficeImage::from_bytes(PNG_1X1)
                .unwrap()
                .with_alt_text("Logo <safe>"),
        )
        .unwrap();
    let document = editor.snapshot().unwrap();

    let html = document.html_view().unwrap();
    let svg = document.svg_view().unwrap();

    assert_eq!(html.unit_count, 1);
    assert_eq!(svg.unit_count, 1);
    assert_eq!(svg.media_type, "image/svg+xml");
    assert!(html.content.contains("data:image/png;base64,"));
    assert!(svg.content.contains("data:image/png;base64,"));
    assert!(html.content.contains("A3S &lt;Native&gt;"));
    assert!(svg.content.contains("A3S &lt;Native&gt;"));
    assert!(svg.content.contains("data-path=\"/slide[1]\""));
    assert!(svg.content.starts_with("<?xml version=\"1.0\""));
    assert!(!html.content.contains("<script"));
    assert!(!svg.content.contains("<script"));
    assert!(!svg.content.contains("href=\"http"));
    assert_well_formed_xml(&svg.content);
}

#[tokio::test]
async fn renderer_enforces_its_bound_while_composing_output() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("bounded.docx");
    let document = NativeOfficeEditor::create(&path)
        .await
        .unwrap()
        .snapshot()
        .unwrap();

    let error = render_with_limit(&document, NativeOfficeRenderFormat::Html, 128).unwrap_err();
    assert_eq!(error.code, "use.office.render_output_too_large");
    assert_eq!(error.details["limitBytes"], 128);
    let error = render_with_limit(&document, NativeOfficeRenderFormat::Svg, 128).unwrap_err();
    assert_eq!(error.code, "use.office.render_output_too_large");
    assert_eq!(error.details["limitBytes"], 128);
}

#[tokio::test]
async fn word_and_spreadsheet_svg_embed_only_validated_internal_images() {
    let temp = tempfile::tempdir().unwrap();
    for (name, parent) in [("image.docx", "/body"), ("image.xlsx", "/Sheet1/A1")] {
        let path = temp.path().join(name);
        let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
        let created = editor
            .add_image(
                parent,
                NativeOfficeImage::from_bytes(PNG_1X1)
                    .unwrap()
                    .with_alt_text("Validated <image>"),
            )
            .unwrap();

        let svg = editor.snapshot().unwrap().svg_view().unwrap();
        assert!(svg.content.contains("data:image/png;base64,"));
        assert!(svg.content.contains("Validated &lt;image&gt;"));
        assert!(svg
            .content
            .contains(&format!("data-path=\"{}\"", created.path)));
        assert!(!svg.content.contains("href=\"http"));
        assert_well_formed_xml(&svg.content);
    }
}

#[tokio::test]
async fn renderer_rejects_corrupt_internal_raster_parts() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("corrupt-image.docx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    let image = editor
        .add_image("/body", NativeOfficeImage::from_bytes(PNG_1X1).unwrap())
        .unwrap();
    let mut package = editor.package().clone();
    package
        .set_part(&image.part, b"not an image".to_vec())
        .unwrap();
    let document = NativeOfficeDocument::from_package(package).unwrap();

    for format in [
        NativeOfficeRenderFormat::Html,
        NativeOfficeRenderFormat::Svg,
    ] {
        let error = document.render(format).unwrap_err();
        assert_eq!(error.code, "use.office.render_image_invalid");
        assert_eq!(error.details["part"], image.part);
    }
}

#[tokio::test]
async fn renderer_never_fetches_or_emits_external_image_relationships() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("external-image.docx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    let image = editor
        .add_image("/body", NativeOfficeImage::from_bytes(PNG_1X1).unwrap())
        .unwrap();
    let mut package = editor.package().clone();
    let relationships_part = "word/_rels/document.xml.rels";
    let relationships =
        String::from_utf8(package.part(relationships_part).unwrap().to_vec()).unwrap();
    let internal_target = image
        .part
        .trim_start_matches('/')
        .strip_prefix("word/")
        .unwrap();
    let original = format!("Target=\"{internal_target}\"");
    assert!(relationships.contains(&original));
    let external = relationships.replacen(
        &original,
        "Target=\"https://example.invalid/pixel.png\" TargetMode=\"External\"",
        1,
    );
    package
        .set_part(relationships_part, external.into_bytes())
        .unwrap();
    let document = NativeOfficeDocument::from_package(package).unwrap();

    let html = document.html_view().unwrap();
    let svg = document.svg_view().unwrap();
    assert!(html.content.contains("Embedded image unavailable"));
    assert!(!html.content.contains("example.invalid"));
    assert!(!html.content.contains("src=\"http"));
    assert!(svg.content.contains("Embedded image unavailable"));
    assert!(!svg.content.contains("example.invalid"));
    assert!(!svg.content.contains("href=\"http"));
}

#[tokio::test]
async fn unit_inventory_preserves_exact_word_worksheet_and_slide_identity() {
    let temp = tempfile::tempdir().unwrap();

    let word_path = temp.path().join("report.docx");
    let word = NativeOfficeEditor::create(&word_path)
        .await
        .unwrap()
        .snapshot()
        .unwrap();
    let inventory = word
        .inventory_units(NativeOfficeUnitInventoryOptions::default())
        .unwrap();
    assert_eq!(inventory.total_units, 1);
    assert_eq!(
        inventory.units,
        vec![NativeOfficeUnit {
            ordinal: 1,
            locator: NativeOfficeUnitLocator::Document,
            path: "/".to_string(),
        }]
    );

    let workbook_path = temp.path().join("workbook.xlsx");
    let mut workbook = NativeOfficeEditor::create(&workbook_path).await.unwrap();
    workbook.add_worksheet("Data Set").unwrap();
    let workbook = workbook.snapshot().unwrap();
    let inventory = workbook
        .inventory_units(NativeOfficeUnitInventoryOptions::default())
        .unwrap();
    assert_eq!(inventory.total_units, 2);
    assert_eq!(
        inventory.units,
        vec![
            NativeOfficeUnit {
                ordinal: 1,
                locator: NativeOfficeUnitLocator::Worksheet {
                    index: 1,
                    name: "Sheet1".to_string(),
                },
                path: "/Sheet1".to_string(),
            },
            NativeOfficeUnit {
                ordinal: 2,
                locator: NativeOfficeUnitLocator::Worksheet {
                    index: 2,
                    name: "Data Set".to_string(),
                },
                path: "/Data Set".to_string(),
            },
        ]
    );

    let deck_path = temp.path().join("deck.pptx");
    let mut deck = NativeOfficeEditor::create(&deck_path).await.unwrap();
    deck.add_slide("/", "First").unwrap();
    deck.add_slide("/", "Second").unwrap();
    let deck = deck.snapshot().unwrap();
    let inventory = deck
        .inventory_units(NativeOfficeUnitInventoryOptions::default())
        .unwrap();
    assert_eq!(inventory.total_units, 2);
    assert_eq!(
        inventory.units,
        vec![
            NativeOfficeUnit {
                ordinal: 1,
                locator: NativeOfficeUnitLocator::Slide { number: 1 },
                path: "/slide[1]".to_string(),
            },
            NativeOfficeUnit {
                ordinal: 2,
                locator: NativeOfficeUnitLocator::Slide { number: 2 },
                path: "/slide[2]".to_string(),
            },
        ]
    );
}

#[tokio::test]
async fn per_unit_html_and_svg_exclude_sibling_worksheets_and_slides() {
    let temp = tempfile::tempdir().unwrap();

    let workbook_path = temp.path().join("isolated.xlsx");
    let mut workbook = NativeOfficeEditor::create(&workbook_path).await.unwrap();
    workbook
        .set_text("/Sheet1/A1", "FIRST-SHEET-MARKER")
        .unwrap();
    workbook.add_worksheet("Data").unwrap();
    workbook
        .set_text("/Data/A1", "SECOND-SHEET-MARKER")
        .unwrap();
    let workbook = workbook.snapshot().unwrap();
    let worksheet = NativeOfficeUnitLocator::Worksheet {
        index: 2,
        name: "Data".to_string(),
    };
    for format in [
        NativeOfficeRenderFormat::Html,
        NativeOfficeRenderFormat::Svg,
    ] {
        let first = workbook
            .render_unit(
                &worksheet,
                NativeOfficeUnitRenderOptions {
                    format,
                    max_output_bytes: MAX_NATIVE_OFFICE_RENDER_BYTES,
                },
            )
            .unwrap();
        let second = workbook
            .render_unit(
                &worksheet,
                NativeOfficeUnitRenderOptions {
                    format,
                    max_output_bytes: MAX_NATIVE_OFFICE_RENDER_BYTES,
                },
            )
            .unwrap();
        assert_eq!(first, second);
        assert_eq!(first.document_unit_count, 2);
        assert_eq!(first.unit.path, "/Data");
        assert_eq!(first.byte_length, first.content.len());
        assert!(first.content.contains("SECOND-SHEET-MARKER"));
        if format == NativeOfficeRenderFormat::Html {
            assert!(first.content.contains(
                "data-document-kind=\"spreadsheet\" data-unit-path=\"/Data\" data-unit-ordinal=\"2\">"
            ));
        }
        assert!(!first.content.contains("FIRST-SHEET-MARKER"));
        assert!(!first.content.contains("data-path=\"/Sheet1\""));
    }

    let deck_path = temp.path().join("isolated.pptx");
    let mut deck = NativeOfficeEditor::create(&deck_path).await.unwrap();
    deck.add_slide("/", "FIRST-SLIDE-MARKER").unwrap();
    deck.add_slide("/", "SECOND-SLIDE-MARKER").unwrap();
    let deck = deck.snapshot().unwrap();
    let slide = NativeOfficeUnitLocator::Slide { number: 2 };
    for format in [
        NativeOfficeRenderFormat::Html,
        NativeOfficeRenderFormat::Svg,
    ] {
        let rendered = deck
            .render_unit(
                &slide,
                NativeOfficeUnitRenderOptions {
                    format,
                    max_output_bytes: MAX_NATIVE_OFFICE_RENDER_BYTES,
                },
            )
            .unwrap();
        assert_eq!(rendered.document_unit_count, 2);
        assert_eq!(rendered.unit.path, "/slide[2]");
        assert!(rendered.content.contains("SECOND-SLIDE-MARKER"));
        assert!(!rendered.content.contains("FIRST-SLIDE-MARKER"));
        assert!(!rendered.content.contains("data-path=\"/slide[1]\""));
    }
}

#[tokio::test]
async fn unit_contract_rejects_wrong_duplicate_and_over_limit_identity() {
    let temp = tempfile::tempdir().unwrap();
    let workbook_path = temp.path().join("identity.xlsx");
    let mut editor = NativeOfficeEditor::create(&workbook_path).await.unwrap();
    editor.add_worksheet("Data").unwrap();
    let document = editor.snapshot().unwrap();

    let error = document
        .render_unit(
            &NativeOfficeUnitLocator::Worksheet {
                index: 1,
                name: "Data".to_string(),
            },
            NativeOfficeUnitRenderOptions::default(),
        )
        .unwrap_err();
    assert_eq!(error.code, "use.office.unit_identity_mismatch");

    let error = document
        .render_unit(
            &NativeOfficeUnitLocator::Worksheet {
                index: 0,
                name: "Sheet1".to_string(),
            },
            NativeOfficeUnitRenderOptions::default(),
        )
        .unwrap_err();
    assert_eq!(error.code, "use.office.unit_locator_invalid");

    let error = document
        .render_unit(
            &NativeOfficeUnitLocator::Document,
            NativeOfficeUnitRenderOptions::default(),
        )
        .unwrap_err();
    assert_eq!(error.code, "use.office.unit_kind_mismatch");

    let deck_path = temp.path().join("identity.pptx");
    let mut deck = NativeOfficeEditor::create(&deck_path).await.unwrap();
    deck.add_slide("/", "Only slide").unwrap();
    let error = deck
        .snapshot()
        .unwrap()
        .render_unit(
            &NativeOfficeUnitLocator::Slide { number: 2 },
            NativeOfficeUnitRenderOptions::default(),
        )
        .unwrap_err();
    assert_eq!(error.code, "use.office.unit_not_found");

    let error = document
        .inventory_units(NativeOfficeUnitInventoryOptions { max_units: 1 })
        .unwrap_err();
    assert_eq!(error.code, "use.office.unit_inventory_too_large");
    assert_eq!(error.details["totalUnits"], 2);
    assert_eq!(error.details["maxUnits"], 1);

    for max_units in [0, MAX_NATIVE_OFFICE_UNIT_INVENTORY_LIMIT + 1] {
        let error = document
            .inventory_units(NativeOfficeUnitInventoryOptions { max_units })
            .unwrap_err();
        assert_eq!(error.code, "use.office.unit_inventory_limit_invalid");
    }

    let mut package = editor.package().clone();
    let workbook_xml =
        String::from_utf8(package.part("xl/workbook.xml").unwrap().to_vec()).unwrap();
    assert!(workbook_xml.contains("name=\"Data\""));
    package
        .set_part(
            "xl/workbook.xml",
            workbook_xml
                .replacen("name=\"Data\"", "name=\"Sheet1\"", 1)
                .into_bytes(),
        )
        .unwrap();
    let duplicate = NativeOfficeDocument::from_package(package).unwrap();
    let error = duplicate
        .inventory_units(NativeOfficeUnitInventoryOptions::default())
        .unwrap_err();
    assert_eq!(error.code, "use.office.unit_identity_duplicate");
}

#[tokio::test]
async fn per_unit_render_has_an_explicit_byte_bound_and_leaks_no_source_path() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("confidential-source-name.docx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor
        .set_text(
            "/body/p[1]",
            "safe text with https://example.invalid/not-a-fetch",
        )
        .unwrap();
    let document = editor.snapshot().unwrap();

    let rendered = document
        .render_unit(
            &NativeOfficeUnitLocator::Document,
            NativeOfficeUnitRenderOptions::default(),
        )
        .unwrap();
    assert_eq!(rendered.unit.path, "/");
    assert_eq!(rendered.document_unit_count, 1);
    assert!(!rendered.content.contains("confidential-source-name.docx"));
    assert!(!rendered.content.contains("src=\"http"));
    assert!(!rendered.content.contains("href=\"http"));

    let error = document
        .render_unit(
            &NativeOfficeUnitLocator::Document,
            NativeOfficeUnitRenderOptions {
                format: NativeOfficeRenderFormat::Html,
                max_output_bytes: 128,
            },
        )
        .unwrap_err();
    assert_eq!(error.code, "use.office.render_output_too_large");
    assert_eq!(error.details["limitBytes"], 128);

    for max_output_bytes in [0, MAX_NATIVE_OFFICE_RENDER_BYTES + 1] {
        let error = document
            .render_unit(
                &NativeOfficeUnitLocator::Document,
                NativeOfficeUnitRenderOptions {
                    format: NativeOfficeRenderFormat::Html,
                    max_output_bytes,
                },
            )
            .unwrap_err();
        assert_eq!(error.code, "use.office.render_limit_invalid");
    }
}

#[test]
fn unit_locator_json_is_stable_and_strict() {
    let locator = NativeOfficeUnitLocator::Worksheet {
        index: 2,
        name: "Data".to_string(),
    };
    assert_eq!(
        serde_json::to_value(&locator).unwrap(),
        serde_json::json!({ "type": "worksheet", "index": 2, "name": "Data" })
    );
    assert!(
        serde_json::from_value::<NativeOfficeUnitLocator>(serde_json::json!({
            "type": "slide",
            "number": 2,
            "unexpected": true
        }))
        .is_err()
    );
}

#[test]
fn public_render_contracts_are_send_and_sync() {
    fn assert_send_sync<T: Send + Sync>() {}

    assert_send_sync::<NativeOfficeRenderFormat>();
    assert_send_sync::<NativeOfficeRenderedView>();
    assert_send_sync::<NativeOfficeUnitLocator>();
    assert_send_sync::<NativeOfficeUnit>();
    assert_send_sync::<NativeOfficeUnitInventoryOptions>();
    assert_send_sync::<NativeOfficeUnitRenderOptions>();
    assert_send_sync::<crate::NativeOfficeUnitInventory>();
    assert_send_sync::<crate::NativeOfficeRenderedUnit>();
}

fn assert_well_formed_xml(xml: &str) {
    let mut reader = quick_xml::Reader::from_str(xml);
    loop {
        if matches!(reader.read_event().unwrap(), quick_xml::events::Event::Eof) {
            break;
        }
    }
}
