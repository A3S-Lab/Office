use crate::{
    NativeOfficeLayoutEnvironment, NativeOfficeLayoutRenderer, NativeOfficeLayoutSourceKind,
    NativeOfficePdfOutline, NativeOfficePdfOutlineEntry, NativeOfficePdfOutlineOptions,
    NativeOfficePdfPageBox, NativeOfficePdfPageGeometry, NativeOfficePdfPageInventory,
    NativeOfficePdfPageInventoryOptions, NativeOfficePdfPageTextLayer,
    NativeOfficePdfTextCharacter, NativeOfficePdfTextLayerOptions,
    NativeOfficePdfiumLayoutRenderer, NativeOfficeUnit, NativeOfficeUnitLocator, PackageRevision,
    NATIVE_OFFICE_PDF_TEXT_SCHEMA_VERSION,
};
use sha2::Digest as _;

const DPI_MILLI: u32 = 144_000;
const TEST_TIMEOUT_MS: u64 = 10_000;
const TEST_MAX_OUTPUT_BYTES: u64 = 1024 * 1024;

#[test]
fn pdf_page_locator_has_a_strict_one_based_json_contract() {
    let locator = NativeOfficeUnitLocator::Page { number: 2 };
    assert_eq!(
        serde_json::to_value(&locator).unwrap(),
        serde_json::json!({ "type": "page", "number": 2 })
    );
    assert_eq!(
        serde_json::from_value::<NativeOfficeUnitLocator>(
            serde_json::json!({ "type": "page", "number": 2 })
        )
        .unwrap(),
        locator
    );
    assert!(serde_json::from_value::<NativeOfficeUnitLocator>(
        serde_json::json!({ "type": "page", "number": 2, "unchecked": true })
    )
    .is_err());
}

#[test]
fn two_page_pdf_inventory_is_complete_ordered_and_source_bound() {
    let inventory = inventory(vec![
        page(1, 0, 612_000, 792_000),
        page(2, 90, 792_000, 612_000),
    ]);

    inventory.validate().unwrap();
    assert_eq!(inventory.kind, NativeOfficeLayoutSourceKind::Pdf);
    assert_eq!(inventory.pages[0].unit.path, "/page[1]");
    assert_eq!(inventory.pages[1].unit.path, "/page[2]");
    assert_eq!(inventory.pages[0].output_width_px, 1_224);
    assert_eq!(inventory.pages[0].output_height_px, 1_584);
    assert_eq!(inventory.pages[1].rotation_degrees, 90);
}

#[test]
fn pdf_inventory_rejects_truncation_reordering_duplicates_and_foreign_sources() {
    let mut truncated = inventory(vec![page(1, 0, 612_000, 792_000)]);
    truncated.total_pages = 2;
    assert_eq!(
        truncated.validate().unwrap_err().code,
        "use.office.pdf_inventory_invalid"
    );

    let reordered = inventory(vec![
        page(2, 0, 612_000, 792_000),
        page(1, 0, 612_000, 792_000),
    ]);
    assert_eq!(
        reordered.validate().unwrap_err().code,
        "use.office.pdf_inventory_invalid"
    );

    let duplicate = inventory(vec![
        page(1, 0, 612_000, 792_000),
        page(1, 0, 612_000, 792_000),
    ]);
    assert_eq!(
        duplicate.validate().unwrap_err().code,
        "use.office.pdf_inventory_invalid"
    );

    let observed = inventory(vec![page(1, 0, 612_000, 792_000)]);
    let foreign = PackageRevision {
        archive_bytes: 24,
        sha256: "b".repeat(64),
    };
    assert_eq!(
        observed.validate_source(&foreign).unwrap_err().code,
        "use.office.pdf_inventory_source_mismatch"
    );
}

#[test]
fn pdf_inventory_rejects_invalid_page_geometry_and_limits() {
    let mut invalid_rotation = inventory(vec![page(1, 45, 612_000, 792_000)]);
    assert_eq!(
        invalid_rotation.validate().unwrap_err().code,
        "use.office.pdf_page_geometry_invalid"
    );

    invalid_rotation.pages[0].rotation_degrees = 0;
    invalid_rotation.pages[0].crop_box.right_millipoints = 0;
    assert_eq!(
        invalid_rotation.validate().unwrap_err().code,
        "use.office.pdf_page_geometry_invalid"
    );

    let mut over_limit = inventory(vec![
        page(1, 0, 612_000, 792_000),
        page(2, 0, 612_000, 792_000),
    ]);
    over_limit.max_pages = 1;
    assert_eq!(
        over_limit.validate().unwrap_err().code,
        "use.office.pdf_page_limit"
    );
}

#[test]
fn pdf_text_layer_preserves_unicode_offsets_geometry_and_source_identity() {
    let inventory = inventory(vec![page(1, 0, 612_000, 792_000)]);
    let text = "A\n🌍".to_string();
    let layer = NativeOfficePdfPageTextLayer {
        schema_version: NATIVE_OFFICE_PDF_TEXT_SCHEMA_VERSION,
        source_revision: inventory.source_revision.clone(),
        unit: inventory.pages[0].unit.clone(),
        page_geometry: inventory.pages[0].clone(),
        engine_version: "chromium/7881".to_string(),
        max_characters: 8,
        max_text_bytes: 32,
        text_sha256: format!("{:x}", sha2::Sha256::digest(text.as_bytes())),
        text,
        characters: vec![
            text_character(
                0,
                "A",
                0,
                1,
                0,
                1,
                Some(text_box(10_000, 20_000, 18_000, 32_000)),
            ),
            text_character(1, "\n", 1, 2, 1, 2, None),
            text_character(
                2,
                "🌍",
                2,
                6,
                2,
                4,
                Some(text_box(10_000, 40_000, 22_000, 52_000)),
            ),
        ],
    };

    layer.validate(&inventory).unwrap();
    assert_eq!(layer.characters[2].utf8_range(), 2..6);
    assert_eq!(layer.characters[2].utf16_range(), 2..4);

    let mut tampered = layer.clone();
    tampered.characters[2].utf16_end = 3;
    assert_eq!(
        tampered.validate(&inventory).unwrap_err().code,
        "use.office.pdf_text_layer_invalid"
    );

    let mut foreign = inventory.clone();
    foreign.source_revision.sha256 = "b".repeat(64);
    assert_eq!(
        layer.validate(&foreign).unwrap_err().code,
        "use.office.pdf_text_layer_source_mismatch"
    );
}

#[test]
fn pdf_outline_preserves_hierarchy_and_exact_page_targets() {
    let inventory = inventory(vec![
        page(1, 0, 612_000, 792_000),
        page(2, 0, 612_000, 792_000),
    ]);
    let outline = NativeOfficePdfOutline {
        schema_version: NATIVE_OFFICE_PDF_TEXT_SCHEMA_VERSION,
        source_revision: inventory.source_revision.clone(),
        engine_version: "chromium/7881".to_string(),
        max_entries: 8,
        max_depth: 4,
        max_title_bytes: 64,
        entries: vec![
            NativeOfficePdfOutlineEntry {
                index: 0,
                parent_index: None,
                depth: 0,
                title: "Chapter".to_string(),
                target_unit: Some(inventory.pages[0].unit.clone()),
            },
            NativeOfficePdfOutlineEntry {
                index: 1,
                parent_index: Some(0),
                depth: 1,
                title: "Section".to_string(),
                target_unit: Some(inventory.pages[1].unit.clone()),
            },
        ],
    };

    outline.validate(&inventory).unwrap();

    let mut invalid_parent = outline.clone();
    invalid_parent.entries[1].parent_index = Some(1);
    assert_eq!(
        invalid_parent.validate(&inventory).unwrap_err().code,
        "use.office.pdf_outline_invalid"
    );
}

#[tokio::test]
async fn explicit_pdfium_library_inventories_and_renders_exact_pages() {
    let Some(library) = std::env::var_os("A3S_OFFICE_TEST_PDFIUM_LIBRARY") else {
        return;
    };
    let directory = tempfile::tempdir().unwrap();
    let source = directory.path().join("two-pages.pdf");
    std::fs::write(&source, two_page_pdf()).unwrap();
    let renderer = NativeOfficePdfiumLayoutRenderer::from_library(
        library,
        format!(
            "{:x}",
            sha2::Sha256::digest(b"a3s-office-test-font-manifest-v1")
        ),
    )
    .await
    .unwrap();
    assert!(renderer.supports(NativeOfficeLayoutSourceKind::Pdf));
    assert!(!renderer.supports(NativeOfficeLayoutSourceKind::Presentation));

    let revision = renderer
        .source_revision(&source, 1024 * 1024, TEST_TIMEOUT_MS)
        .await
        .unwrap();
    let options = NativeOfficePdfPageInventoryOptions {
        max_pages: 2,
        max_source_bytes: 1024 * 1024,
        dpi_milli: 72_000,
        timeout_ms: TEST_TIMEOUT_MS,
    };
    let inventory = renderer
        .inventory_pages(&source, revision.clone(), options)
        .await
        .unwrap();
    assert_eq!(renderer.inventory_call_count(), 1);
    assert_eq!(inventory.total_pages, 2);
    assert_eq!(inventory.pages[0].media_box.right_millipoints, 200_000);
    assert_eq!(inventory.pages[0].crop_box.left_millipoints, 10_000);
    assert_eq!(inventory.pages[0].crop_box.bottom_millipoints, 20_000);
    assert_eq!(inventory.pages[0].output_width_px, 180);
    assert_eq!(inventory.pages[0].output_height_px, 60);
    assert_eq!(inventory.pages[1].rotation_degrees, 90);
    assert_eq!(inventory.pages[1].output_width_px, 240);
    assert_eq!(inventory.pages[1].output_height_px, 120);

    let first_text = renderer
        .extract_page_text(
            &source,
            &inventory,
            inventory.pages[0].unit.clone(),
            NativeOfficePdfTextLayerOptions {
                max_characters: 64,
                max_text_bytes: 1024,
                timeout_ms: TEST_TIMEOUT_MS,
            },
        )
        .await
        .unwrap();
    assert_eq!(renderer.inventory_call_count(), 1);
    assert_eq!(first_text.text, "Hello PDF");
    assert_eq!(first_text.characters.len(), 9);
    assert!(first_text
        .characters
        .iter()
        .filter_map(|character| character.bounds)
        .all(|bounds| bounds.right_millipoints > bounds.left_millipoints));
    first_text.validate(&inventory).unwrap();

    let outline = renderer
        .extract_outline(
            &source,
            &inventory,
            NativeOfficePdfOutlineOptions {
                max_entries: 4,
                max_depth: 4,
                max_title_bytes: 64,
                timeout_ms: TEST_TIMEOUT_MS,
            },
        )
        .await
        .unwrap();
    assert_eq!(renderer.inventory_call_count(), 1);
    assert_eq!(outline.entries.len(), 2);
    assert_eq!(outline.entries[0].title, "First page");
    assert_eq!(
        outline.entries[0].target_unit,
        Some(inventory.pages[0].unit.clone())
    );
    assert_eq!(outline.entries[1].title, "Second page");
    assert_eq!(
        outline.entries[1].target_unit,
        Some(inventory.pages[1].unit.clone())
    );

    let text_limit = renderer
        .extract_page_text(
            &source,
            &inventory,
            inventory.pages[0].unit.clone(),
            NativeOfficePdfTextLayerOptions {
                max_characters: 1,
                max_text_bytes: 1024,
                timeout_ms: TEST_TIMEOUT_MS,
            },
        )
        .await
        .unwrap_err();
    assert_eq!(text_limit.code, "use.office.pdf_text_character_limit");

    let outline_limit = renderer
        .extract_outline(
            &source,
            &inventory,
            NativeOfficePdfOutlineOptions {
                max_entries: 1,
                max_depth: 4,
                max_title_bytes: 64,
                timeout_ms: TEST_TIMEOUT_MS,
            },
        )
        .await
        .unwrap_err();
    assert_eq!(outline_limit.code, "use.office.pdf_outline_entry_limit");

    let page_one = inventory.pages[0].unit.clone();
    let inspection = renderer
        .inspect_inventoried_page(
            &source,
            &inventory,
            page_one.clone(),
            NativeOfficeLayoutEnvironment::new("en-US", "UTC"),
        )
        .unwrap();
    assert_eq!(renderer.inventory_call_count(), 1);
    let first_output = directory.path().join("page-1.png");
    let first_request = inspection.clone().into_render_request(
        &first_output,
        TEST_MAX_OUTPUT_BYTES,
        TEST_TIMEOUT_MS,
    );
    let first = renderer.render(first_request.clone()).await.unwrap();
    assert_eq!(first.unit, page_one);
    assert_eq!(first.raster.rotation_degrees, 0);
    assert_eq!((first.raster.width_px, first.raster.height_px), (180, 60));
    assert_eq!(
        rgba_pixel(&std::fs::read(&first_output).unwrap(), 90, 30),
        [255, 0, 0, 255]
    );

    let no_clobber = renderer.render(first_request).await.unwrap_err();
    assert_eq!(no_clobber.code, "use.office.layout_output_exists");
    assert_eq!(first.raster.sha256, sha256_file(&first_output));

    let repeat_output = directory.path().join("page-1-repeat.png");
    let repeat = renderer
        .render(inspection.into_render_request(
            &repeat_output,
            TEST_MAX_OUTPUT_BYTES,
            TEST_TIMEOUT_MS,
        ))
        .await
        .unwrap();
    assert_eq!(first.profile_sha256, repeat.profile_sha256);
    assert_eq!(first.raster.sha256, repeat.raster.sha256);

    let page_two_inspection = renderer
        .inspect_inventoried_page(
            &source,
            &inventory,
            inventory.pages[1].unit.clone(),
            NativeOfficeLayoutEnvironment::new("en-US", "UTC"),
        )
        .unwrap();
    assert_eq!(renderer.inventory_call_count(), 1);
    let second_output = directory.path().join("page-2.png");
    let second = renderer
        .render(page_two_inspection.into_render_request(
            &second_output,
            TEST_MAX_OUTPUT_BYTES,
            TEST_TIMEOUT_MS,
        ))
        .await
        .unwrap();
    assert_eq!(second.raster.rotation_degrees, 90);
    assert_eq!(
        (second.raster.width_px, second.raster.height_px),
        (240, 120)
    );
    assert_eq!(
        rgba_pixel(&std::fs::read(&second_output).unwrap(), 120, 60),
        [0, 0, 255, 255]
    );
    assert_ne!(first.raster.sha256, second.raster.sha256);

    let mut tampered_inventory = inventory.clone();
    tampered_inventory.pages[0].output_width_px += 1;
    let tampered_page = renderer
        .inspect_inventoried_page(
            &source,
            &tampered_inventory,
            page_one.clone(),
            NativeOfficeLayoutEnvironment::new("en-US", "UTC"),
        )
        .unwrap_err();
    assert_eq!(tampered_page.code, "use.office.pdf_page_geometry_invalid");
    assert_eq!(renderer.inventory_call_count(), 1);

    let missing_cached_page = renderer
        .inspect_inventoried_page(
            &source,
            &inventory,
            NativeOfficeUnit {
                ordinal: 3,
                locator: NativeOfficeUnitLocator::Page { number: 3 },
                path: "/page[3]".to_string(),
            },
            NativeOfficeLayoutEnvironment::new("en-US", "UTC"),
        )
        .unwrap_err();
    assert_eq!(
        missing_cached_page.code,
        "use.office.pdf_page_identity_mismatch"
    );
    assert_eq!(renderer.inventory_call_count(), 1);

    let over_limit = renderer
        .inventory_pages(
            &source,
            revision.clone(),
            NativeOfficePdfPageInventoryOptions {
                max_pages: 1,
                ..options
            },
        )
        .await
        .unwrap_err();
    assert_eq!(over_limit.code, "use.office.pdf_page_limit");

    let invalid_zero = renderer
        .inspect_page(
            &source,
            revision.clone(),
            NativeOfficeUnit {
                ordinal: 0,
                locator: NativeOfficeUnitLocator::Page { number: 0 },
                path: "/page[0]".to_string(),
            },
            NativeOfficeLayoutEnvironment::new("en-US", "UTC"),
            options,
        )
        .await
        .unwrap_err();
    assert_eq!(invalid_zero.code, "use.office.layout_unit_mismatch");

    let missing_page = renderer
        .inspect_page(
            &source,
            revision.clone(),
            NativeOfficeUnit {
                ordinal: 3,
                locator: NativeOfficeUnitLocator::Page { number: 3 },
                path: "/page[3]".to_string(),
            },
            NativeOfficeLayoutEnvironment::new("en-US", "UTC"),
            options,
        )
        .await
        .unwrap_err();
    assert_eq!(missing_page.code, "use.office.pdf_page_identity_mismatch");

    let bounded_inspection = renderer
        .inspect_page(
            &source,
            revision.clone(),
            inventory.pages[0].unit.clone(),
            NativeOfficeLayoutEnvironment::new("en-US", "UTC"),
            options,
        )
        .await
        .unwrap();
    let bounded_output = directory.path().join("bounded.png");
    let bounded_error = renderer
        .render(bounded_inspection.into_render_request(&bounded_output, 1, TEST_TIMEOUT_MS))
        .await
        .unwrap_err();
    assert_eq!(bounded_error.code, "use.office.layout_output_too_large");
    assert!(!bounded_output.exists());

    let timeout_inspection = renderer
        .inspect_page(
            &source,
            revision.clone(),
            inventory.pages[0].unit.clone(),
            NativeOfficeLayoutEnvironment::new("en-US", "UTC"),
            options,
        )
        .await
        .unwrap();
    let timeout_output = directory.path().join("timeout.png");
    let timeout_error = renderer
        .render_with_before_final_source_check(
            timeout_inspection.into_render_request(&timeout_output, TEST_MAX_OUTPUT_BYTES, 1),
            || std::thread::sleep(std::time::Duration::from_millis(25)),
        )
        .await
        .unwrap_err();
    assert_eq!(timeout_error.code, "use.office.layout_timeout");
    assert!(!timeout_output.exists());

    let corrupt_source = directory.path().join("corrupt.pdf");
    std::fs::write(&corrupt_source, b"%PDF-1.7\nnot-a-document\n%%EOF\n").unwrap();
    let corrupt_revision = renderer
        .source_revision(&corrupt_source, 1024 * 1024, TEST_TIMEOUT_MS)
        .await
        .unwrap();
    let corrupt = renderer
        .inventory_pages(&corrupt_source, corrupt_revision, options)
        .await
        .unwrap_err();
    assert_eq!(corrupt.code, "use.office.pdf_corrupt");

    let zero_source = directory.path().join("zero-pages.pdf");
    std::fs::write(&zero_source, zero_page_pdf()).unwrap();
    let zero_revision = renderer
        .source_revision(&zero_source, 1024 * 1024, TEST_TIMEOUT_MS)
        .await
        .unwrap();
    let zero = renderer
        .inventory_pages(&zero_source, zero_revision, options)
        .await
        .unwrap_err();
    assert_eq!(zero.code, "use.office.pdf_zero_pages");

    let mutation_inspection = renderer
        .inspect_page(
            &source,
            revision,
            inventory.pages[0].unit.clone(),
            NativeOfficeLayoutEnvironment::new("en-US", "UTC"),
            options,
        )
        .await
        .unwrap();
    let mutation_output = directory.path().join("mutated.png");
    let mutation_request = mutation_inspection.into_render_request(
        &mutation_output,
        TEST_MAX_OUTPUT_BYTES,
        TEST_TIMEOUT_MS,
    );
    let source_for_hook = source.clone();
    let mutation = renderer
        .render_with_before_final_source_check(mutation_request, move || {
            use std::io::Write as _;
            std::fs::OpenOptions::new()
                .append(true)
                .open(source_for_hook)
                .unwrap()
                .write_all(b"mutated")
                .unwrap();
        })
        .await
        .unwrap_err();
    assert_eq!(mutation.code, "use.office.layout_source_mutated");
    assert!(!mutation_output.exists());

    let blank_source = directory.path().join("blank.pdf");
    std::fs::write(&blank_source, blank_page_pdf()).unwrap();
    let blank_revision = renderer
        .source_revision(&blank_source, 1024 * 1024, TEST_TIMEOUT_MS)
        .await
        .unwrap();
    let blank_inventory = renderer
        .inventory_pages(
            &blank_source,
            blank_revision,
            NativeOfficePdfPageInventoryOptions {
                max_pages: 1,
                ..options
            },
        )
        .await
        .unwrap();
    let blank_text = renderer
        .extract_page_text(
            &blank_source,
            &blank_inventory,
            blank_inventory.pages[0].unit.clone(),
            NativeOfficePdfTextLayerOptions {
                max_characters: 8,
                max_text_bytes: 8,
                timeout_ms: TEST_TIMEOUT_MS,
            },
        )
        .await
        .unwrap();
    assert!(blank_text.text.is_empty());
    assert!(blank_text.characters.is_empty());
    blank_text.validate(&blank_inventory).unwrap();
}

fn inventory(pages: Vec<NativeOfficePdfPageGeometry>) -> NativeOfficePdfPageInventory {
    NativeOfficePdfPageInventory {
        kind: NativeOfficeLayoutSourceKind::Pdf,
        source_revision: PackageRevision {
            archive_bytes: 24,
            sha256: "a".repeat(64),
        },
        max_pages: 10,
        total_pages: pages.len(),
        dpi_milli: DPI_MILLI,
        pages,
    }
}

fn page(
    number: u32,
    rotation_degrees: u16,
    surface_width_millipoints: i64,
    surface_height_millipoints: i64,
) -> NativeOfficePdfPageGeometry {
    let (effective_width, effective_height) = if matches!(rotation_degrees, 90 | 270) {
        (surface_height_millipoints, surface_width_millipoints)
    } else {
        (surface_width_millipoints, surface_height_millipoints)
    };
    let width_px = ((effective_width as u128 * DPI_MILLI as u128 + 36_000_000) / 72_000_000) as u32;
    let height_px =
        ((effective_height as u128 * DPI_MILLI as u128 + 36_000_000) / 72_000_000) as u32;
    NativeOfficePdfPageGeometry {
        unit: NativeOfficeUnit {
            ordinal: number,
            locator: NativeOfficeUnitLocator::Page { number },
            path: format!("/page[{number}]"),
        },
        media_box: NativeOfficePdfPageBox {
            left_millipoints: 0,
            bottom_millipoints: 0,
            right_millipoints: surface_width_millipoints,
            top_millipoints: surface_height_millipoints,
        },
        crop_box: NativeOfficePdfPageBox {
            left_millipoints: 0,
            bottom_millipoints: 0,
            right_millipoints: surface_width_millipoints,
            top_millipoints: surface_height_millipoints,
        },
        rotation_degrees,
        surface_width_micrometers: ((effective_width as u128 * 25_400 + 36_000) / 72_000) as u32,
        surface_height_micrometers: ((effective_height as u128 * 25_400 + 36_000) / 72_000) as u32,
        output_width_px: width_px,
        output_height_px: height_px,
    }
}

fn text_character(
    index: u32,
    text: &str,
    utf8_start: u64,
    utf8_end: u64,
    utf16_start: u64,
    utf16_end: u64,
    bounds: Option<NativeOfficePdfPageBox>,
) -> NativeOfficePdfTextCharacter {
    NativeOfficePdfTextCharacter {
        index,
        text: text.to_string(),
        utf8_start,
        utf8_end,
        utf16_start,
        utf16_end,
        bounds,
        font_size_millipoints: Some(12_000),
        rotation_millidegrees: Some(0),
        generated: Some(false),
    }
}

fn text_box(
    left_millipoints: i64,
    bottom_millipoints: i64,
    right_millipoints: i64,
    top_millipoints: i64,
) -> NativeOfficePdfPageBox {
    NativeOfficePdfPageBox {
        left_millipoints,
        bottom_millipoints,
        right_millipoints,
        top_millipoints,
    }
}

fn two_page_pdf() -> Vec<u8> {
    let red = b"1 0 0 rg\n0 0 200 100 re f\nBT /F1 12 Tf 20 50 Td (Hello PDF) Tj ET\n";
    let blue = b"0 0 1 rg\n0 0 120 240 re f\nBT /F1 12 Tf 20 100 Td (Rotated PDF) Tj ET\n";
    let objects = vec![
        b"<< /Type /Catalog /Pages 2 0 R /Outlines 8 0 R /PageMode /UseOutlines >>".to_vec(),
        b"<< /Type /Pages /Count 2 /Kids [3 0 R 4 0 R] >>".to_vec(),
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 100] /CropBox [10 20 190 80] /Resources << /Font << /F1 7 0 R >> >> /Contents 5 0 R >>".to_vec(),
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 120 240] /Rotate 90 /Resources << /Font << /F1 7 0 R >> >> /Contents 6 0 R >>".to_vec(),
        pdf_stream(red),
        pdf_stream(blue),
        b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>".to_vec(),
        b"<< /Type /Outlines /First 9 0 R /Last 10 0 R /Count 2 >>".to_vec(),
        b"<< /Title (First page) /Parent 8 0 R /Next 10 0 R /Dest [3 0 R /Fit] >>".to_vec(),
        b"<< /Title (Second page) /Parent 8 0 R /Prev 9 0 R /Dest [4 0 R /Fit] >>".to_vec(),
    ];
    encode_pdf(&objects)
}

fn zero_page_pdf() -> Vec<u8> {
    encode_pdf(&[
        b"<< /Type /Catalog /Pages 2 0 R >>".to_vec(),
        b"<< /Type /Pages /Count 0 /Kids [] >>".to_vec(),
    ])
}

fn blank_page_pdf() -> Vec<u8> {
    encode_pdf(&[
        b"<< /Type /Catalog /Pages 2 0 R >>".to_vec(),
        b"<< /Type /Pages /Count 1 /Kids [3 0 R] >>".to_vec(),
        b"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 100] /Resources << >> >>".to_vec(),
    ])
}

fn pdf_stream(content: &[u8]) -> Vec<u8> {
    let mut stream = format!("<< /Length {} >>\nstream\n", content.len()).into_bytes();
    stream.extend_from_slice(content);
    stream.extend_from_slice(b"endstream");
    stream
}

fn encode_pdf(objects: &[Vec<u8>]) -> Vec<u8> {
    let mut pdf = b"%PDF-1.7\n%\xE2\xE3\xCF\xD3\n".to_vec();
    let mut offsets = Vec::with_capacity(objects.len());
    for (offset, object) in objects.iter().enumerate() {
        offsets.push(pdf.len());
        pdf.extend_from_slice(format!("{} 0 obj\n", offset + 1).as_bytes());
        pdf.extend_from_slice(object);
        pdf.extend_from_slice(b"\nendobj\n");
    }
    let xref = pdf.len();
    pdf.extend_from_slice(format!("xref\n0 {}\n", objects.len() + 1).as_bytes());
    pdf.extend_from_slice(b"0000000000 65535 f \n");
    for offset in offsets {
        pdf.extend_from_slice(format!("{offset:010} 00000 n \n").as_bytes());
    }
    pdf.extend_from_slice(
        format!(
            "trailer\n<< /Size {} /Root 1 0 R >>\nstartxref\n{xref}\n%%EOF\n",
            objects.len() + 1
        )
        .as_bytes(),
    );
    pdf
}

fn rgba_pixel(png: &[u8], x: usize, y: usize) -> [u8; 4] {
    let mut reader = png::Decoder::new(png).read_info().unwrap();
    let mut pixels = vec![0; reader.output_buffer_size()];
    let output = reader.next_frame(&mut pixels).unwrap();
    assert_eq!(output.color_type, png::ColorType::Rgba);
    let width = output.width as usize;
    let offset = (y * width + x) * 4;
    pixels[offset..offset + 4].try_into().unwrap()
}

fn sha256_file(path: &std::path::Path) -> String {
    format!("{:x}", sha2::Sha256::digest(std::fs::read(path).unwrap()))
}
