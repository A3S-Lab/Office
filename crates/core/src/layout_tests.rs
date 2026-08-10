use std::io::Write as _;
use std::path::PathBuf;

use sha2::{Digest, Sha256};

use crate::{
    NativeOfficeEditor, NativeOfficeImage, NativeOfficeLayoutAuthority,
    NativeOfficeLayoutEnvironment, NativeOfficeLayoutRenderer, NativeOfficePackage,
    NativeOfficePptxImageLayoutRenderer, NativeOfficeUnit, NativeOfficeUnitLocator,
};

const TEST_TIMEOUT_MS: u64 = 10_000;
const TEST_MAX_OUTPUT_BYTES: u64 = 1024 * 1024;
const CONSTRAINED_ASYNC_STACK_BYTES: usize = 1024 * 1024;
const SLIDE_WIDTH_EMU: &str = "12192000";
const SLIDE_HEIGHT_EMU: &str = "6858000";

#[tokio::test]
async fn exact_pptx_image_slide_is_a_source_layout_raster() {
    let fixture = exact_image_deck().await;
    let renderer = renderer();
    let inspection = renderer
        .inspect_unit(
            &fixture.path,
            fixture.revision.clone(),
            slide_unit(2),
            environment(),
            TEST_TIMEOUT_MS,
        )
        .await
        .unwrap();
    assert_eq!(
        inspection.profile.authority,
        NativeOfficeLayoutAuthority::SourceLayout
    );
    assert_eq!(inspection.profile.output_width_px, 160);
    assert_eq!(inspection.profile.output_height_px, 90);
    assert_eq!(inspection.profile.surface_width_micrometers, 338_667);
    assert_eq!(inspection.profile.surface_height_micrometers, 190_500);

    let output = fixture.directory.path().join("slide-2.png");
    let request = inspection.into_render_request(&output, TEST_MAX_OUTPUT_BYTES, TEST_TIMEOUT_MS);
    let receipt = renderer.render(request.clone()).await.unwrap();

    let published = std::fs::read(&output).unwrap();
    assert_eq!(published, fixture.second_png);
    assert_ne!(published, fixture.first_png);
    assert_eq!(receipt.source_revision, fixture.revision);
    assert_eq!(receipt.unit, slide_unit(2));
    assert_eq!(receipt.raster.output_path, output);
    assert_eq!(receipt.raster.byte_length, published.len() as u64);
    assert_eq!(
        receipt.raster.sha256,
        format!("{:x}", Sha256::digest(&published))
    );
    assert_eq!(receipt.profile_sha256, receipt.profile.sha256().unwrap());

    // Normalized overlay x=0.25 and x=0.75 still address the original red and
    // blue source pixels because the provider performs no resampling.
    let pixels = decode_rgb(&published);
    assert_eq!(pixel(&pixels, 160, 40, 45), [220, 20, 60]);
    assert_eq!(pixel(&pixels, 160, 120, 45), [30, 90, 220]);

    let error = renderer.render(request).await.unwrap_err();
    assert_eq!(error.code, "use.office.layout_output_exists");
    assert_eq!(
        std::fs::read(&receipt.raster.output_path).unwrap(),
        published
    );
}

#[test]
fn exact_pptx_image_render_supports_a_constrained_async_thread_stack() {
    std::thread::Builder::new()
        .name("office-layout-constrained-stack".to_string())
        .stack_size(CONSTRAINED_ASYNC_STACK_BYTES)
        .spawn(|| {
            tokio::runtime::Builder::new_current_thread()
                .enable_all()
                .build()
                .unwrap()
                .block_on(async {
                    let fixture = exact_image_deck().await;
                    let renderer = renderer();
                    let inspection = renderer
                        .inspect_unit(
                            &fixture.path,
                            fixture.revision,
                            slide_unit(1),
                            environment(),
                            TEST_TIMEOUT_MS,
                        )
                        .await
                        .unwrap();
                    let output = fixture.directory.path().join("constrained-stack.png");
                    let request = inspection.into_render_request(
                        &output,
                        TEST_MAX_OUTPUT_BYTES,
                        TEST_TIMEOUT_MS,
                    );

                    renderer.render(request).await.unwrap();

                    assert_eq!(std::fs::read(output).unwrap(), fixture.first_png);
                });
        })
        .unwrap()
        .join()
        .unwrap();
}

#[tokio::test]
async fn layout_selection_rejects_locator_identity_drift() {
    let fixture = exact_image_deck().await;
    let renderer = renderer();
    let error = renderer
        .inspect_unit(
            &fixture.path,
            fixture.revision,
            NativeOfficeUnit {
                ordinal: 1,
                locator: NativeOfficeUnitLocator::Slide { number: 2 },
                path: "/slide[1]".to_string(),
            },
            environment(),
            TEST_TIMEOUT_MS,
        )
        .await
        .unwrap_err();

    assert_eq!(error.code, "use.office.layout_unit_mismatch");
}

#[tokio::test]
async fn layout_render_fails_closed_after_source_mutation() {
    let fixture = exact_image_deck().await;
    let renderer = renderer();
    let inspection = renderer
        .inspect_unit(
            &fixture.path,
            fixture.revision,
            slide_unit(1),
            environment(),
            TEST_TIMEOUT_MS,
        )
        .await
        .unwrap();
    let output = fixture.directory.path().join("mutated.png");
    let request = inspection.into_render_request(&output, TEST_MAX_OUTPUT_BYTES, TEST_TIMEOUT_MS);
    std::fs::OpenOptions::new()
        .append(true)
        .open(&fixture.path)
        .unwrap()
        .write_all(b"mutated")
        .unwrap();

    let error = renderer.render(request).await.unwrap_err();

    assert_eq!(error.code, "use.office.layout_source_mutated");
    assert!(!output.exists());
}

#[tokio::test]
async fn layout_render_fails_closed_when_source_changes_during_render() {
    let fixture = exact_image_deck().await;
    let renderer = renderer();
    let inspection = renderer
        .inspect_unit(
            &fixture.path,
            fixture.revision,
            slide_unit(1),
            environment(),
            TEST_TIMEOUT_MS,
        )
        .await
        .unwrap();
    let output = fixture.directory.path().join("mutated-during-render.png");
    let request = inspection.into_render_request(&output, TEST_MAX_OUTPUT_BYTES, TEST_TIMEOUT_MS);
    let source_path = fixture.path.clone();

    let error = renderer
        .render_with_before_final_source_check(request, move || {
            std::fs::OpenOptions::new()
                .append(true)
                .open(source_path)
                .unwrap()
                .write_all(b"mutated-during-render")
                .unwrap();
        })
        .await
        .unwrap_err();

    assert_eq!(error.code, "use.office.layout_source_mutated");
    assert!(!output.exists());
}

#[tokio::test]
async fn semantic_authority_and_non_exact_slides_cannot_be_promoted() {
    let fixture = exact_image_deck().await;
    let renderer = renderer();
    let inspection = renderer
        .inspect_unit(
            &fixture.path,
            fixture.revision,
            slide_unit(1),
            environment(),
            TEST_TIMEOUT_MS,
        )
        .await
        .unwrap();
    let output = fixture.directory.path().join("semantic.png");
    let mut request =
        inspection.into_render_request(&output, TEST_MAX_OUTPUT_BYTES, TEST_TIMEOUT_MS);
    request.profile.authority = NativeOfficeLayoutAuthority::SemanticPreview;
    let error = renderer.render(request).await.unwrap_err();
    assert_eq!(error.code, "use.office.layout_profile_invalid");
    assert!(!output.exists());

    let (directory, path, revision) = non_exact_deck().await;
    let error = renderer
        .inspect_unit(
            &path,
            revision,
            slide_unit(1),
            environment(),
            TEST_TIMEOUT_MS,
        )
        .await
        .unwrap_err();
    assert_eq!(error.code, "use.office.layout_unsupported");
    drop(directory);
}

#[tokio::test]
async fn hidden_picture_cannot_be_claimed_as_the_visible_slide_surface() {
    let fixture = exact_image_deck().await;
    let revision =
        replace_slide_xml(&fixture.path, 1, "descr=\"\"", "descr=\"\" hidden=\"1\"").await;

    let error = renderer()
        .inspect_unit(
            &fixture.path,
            revision,
            slide_unit(1),
            environment(),
            TEST_TIMEOUT_MS,
        )
        .await
        .unwrap_err();

    assert_eq!(error.code, "use.office.layout_unsupported");
}

#[tokio::test]
async fn layout_profile_has_a_strict_deterministic_json_contract() {
    let fixture = exact_image_deck().await;
    let inspection = renderer()
        .inspect_unit(
            &fixture.path,
            fixture.revision,
            slide_unit(1),
            environment(),
            TEST_TIMEOUT_MS,
        )
        .await
        .unwrap();
    let profile = inspection.profile;
    let original_hash = profile.sha256().unwrap();
    let mut value = serde_json::to_value(&profile).unwrap();
    assert_eq!(value["authority"], "source-layout");
    assert_eq!(value["schemaVersion"], 1);
    assert_eq!(
        serde_json::from_value::<crate::NativeOfficeLayoutProfile>(value.clone()).unwrap(),
        profile
    );

    let mut different_timezone = profile;
    different_timezone.timezone = "Asia/Shanghai".to_string();
    assert_ne!(original_hash, different_timezone.sha256().unwrap());

    value
        .as_object_mut()
        .unwrap()
        .insert("unreviewedField".to_string(), serde_json::Value::Bool(true));
    assert!(serde_json::from_value::<crate::NativeOfficeLayoutProfile>(value).is_err());
}

#[tokio::test]
async fn layout_profile_accepts_independent_surface_and_pixel_quantization() {
    let fixture = exact_image_deck().await;
    let mut profile = renderer()
        .inspect_unit(
            &fixture.path,
            fixture.revision,
            slide_unit(1),
            environment(),
            TEST_TIMEOUT_MS,
        )
        .await
        .unwrap()
        .profile;
    profile.dpi_x_milli = 144_000;
    profile.dpi_y_milli = 144_000;
    profile.surface_width_micrometers = 210_009;
    profile.surface_height_micrometers = 297_004;
    profile.output_width_px = 1_191;
    profile.output_height_px = 1_684;

    profile.validate().unwrap();

    profile.output_width_px += 1;
    assert_eq!(
        profile.validate().unwrap_err().code,
        "use.office.layout_profile_invalid"
    );
}

#[tokio::test]
async fn layout_output_budget_is_checked_before_publication() {
    let fixture = exact_image_deck().await;
    let renderer = renderer();
    let inspection = renderer
        .inspect_unit(
            &fixture.path,
            fixture.revision,
            slide_unit(1),
            environment(),
            TEST_TIMEOUT_MS,
        )
        .await
        .unwrap();
    let output = fixture.directory.path().join("too-small.png");
    let request = inspection.into_render_request(
        &output,
        u64::try_from(fixture.first_png.len()).unwrap() - 1,
        TEST_TIMEOUT_MS,
    );

    let error = renderer.render(request).await.unwrap_err();

    assert_eq!(error.code, "use.office.layout_output_too_large");
    assert!(!output.exists());
}

#[test]
fn layout_renderer_contract_is_object_safe_send_and_sync() {
    fn assert_send_sync<T: Send + Sync + ?Sized>() {}
    assert_send_sync::<dyn NativeOfficeLayoutRenderer>();
}

struct ExactDeckFixture {
    directory: tempfile::TempDir,
    path: PathBuf,
    revision: crate::PackageRevision,
    first_png: Vec<u8>,
    second_png: Vec<u8>,
}

async fn exact_image_deck() -> ExactDeckFixture {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("exact-images.pptx");
    let first_png = split_png([220, 20, 60], [245, 160, 20]);
    let second_png = split_png([220, 20, 60], [30, 90, 220]);
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    for png in [&first_png, &second_png] {
        let slide = editor.add_slide("/", "").unwrap();
        editor
            .add_image(
                &slide,
                NativeOfficeImage::from_bytes(png)
                    .unwrap()
                    .with_width_px(1_280)
                    .with_height_px(720),
            )
            .unwrap();
    }
    let mut package = editor.package().clone();
    for number in 1..=2 {
        make_picture_fill_slide(&mut package, number);
    }
    let mut editor = NativeOfficeEditor::from_package(package).unwrap();
    editor.save().await.unwrap();
    let package = NativeOfficePackage::open(&path).await.unwrap();
    let revision = package.source_revision().clone();
    ExactDeckFixture {
        directory,
        path,
        revision,
        first_png,
        second_png,
    }
}

async fn non_exact_deck() -> (tempfile::TempDir, PathBuf, crate::PackageRevision) {
    let directory = tempfile::tempdir().unwrap();
    let path = directory.path().join("semantic.pptx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor.add_slide("/", "Not an image-only slide").unwrap();
    editor.save().await.unwrap();
    let package = NativeOfficePackage::open(&path).await.unwrap();
    let revision = package.source_revision().clone();
    (directory, path, revision)
}

fn make_picture_fill_slide(package: &mut NativeOfficePackage, number: u32) {
    let part = format!("ppt/slides/slide{number}.xml");
    let original = std::str::from_utf8(package.part(&part).unwrap()).unwrap();
    assert!(original.contains(&format!(
        "<a:ext cx=\"{SLIDE_WIDTH_EMU}\" cy=\"{SLIDE_HEIGHT_EMU}\"/>"
    )));
    let updated = original.replace(
        "<a:off x=\"914400\" y=\"914400\"/>",
        "<a:off x=\"0\" y=\"0\"/>",
    );
    assert_ne!(updated, original);
    package.set_part(&part, updated.into_bytes()).unwrap();
}

async fn replace_slide_xml(
    path: &PathBuf,
    number: u32,
    from: &str,
    to: &str,
) -> crate::PackageRevision {
    let mut package = NativeOfficePackage::open(path).await.unwrap();
    let part = format!("ppt/slides/slide{number}.xml");
    let original = std::str::from_utf8(package.part(&part).unwrap()).unwrap();
    let updated = original.replacen(from, to, 1);
    assert_ne!(updated, original);
    package.set_part(&part, updated.into_bytes()).unwrap();
    let mut editor = NativeOfficeEditor::from_package(package).unwrap();
    editor.save().await.unwrap();
    NativeOfficePackage::open(path)
        .await
        .unwrap()
        .source_revision()
        .clone()
}

fn renderer() -> NativeOfficePptxImageLayoutRenderer {
    NativeOfficePptxImageLayoutRenderer::new("a".repeat(64)).unwrap()
}

fn environment() -> NativeOfficeLayoutEnvironment {
    NativeOfficeLayoutEnvironment::new("en-US", "UTC")
}

fn slide_unit(number: u32) -> NativeOfficeUnit {
    NativeOfficeUnit {
        ordinal: number,
        locator: NativeOfficeUnitLocator::Slide { number },
        path: format!("/slide[{number}]"),
    }
}

fn split_png(left: [u8; 3], right: [u8; 3]) -> Vec<u8> {
    let mut pixels = Vec::with_capacity(160 * 90 * 3);
    for _y in 0..90 {
        for x in 0..160 {
            pixels.extend_from_slice(if x < 80 { &left } else { &right });
        }
    }
    let mut bytes = Vec::new();
    {
        let mut encoder = png::Encoder::new(&mut bytes, 160, 90);
        encoder.set_color(png::ColorType::Rgb);
        encoder.set_depth(png::BitDepth::Eight);
        let mut writer = encoder.write_header().unwrap();
        writer.write_image_data(&pixels).unwrap();
    }
    bytes
}

fn decode_rgb(bytes: &[u8]) -> Vec<u8> {
    let decoder = png::Decoder::new(bytes);
    let mut reader = decoder.read_info().unwrap();
    let mut pixels = vec![0; reader.output_buffer_size()];
    let output = reader.next_frame(&mut pixels).unwrap();
    assert_eq!(output.color_type, png::ColorType::Rgb);
    assert_eq!((output.width, output.height), (160, 90));
    pixels.truncate(output.buffer_size());
    pixels
}

fn pixel(pixels: &[u8], width: usize, x: usize, y: usize) -> [u8; 3] {
    let offset = (y * width + x) * 3;
    pixels[offset..offset + 3].try_into().unwrap()
}
