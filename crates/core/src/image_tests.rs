use crate::{
    DocumentKind, NativeOfficeEditor, NativeOfficeImage, NativeOfficeImageFormat,
    NativeOfficeMutation, OfficeNodeType,
};

pub(crate) const PNG_1X1: &[u8] = &[
    0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52,
    0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x06, 0x00, 0x00, 0x00, 0x1f, 0x15, 0xc4,
    0x89, 0x00, 0x00, 0x00, 0x0a, 0x49, 0x44, 0x41, 0x54, 0x78, 0x9c, 0x63, 0x00, 0x01, 0x00, 0x00,
    0x05, 0x00, 0x01, 0x0d, 0x0a, 0x2d, 0xb4, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae,
    0x42, 0x60, 0x82,
];

const GIF_1X1: &[u8] = &[
    0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x01, 0x00, 0x01, 0x00, 0x80, 0x00, 0x00, 0x00, 0x00, 0x00,
    0xff, 0xff, 0xff, 0x2c, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00, 0x02, 0x02, 0x44,
    0x01, 0x00, 0x3b,
];

const JPEG_2X2_HEX: &str = concat!(
    "ffd8ffe000104a46494600010200000100010000fffe00104c61766336322e32382e31303000",
    "ffdb004300080404040404050505050505060606060606060606060606060707070808080707",
    "0706060707080808080909090808080809090a0a0a0c0c0b0b0e0e0e111114ffc4004c0001",
    "0100000000000000000000000000000006010101000000000000000000000000000006071001",
    "00000000000000000000000000000000110100000000000000000000000000000000ffc00011",
    "080002000203012200021100031100ffda000c03010002110311003f008b004d7f7fffd9"
);

fn decode_hex(value: &str) -> Vec<u8> {
    value
        .as_bytes()
        .chunks_exact(2)
        .map(|pair| {
            let pair = std::str::from_utf8(pair).unwrap();
            u8::from_str_radix(pair, 16).unwrap()
        })
        .collect()
}

#[test]
fn native_images_detect_png_jpeg_and_gif_and_infer_dimensions() {
    let jpeg = decode_hex(JPEG_2X2_HEX);
    for (bytes, format) in [
        (PNG_1X1, NativeOfficeImageFormat::Png),
        (jpeg.as_slice(), NativeOfficeImageFormat::Jpeg),
        (GIF_1X1, NativeOfficeImageFormat::Gif),
    ] {
        assert_eq!(
            NativeOfficeImage::from_bytes(bytes).unwrap().format(),
            format
        );
    }

    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("dimensions.docx");
    let runtime = tokio::runtime::Runtime::new().unwrap();
    let mut editor = runtime.block_on(NativeOfficeEditor::create(&path)).unwrap();
    let created = editor
        .add_image(
            "/body",
            NativeOfficeImage::from_bytes(&jpeg)
                .unwrap()
                .with_width_px(100),
        )
        .unwrap();
    assert_eq!(created.width_px, 100);
    assert_eq!(created.height_px, 100);
}

#[tokio::test]
async fn native_editor_adds_reads_and_removes_embedded_images_in_all_formats() {
    let temp = tempfile::tempdir().unwrap();

    for (kind, extension, parent) in [
        (DocumentKind::Word, "docx", "/body"),
        (DocumentKind::Spreadsheet, "xlsx", "/Sheet1/A1"),
        (DocumentKind::Presentation, "pptx", "/slide[1]"),
    ] {
        let path = temp.path().join(format!("image.{extension}"));
        let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
        if kind == DocumentKind::Presentation {
            editor.add_slide("/", "Images").unwrap();
        }
        let image = NativeOfficeImage::from_bytes(PNG_1X1)
            .unwrap()
            .with_name("A3S Logo")
            .with_alt_text("A3S test logo")
            .with_width_px(96)
            .with_height_px(48);
        let created = editor.add_image(parent, image).unwrap();

        assert_eq!(created.format, NativeOfficeImageFormat::Png);
        assert_eq!(created.width_px, 96);
        assert_eq!(created.height_px, 48);
        assert_eq!(
            editor
                .package()
                .part(created.part.trim_start_matches('/'))
                .unwrap(),
            PNG_1X1
        );
        assert_eq!(editor.snapshot().unwrap().statistics().picture_count, 1);
        let node = editor.snapshot().unwrap().get(&created.path, 0).unwrap();
        assert_eq!(node.node_type, OfficeNodeType::Picture);
        assert_eq!(node.format["relationshipId"], created.relationship_id);
        assert_eq!(node.format["name"], "A3S Logo");
        assert_eq!(node.format["alt"], "A3S test logo");
        assert_eq!(node.format["widthPx"], "96");
        assert_eq!(node.format["heightPx"], "48");

        editor.remove(&created.path).unwrap();
        assert!(!editor
            .package()
            .contains_part(created.part.trim_start_matches('/')));
        assert_eq!(
            editor
                .snapshot()
                .unwrap()
                .get(&created.path, 0)
                .unwrap_err()
                .code,
            "use.office.node_not_found"
        );
        assert_eq!(editor.snapshot().unwrap().statistics().picture_count, 0);
    }
}

#[tokio::test]
async fn invalid_image_data_rolls_back_the_whole_native_batch() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("rollback.docx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    let before = editor.package().content_sha256();
    let invalid: NativeOfficeImage = serde_json::from_value(serde_json::json!({
        "format": "png",
        "data": "bm90IGFuIGltYWdl",
        "name": "Invalid"
    }))
    .unwrap();

    let error = editor
        .apply_batch(&[
            NativeOfficeMutation::AddParagraph {
                parent: "/body".into(),
                text: "must roll back".into(),
            },
            NativeOfficeMutation::AddImage {
                parent: "/body".into(),
                image: invalid,
            },
        ])
        .unwrap_err();
    assert_eq!(error.code, "use.office.image_invalid");
    assert_eq!(editor.package().content_sha256(), before);

    let error = editor
        .apply_batch(&[
            NativeOfficeMutation::AddImage {
                parent: "/body".into(),
                image: NativeOfficeImage::from_bytes(PNG_1X1).unwrap(),
            },
            NativeOfficeMutation::AddSlide {
                parent: "/".into(),
                title: "wrong document kind".into(),
            },
        ])
        .unwrap_err();
    assert_eq!(error.code, "use.office.mutation_type_unsupported");
    assert_eq!(editor.package().content_sha256(), before);
}

#[tokio::test]
async fn image_removal_preserves_media_targeted_by_another_relationship() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("shared.docx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    let created = editor
        .add_image("/body", NativeOfficeImage::from_bytes(PNG_1X1).unwrap())
        .unwrap();
    let media = created.part.trim_start_matches('/');
    let mut package = editor.package().clone();
    let shared_id = crate::opc_edit::add_relationship(
        &mut package,
        "word/_rels/document.xml.rels",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
        &format!("media/{}", media.rsplit('/').next().unwrap()),
    )
    .unwrap();
    let mut editor = NativeOfficeEditor::from_package(package).unwrap();

    editor.remove(&created.path).unwrap();
    assert!(editor.package().contains_part(media));
    let source = crate::RelationshipSource::Part {
        part_name: "word/document.xml".into(),
    };
    let model = editor.package().opc_model().unwrap();
    assert!(model
        .relationships()
        .relationship(&source, &created.relationship_id)
        .is_none());
    assert!(model
        .relationships()
        .relationship(&source, &shared_id)
        .is_some());
}

#[test]
fn image_mutations_and_receipts_have_a_stable_batch_schema() {
    let image = NativeOfficeImage::from_bytes(PNG_1X1)
        .unwrap()
        .with_name("Logo")
        .with_width_px(320);
    let mutation = NativeOfficeMutation::AddImage {
        parent: "/slide[1]".into(),
        image: image.clone(),
    };
    let value = serde_json::to_value(&mutation).unwrap();
    assert_eq!(value["operation"], "add-image");
    assert_eq!(value["parent"], "/slide[1]");
    assert_eq!(value["image"]["format"], "png");
    assert_eq!(value["image"]["name"], "Logo");
    assert_eq!(value["image"]["widthPx"], 320);
    assert!(value["image"]["data"].as_str().unwrap().len() > PNG_1X1.len());
    assert_eq!(
        serde_json::from_value::<NativeOfficeMutation>(value).unwrap(),
        mutation
    );

    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<NativeOfficeImage>();
    assert_send_sync::<crate::NativeCreatedImage>();
}
