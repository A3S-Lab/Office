use crate::{
    DocumentKind, NativeOfficeEditor, NativeOfficeHyperlink, NativeOfficeHyperlinkTarget,
    NativeOfficeMutation, NativeOfficePartType, OfficeNodeType, RelationshipSource,
    RelationshipTarget,
};

fn external(uri: &str) -> NativeOfficeHyperlink {
    NativeOfficeHyperlink::external(uri).unwrap()
}

fn relationship_count(editor: &NativeOfficeEditor, owner: &str, suffix: &str) -> usize {
    let source = RelationshipSource::Part {
        part_name: owner.to_string(),
    };
    editor
        .package()
        .opc_model()
        .unwrap()
        .relationships()
        .relationships_from(&source)
        .iter()
        .filter(|relationship| relationship.relationship_type.ends_with(suffix))
        .count()
}

#[test]
fn hyperlink_mutation_has_a_typed_stable_json_contract() {
    let mutation = NativeOfficeMutation::SetHyperlink {
        path: "/body/p[1]".into(),
        hyperlink: external("https://example.com/report")
            .with_display("Open report")
            .with_tooltip("A3S report"),
    };
    assert_eq!(
        serde_json::to_value(&mutation).unwrap(),
        serde_json::json!({
            "operation": "set-hyperlink",
            "path": "/body/p[1]",
            "hyperlink": {
                "target": {
                    "kind": "external",
                    "uri": "https://example.com/report"
                },
                "display": "Open report",
                "tooltip": "A3S report"
            }
        })
    );
    assert_eq!(
        serde_json::from_value::<NativeOfficeMutation>(serde_json::to_value(&mutation).unwrap())
            .unwrap(),
        mutation
    );
    assert!(
        serde_json::from_value::<NativeOfficeMutation>(serde_json::json!({
            "operation": "set-hyperlink",
            "path": "/body/p[1]",
            "hyperlink": {
                "target": { "kind": "external", "uri": "https://example.com" },
                "unknown": true
            }
        }))
        .is_err()
    );
}

#[tokio::test]
async fn native_word_adds_updates_reads_queries_and_removes_hyperlinks() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("links.docx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor.set_text("/body/p[1]", "Before ").unwrap();

    let link_path = editor
        .set_hyperlink(
            "/body/p[1]",
            external("https://example.com/report")
                .with_display("Open report")
                .with_tooltip("A3S report"),
        )
        .unwrap();
    assert_eq!(link_path, "/body/p[1]/hyperlink[1]");

    let link = editor.snapshot().unwrap().get(&link_path, 1).unwrap();
    assert_eq!(link.node_type, OfficeNodeType::Hyperlink);
    assert_eq!(link.text, "Open report");
    assert_eq!(link.format["targetKind"], "external");
    assert_eq!(link.format["target"], "https://example.com/report");
    assert_eq!(link.format["tooltip"], "A3S report");
    assert_eq!(
        editor.snapshot().unwrap().query("hyperlink").unwrap().len(),
        1
    );
    assert_eq!(
        relationship_count(&editor, "word/document.xml", "/hyperlink"),
        1
    );

    let internal = NativeOfficeHyperlink::internal("section_1")
        .unwrap()
        .with_display("Jump inside");
    assert_eq!(
        editor.set_hyperlink(&link_path, internal).unwrap(),
        link_path
    );
    let link = editor.snapshot().unwrap().get(&link_path, 1).unwrap();
    assert_eq!(link.text, "Jump inside");
    assert_eq!(link.format["targetKind"], "internal");
    assert_eq!(link.format["target"], "section_1");
    assert!(!link.format.contains_key("relationshipId"));
    assert_eq!(
        relationship_count(&editor, "word/document.xml", "/hyperlink"),
        0
    );

    editor.remove(&link_path).unwrap();
    assert_eq!(
        editor.snapshot().unwrap().query("hyperlink").unwrap().len(),
        0
    );
    assert_eq!(
        editor
            .snapshot()
            .unwrap()
            .get("/body/p[1]", 0)
            .unwrap()
            .text,
        "Before "
    );
}

#[tokio::test]
async fn native_word_manages_header_and_footer_hyperlinks_in_their_own_parts() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("header-footer-links.docx");
    let mut seed = NativeOfficeEditor::create(&path).await.unwrap();
    let header = seed.add_part("/", NativeOfficePartType::Header).unwrap();
    let footer = seed.add_part("/", NativeOfficePartType::Footer).unwrap();
    let mut package = seed.package().clone();
    package
        .set_part(
            header.part.trim_start_matches('/'),
            br#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:hdr>"#.to_vec(),
        )
        .unwrap();
    package
        .set_part(
            footer.part.trim_start_matches('/'),
            br#"<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Footer</w:t></w:r></w:p></w:ftr>"#.to_vec(),
        )
        .unwrap();
    let mut editor = NativeOfficeEditor::from_package(package).unwrap();

    let header_link = editor
        .set_hyperlink(
            "/header[1]/p[1]",
            external("https://example.com/header")
                .with_display("Header link")
                .with_tooltip("Header tooltip"),
        )
        .unwrap();
    assert_eq!(header_link, "/header[1]/p[1]/hyperlink[1]");
    let node = editor.snapshot().unwrap().get(&header_link, 0).unwrap();
    assert_eq!(node.text, "Header link");
    assert_eq!(node.format["target"], "https://example.com/header");
    assert_eq!(
        relationship_count(&editor, "word/header1.xml", "/hyperlink"),
        1
    );
    assert_eq!(
        relationship_count(&editor, "word/document.xml", "/hyperlink"),
        0
    );

    editor
        .set_hyperlink(
            &header_link,
            NativeOfficeHyperlink::internal("header_bookmark")
                .unwrap()
                .with_display("Inside header"),
        )
        .unwrap();
    let node = editor.snapshot().unwrap().get(&header_link, 0).unwrap();
    assert_eq!(node.format["targetKind"], "internal");
    assert_eq!(node.format["target"], "header_bookmark");
    assert_eq!(
        relationship_count(&editor, "word/header1.xml", "/hyperlink"),
        0
    );

    let footer_link = editor
        .set_hyperlink(
            "/footer[1]/p[1]",
            external("mailto:office@example.com").with_display("Mail footer"),
        )
        .unwrap();
    assert_eq!(footer_link, "/footer[1]/p[1]/hyperlink[1]");
    assert_eq!(
        editor
            .snapshot()
            .unwrap()
            .get(&footer_link, 0)
            .unwrap()
            .format["target"],
        "mailto:office@example.com"
    );

    editor.remove(&header_link).unwrap();
    editor.remove(&footer_link).unwrap();
    assert!(editor
        .snapshot()
        .unwrap()
        .query("hyperlink")
        .unwrap()
        .is_empty());
    assert_eq!(
        relationship_count(&editor, "word/footer1.xml", "/hyperlink"),
        0
    );
}

#[tokio::test]
async fn native_spreadsheet_sets_internal_and_external_cell_hyperlinks() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("links.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();

    let link_path = editor
        .set_hyperlink(
            "/Sheet1/A1",
            external("https://example.com/data")
                .with_display("Data")
                .with_tooltip("Open data"),
        )
        .unwrap();
    assert_eq!(link_path, "/Sheet1/A1/hyperlink");
    assert_eq!(
        editor
            .snapshot()
            .unwrap()
            .get("/Sheet1/A1", 1)
            .unwrap()
            .text,
        "Data"
    );
    let link = editor.snapshot().unwrap().get(&link_path, 0).unwrap();
    assert_eq!(link.format["targetKind"], "external");
    assert_eq!(link.format["target"], "https://example.com/data");
    assert_eq!(link.format["display"], "Data");
    assert_eq!(link.format["tooltip"], "Open data");
    assert_eq!(
        relationship_count(&editor, "xl/worksheets/sheet1.xml", "/hyperlink"),
        1
    );

    let internal = NativeOfficeHyperlink::internal("Sheet1!B2")
        .unwrap()
        .with_display("B2");
    editor.set_hyperlink("/Sheet1/A1", internal).unwrap();
    let link = editor.snapshot().unwrap().get(&link_path, 0).unwrap();
    assert_eq!(link.format["targetKind"], "internal");
    assert_eq!(link.format["target"], "Sheet1!B2");
    assert_eq!(
        relationship_count(&editor, "xl/worksheets/sheet1.xml", "/hyperlink"),
        0
    );

    editor.remove(&link_path).unwrap();
    assert_eq!(
        editor.snapshot().unwrap().query("hyperlink").unwrap().len(),
        0
    );
    assert_eq!(
        editor
            .snapshot()
            .unwrap()
            .get("/Sheet1/A1", 0)
            .unwrap()
            .text,
        "Data"
    );
}

#[tokio::test]
async fn native_spreadsheet_manages_bounded_range_hyperlinks_without_overlaps() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("range-links.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor.set_text("/Sheet1/A1", "Keep A1").unwrap();
    editor.set_text("/Sheet1/B2", "Keep B2").unwrap();

    let link_path = editor
        .set_hyperlink(
            "/Sheet1/A1:B2",
            external("https://example.com/range")
                .with_display("Range")
                .with_tooltip("Open range"),
        )
        .unwrap();
    assert_eq!(link_path, "/Sheet1/hyperlink[1]");
    let link = editor.snapshot().unwrap().get(&link_path, 0).unwrap();
    assert_eq!(link.format["ref"], "A1:B2");
    assert_eq!(link.format["targetKind"], "external");
    assert_eq!(link.format["target"], "https://example.com/range");
    assert_eq!(
        editor
            .snapshot()
            .unwrap()
            .get("/Sheet1/A1", 0)
            .unwrap()
            .text,
        "Keep A1"
    );
    assert_eq!(
        relationship_count(&editor, "xl/worksheets/sheet1.xml", "/hyperlink"),
        1
    );

    editor
        .set_hyperlink(
            &link_path,
            NativeOfficeHyperlink::internal("Sheet1!D4")
                .unwrap()
                .with_display("Internal range"),
        )
        .unwrap();
    let link = editor.snapshot().unwrap().get(&link_path, 0).unwrap();
    assert_eq!(link.format["targetKind"], "internal");
    assert_eq!(link.format["target"], "Sheet1!D4");
    assert_eq!(
        relationship_count(&editor, "xl/worksheets/sheet1.xml", "/hyperlink"),
        0
    );

    let before = editor.package().content_sha256();
    let conflict = editor
        .set_hyperlink("/Sheet1/B2:C3", external("https://example.com/overlap"))
        .unwrap_err();
    assert_eq!(conflict.code, "use.office.hyperlink_range_conflict");
    assert_eq!(editor.package().content_sha256(), before);

    editor.remove(&link_path).unwrap();
    assert!(editor
        .snapshot()
        .unwrap()
        .query("hyperlink")
        .unwrap()
        .is_empty());
}

#[tokio::test]
async fn native_presentation_switches_between_external_links_and_internal_slide_jumps() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("links.pptx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor.add_slide("/", "Linked shape").unwrap();
    editor.add_slide("/", "Target slide").unwrap();

    let link_path = editor
        .set_hyperlink(
            "/slide[1]/shape[1]",
            external("https://example.com/slides").with_tooltip("Open slides"),
        )
        .unwrap();
    assert_eq!(link_path, "/slide[1]/shape[1]/hyperlink");
    let link = editor.snapshot().unwrap().get(&link_path, 0).unwrap();
    assert_eq!(link.format["targetKind"], "external");
    assert_eq!(link.format["target"], "https://example.com/slides");
    assert_eq!(link.format["tooltip"], "Open slides");
    assert_eq!(
        relationship_count(&editor, "ppt/slides/slide1.xml", "/hyperlink"),
        1
    );

    editor
        .set_hyperlink(
            &link_path,
            external("https://example.com/slides").with_tooltip("Updated"),
        )
        .unwrap();
    assert_eq!(
        relationship_count(&editor, "ppt/slides/slide1.xml", "/hyperlink"),
        1
    );
    assert_eq!(
        editor
            .snapshot()
            .unwrap()
            .get(&link_path, 0)
            .unwrap()
            .format["tooltip"],
        "Updated"
    );

    editor
        .set_hyperlink(
            &link_path,
            NativeOfficeHyperlink::internal("slide[2]")
                .unwrap()
                .with_tooltip("Jump to target"),
        )
        .unwrap();
    let link = editor.snapshot().unwrap().get(&link_path, 0).unwrap();
    assert_eq!(link.format["targetKind"], "internal");
    assert_eq!(link.format["target"], "/slide[2]");
    assert_eq!(link.format["action"], "ppaction://hlinksldjump");
    assert_eq!(link.format["tooltip"], "Jump to target");
    assert_eq!(
        relationship_count(&editor, "ppt/slides/slide1.xml", "/hyperlink"),
        0
    );
    assert_eq!(
        relationship_count(&editor, "ppt/slides/slide1.xml", "/slide"),
        1
    );

    editor
        .set_hyperlink(
            &link_path,
            NativeOfficeHyperlink::internal("/slide[2]")
                .unwrap()
                .with_tooltip("Updated jump"),
        )
        .unwrap();
    assert_eq!(
        relationship_count(&editor, "ppt/slides/slide1.xml", "/slide"),
        1
    );

    let before = editor.package().content_sha256();
    let error = editor
        .set_hyperlink(
            &link_path,
            NativeOfficeHyperlink::internal("target-slide").unwrap(),
        )
        .unwrap_err();
    assert_eq!(error.code, "use.office.hyperlink_location_invalid");
    assert_eq!(editor.package().content_sha256(), before);

    editor.remove(&link_path).unwrap();
    assert_eq!(
        editor.snapshot().unwrap().query("hyperlink").unwrap().len(),
        0
    );
    assert_eq!(
        relationship_count(&editor, "ppt/slides/slide1.xml", "/hyperlink"),
        0
    );
    assert_eq!(
        relationship_count(&editor, "ppt/slides/slide1.xml", "/slide"),
        0
    );
}

#[tokio::test]
async fn hyperlink_validation_and_owner_removal_are_atomic_and_relationship_safe() {
    let temp = tempfile::tempdir().unwrap();
    for (kind, extension, owner, owner_part) in [
        (
            DocumentKind::Word,
            "docx",
            "/body/p[1]",
            "word/document.xml",
        ),
        (
            DocumentKind::Spreadsheet,
            "xlsx",
            "/Sheet1/A1",
            "xl/worksheets/sheet1.xml",
        ),
        (
            DocumentKind::Presentation,
            "pptx",
            "/slide[1]/shape[1]",
            "ppt/slides/slide1.xml",
        ),
    ] {
        let path = temp.path().join(format!("owner.{extension}"));
        let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
        match kind {
            DocumentKind::Word => editor.set_text(owner, "Owner").unwrap(),
            DocumentKind::Spreadsheet => editor.set_text(owner, "Owner").unwrap(),
            DocumentKind::Presentation => {
                editor.add_slide("/", "Owner").unwrap();
            }
        }
        editor
            .set_hyperlink(owner, external("https://example.com/owner"))
            .unwrap();
        assert_eq!(relationship_count(&editor, owner_part, "/hyperlink"), 1);

        editor.remove(owner).unwrap();
        assert_eq!(relationship_count(&editor, owner_part, "/hyperlink"), 0);
    }

    let path = temp.path().join("invalid.docx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    let before = editor.package().content_sha256();
    let invalid = NativeOfficeHyperlink {
        target: NativeOfficeHyperlinkTarget::External {
            uri: "javascript:alert(1)".into(),
        },
        display: None,
        tooltip: None,
    };
    let error = editor
        .apply_batch(&[
            NativeOfficeMutation::SetText {
                path: "/body/p[1]".into(),
                text: "must roll back".into(),
            },
            NativeOfficeMutation::SetHyperlink {
                path: "/body/p[1]".into(),
                hyperlink: invalid,
            },
        ])
        .unwrap_err();
    assert_eq!(error.code, "use.office.hyperlink_uri_invalid");
    assert_eq!(editor.package().content_sha256(), before);
}

#[test]
fn hyperlink_relationship_targets_remain_inert_typed_data() {
    let target = RelationshipTarget::External {
        uri: "https://example.com".into(),
    };
    assert_eq!(
        target,
        RelationshipTarget::External {
            uri: "https://example.com".into()
        }
    );

    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<NativeOfficeHyperlink>();
    assert_send_sync::<NativeOfficeHyperlinkTarget>();
}

#[test]
fn hyperlink_targets_reject_active_relative_and_credentialed_uris() {
    for uri in [
        "https://example.com/report",
        "http://example.com/report",
        "mailto:office@example.com",
    ] {
        NativeOfficeHyperlink::external(uri).unwrap();
    }

    for uri in [
        "javascript:alert(1)",
        "file:///tmp/report",
        "../relative",
        "//example.com/relative",
        "https://user:password@example.com/report",
        "https://example.com/report\nInjected: value",
    ] {
        let error = NativeOfficeHyperlink::external(uri).unwrap_err();
        assert_eq!(error.code, "use.office.hyperlink_uri_invalid", "{uri}");
    }
}

#[tokio::test]
async fn native_hyperlinks_preserve_strict_ooxml_dialects() {
    const TRANSITIONAL_RELATIONSHIPS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const STRICT_RELATIONSHIPS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";
    const NAMESPACE_PAIRS: &[(&str, &str)] = &[
        (
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
            "http://purl.oclc.org/ooxml/wordprocessingml/main",
        ),
        (
            "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
            "http://purl.oclc.org/ooxml/spreadsheetml/main",
        ),
        (
            "http://schemas.openxmlformats.org/presentationml/2006/main",
            "http://purl.oclc.org/ooxml/presentationml/main",
        ),
        (
            "http://schemas.openxmlformats.org/drawingml/2006/main",
            "http://purl.oclc.org/ooxml/drawingml/main",
        ),
        (TRANSITIONAL_RELATIONSHIPS, STRICT_RELATIONSHIPS),
    ];

    let temp = tempfile::tempdir().unwrap();
    for (extension, owner, owner_part, relationship_part) in [
        (
            "docx",
            "/body/p[1]",
            "word/document.xml",
            "word/_rels/document.xml.rels",
        ),
        (
            "xlsx",
            "/Sheet1/A1",
            "xl/worksheets/sheet1.xml",
            "xl/worksheets/_rels/sheet1.xml.rels",
        ),
        (
            "pptx",
            "/slide[1]/shape[1]",
            "ppt/slides/slide1.xml",
            "ppt/slides/_rels/slide1.xml.rels",
        ),
    ] {
        let path = temp.path().join(format!("strict.{extension}"));
        let mut seed = NativeOfficeEditor::create(&path).await.unwrap();
        if extension == "pptx" {
            seed.add_slide("/", "Strict").unwrap();
            seed.add_slide("/", "Target").unwrap();
        }
        let mut package = seed.package().clone();
        let parts = package.part_names().map(str::to_string).collect::<Vec<_>>();
        for part_name in parts {
            if !(part_name.ends_with(".xml") || part_name.ends_with(".rels")) {
                continue;
            }
            let Ok(mut xml) = String::from_utf8(package.part(&part_name).unwrap().to_vec()) else {
                continue;
            };
            for (transitional, strict) in NAMESPACE_PAIRS {
                xml = xml.replace(transitional, strict);
            }
            package.set_part(&part_name, xml.into_bytes()).unwrap();
        }
        let mut editor = NativeOfficeEditor::from_package(package).unwrap();
        editor
            .set_hyperlink(owner, external("https://example.com/strict"))
            .unwrap();

        let owner_xml =
            String::from_utf8(editor.package().part(owner_part).unwrap().to_vec()).unwrap();
        assert!(owner_xml.contains(STRICT_RELATIONSHIPS));
        assert!(!owner_xml.contains(TRANSITIONAL_RELATIONSHIPS));
        let relationships =
            String::from_utf8(editor.package().part(relationship_part).unwrap().to_vec()).unwrap();
        assert!(relationships.contains(&format!("{STRICT_RELATIONSHIPS}/hyperlink")));
        assert!(!relationships.contains(&format!("{TRANSITIONAL_RELATIONSHIPS}/hyperlink")));

        if extension == "pptx" {
            editor
                .set_hyperlink(owner, NativeOfficeHyperlink::internal("slide[2]").unwrap())
                .unwrap();
            let relationships =
                String::from_utf8(editor.package().part(relationship_part).unwrap().to_vec())
                    .unwrap();
            assert!(relationships.contains(&format!("{STRICT_RELATIONSHIPS}/slide")));
            assert!(!relationships.contains(&format!("{TRANSITIONAL_RELATIONSHIPS}/slide")));
        }
    }
}
