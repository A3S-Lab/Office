use crate::{
    NativeOfficeComment, NativeOfficeCommentPosition, NativeOfficeCommentUpdate,
    NativeOfficeEditor, NativeOfficeMutation, OfficeNodeType, RelationshipSource,
};

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

fn part_names(editor: &NativeOfficeEditor, prefix: &str) -> Vec<String> {
    editor
        .package()
        .part_names()
        .filter(|part| part.starts_with(prefix))
        .map(str::to_string)
        .collect()
}

#[test]
fn comment_mutations_have_typed_stable_json_contracts() {
    let add = NativeOfficeMutation::AddComment {
        parent: "/slide[1]".into(),
        comment: NativeOfficeComment::new("Alice", "Review this")
            .unwrap()
            .with_initials("AL")
            .with_position(NativeOfficeCommentPosition::new(914_400, 457_200)),
    };
    assert_eq!(
        serde_json::to_value(&add).unwrap(),
        serde_json::json!({
            "operation": "add-comment",
            "parent": "/slide[1]",
            "comment": {
                "author": "Alice",
                "text": "Review this",
                "initials": "AL",
                "position": { "xEmu": 914400, "yEmu": 457200 }
            }
        })
    );

    let set = NativeOfficeMutation::SetComment {
        path: "/slide[1]/comment[1]".into(),
        update: NativeOfficeCommentUpdate {
            text: Some("Updated".into()),
            ..NativeOfficeCommentUpdate::default()
        },
    };
    let encoded = serde_json::to_value(&set).unwrap();
    assert_eq!(
        encoded,
        serde_json::json!({
            "operation": "set-comment",
            "path": "/slide[1]/comment[1]",
            "update": { "text": "Updated" }
        })
    );
    assert_eq!(
        serde_json::from_value::<NativeOfficeMutation>(encoded).unwrap(),
        set
    );
}

#[tokio::test]
async fn native_word_adds_updates_reads_queries_and_removes_comments() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("comments.docx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor
        .set_text("/body/p[1]", "Review this paragraph")
        .unwrap();

    let comment_path = editor
        .add_comment(
            "/body/p[1]",
            NativeOfficeComment::new("Alice", "Please reword this")
                .unwrap()
                .with_initials("AL"),
        )
        .unwrap();
    assert_eq!(comment_path, "/comments/comment[1]");
    let node = editor.snapshot().unwrap().get(&comment_path, 0).unwrap();
    assert_eq!(node.node_type, OfficeNodeType::Comment);
    assert_eq!(node.text, "Please reword this");
    assert_eq!(node.format["author"], "Alice");
    assert_eq!(node.format["initials"], "AL");
    assert_eq!(node.format["anchoredTo"], "/body/p[1]");
    assert_eq!(
        editor.snapshot().unwrap().query("comment").unwrap().len(),
        1
    );
    assert_eq!(
        relationship_count(&editor, "word/document.xml", "/comments"),
        1
    );

    editor
        .set_comment(
            &comment_path,
            NativeOfficeCommentUpdate {
                author: Some("Bob".into()),
                text: Some("Looks good now".into()),
                initials: Some("BO".into()),
                position: None,
            },
        )
        .unwrap();
    let updated = editor.snapshot().unwrap().get(&comment_path, 0).unwrap();
    assert_eq!(updated.text, "Looks good now");
    assert_eq!(updated.format["author"], "Bob");
    assert_eq!(updated.format["initials"], "BO");

    editor.remove(&comment_path).unwrap();
    assert!(editor
        .snapshot()
        .unwrap()
        .query("comment")
        .unwrap()
        .is_empty());
    assert_eq!(
        relationship_count(&editor, "word/document.xml", "/comments"),
        0
    );
    assert!(part_names(&editor, "word/comments").is_empty());
}

#[tokio::test]
async fn native_spreadsheet_manages_classic_cell_comments_and_vml_notes() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("comments.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();

    let comment_path = editor
        .add_comment(
            "/Sheet1/B2",
            NativeOfficeComment::new("Alice", "Check this formula").unwrap(),
        )
        .unwrap();
    assert_eq!(comment_path, "/Sheet1/B2/comment");
    let node = editor.snapshot().unwrap().get(&comment_path, 0).unwrap();
    assert_eq!(node.node_type, OfficeNodeType::Comment);
    assert_eq!(node.text, "Check this formula");
    assert_eq!(node.format["author"], "Alice");
    assert_eq!(node.format["ref"], "B2");
    assert_eq!(
        relationship_count(&editor, "xl/worksheets/sheet1.xml", "/comments"),
        1
    );
    assert_eq!(
        relationship_count(&editor, "xl/worksheets/sheet1.xml", "/vmlDrawing"),
        1
    );
    assert_eq!(part_names(&editor, "xl/comments").len(), 1);
    assert_eq!(part_names(&editor, "xl/drawings/vmlDrawing").len(), 1);

    editor
        .set_comment(
            &comment_path,
            NativeOfficeCommentUpdate {
                author: Some("Bob".into()),
                text: Some("Formula checked".into()),
                ..NativeOfficeCommentUpdate::default()
            },
        )
        .unwrap();
    let updated = editor.snapshot().unwrap().get(&comment_path, 0).unwrap();
    assert_eq!(updated.text, "Formula checked");
    assert_eq!(updated.format["author"], "Bob");

    editor.remove(&comment_path).unwrap();
    assert!(editor
        .snapshot()
        .unwrap()
        .query("comment")
        .unwrap()
        .is_empty());
    assert_eq!(
        relationship_count(&editor, "xl/worksheets/sheet1.xml", "/comments"),
        0
    );
    assert_eq!(
        relationship_count(&editor, "xl/worksheets/sheet1.xml", "/vmlDrawing"),
        0
    );
    assert!(part_names(&editor, "xl/comments").is_empty());
    assert!(part_names(&editor, "xl/drawings/vmlDrawing").is_empty());
}

#[tokio::test]
async fn native_presentation_manages_legacy_slide_comments_and_shared_authors() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("comments.pptx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor.add_slide("/", "Review").unwrap();

    let comment_path = editor
        .add_comment(
            "/slide[1]",
            NativeOfficeComment::new("Alice", "Reword this slide")
                .unwrap()
                .with_initials("AL")
                .with_position(NativeOfficeCommentPosition::new(914_400, 457_200)),
        )
        .unwrap();
    assert_eq!(comment_path, "/slide[1]/comment[1]");
    let node = editor.snapshot().unwrap().get(&comment_path, 0).unwrap();
    assert_eq!(node.node_type, OfficeNodeType::Comment);
    assert_eq!(node.text, "Reword this slide");
    assert_eq!(node.format["author"], "Alice");
    assert_eq!(node.format["initials"], "AL");
    assert_eq!(node.format["xEmu"], "914400");
    assert_eq!(node.format["yEmu"], "457200");
    assert_eq!(
        relationship_count(&editor, "ppt/slides/slide1.xml", "/comments"),
        1
    );
    assert_eq!(
        relationship_count(&editor, "ppt/presentation.xml", "/commentAuthors"),
        1
    );

    editor
        .set_comment(
            &comment_path,
            NativeOfficeCommentUpdate {
                author: Some("Bob".into()),
                text: Some("Updated review".into()),
                initials: Some("BO".into()),
                position: Some(NativeOfficeCommentPosition::new(1_828_800, 914_400)),
            },
        )
        .unwrap();
    let updated = editor.snapshot().unwrap().get(&comment_path, 0).unwrap();
    assert_eq!(updated.text, "Updated review");
    assert_eq!(updated.format["author"], "Bob");
    assert_eq!(updated.format["xEmu"], "1828800");

    editor.remove(&comment_path).unwrap();
    assert!(editor
        .snapshot()
        .unwrap()
        .query("comment")
        .unwrap()
        .is_empty());
    assert_eq!(
        relationship_count(&editor, "ppt/slides/slide1.xml", "/comments"),
        0
    );
    assert!(part_names(&editor, "ppt/comments/comment").is_empty());
}

#[tokio::test]
async fn comment_validation_and_format_rejections_are_atomic() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("atomic.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    let before = editor.package().content_sha256();

    let invalid = NativeOfficeComment {
        author: "Alice".into(),
        text: "bad\0text".into(),
        initials: None,
        position: None,
    };
    assert_eq!(
        editor.add_comment("/Sheet1/A1", invalid).unwrap_err().code,
        "use.office.comment_text_invalid"
    );
    assert_eq!(editor.package().content_sha256(), before);

    let positioned = NativeOfficeComment::new("Alice", "No coordinates here")
        .unwrap()
        .with_position(NativeOfficeCommentPosition::new(1, 2));
    assert_eq!(
        editor
            .add_comment("/Sheet1/A1", positioned)
            .unwrap_err()
            .code,
        "use.office.comment_position_unsupported"
    );
    assert_eq!(editor.package().content_sha256(), before);
}

#[tokio::test]
async fn comment_batches_and_owner_removal_are_atomic_and_garbage_collect_parts() {
    let temp = tempfile::tempdir().unwrap();

    let word_path = temp.path().join("owners.docx");
    let mut word = NativeOfficeEditor::create(&word_path).await.unwrap();
    let paragraph = word.add_paragraph("/body", "Owned").unwrap();
    word.add_comment(
        &paragraph,
        NativeOfficeComment::new("Alice", "Owned comment").unwrap(),
    )
    .unwrap();
    word.remove(&paragraph).unwrap();
    assert!(word
        .snapshot()
        .unwrap()
        .query("comment")
        .unwrap()
        .is_empty());
    assert_eq!(
        relationship_count(&word, "word/document.xml", "/comments"),
        0
    );

    word.set_text("/body/p[1]", "Owned run").unwrap();
    for text in ["First run comment", "Second run comment"] {
        word.add_comment(
            "/body/p[1]/r[1]",
            NativeOfficeComment::new("Alice", text).unwrap(),
        )
        .unwrap();
    }
    assert_eq!(word.snapshot().unwrap().query("comment").unwrap().len(), 2);
    word.remove("/body/p[1]/r[1]").unwrap();
    assert!(word
        .snapshot()
        .unwrap()
        .query("comment")
        .unwrap()
        .is_empty());
    assert_eq!(
        relationship_count(&word, "word/document.xml", "/comments"),
        0
    );

    let sheet_path = temp.path().join("owners.xlsx");
    let mut sheet = NativeOfficeEditor::create(&sheet_path).await.unwrap();
    sheet.set_text("/Sheet1/A1", "Owned").unwrap();
    sheet
        .add_comment(
            "/Sheet1/A1",
            NativeOfficeComment::new("Alice", "Owned note").unwrap(),
        )
        .unwrap();
    sheet.remove("/Sheet1/A1").unwrap();
    assert!(sheet
        .snapshot()
        .unwrap()
        .query("comment")
        .unwrap()
        .is_empty());
    assert_eq!(
        relationship_count(&sheet, "xl/worksheets/sheet1.xml", "/comments"),
        0
    );
    assert_eq!(
        relationship_count(&sheet, "xl/worksheets/sheet1.xml", "/vmlDrawing"),
        0
    );

    let deck_path = temp.path().join("owners.pptx");
    let mut deck = NativeOfficeEditor::create(&deck_path).await.unwrap();
    let slide = deck.add_slide("/", "Owned").unwrap();
    deck.add_comment(
        &slide,
        NativeOfficeComment::new("Alice", "Owned review").unwrap(),
    )
    .unwrap();
    deck.remove(&slide).unwrap();
    assert!(deck
        .snapshot()
        .unwrap()
        .query("comment")
        .unwrap()
        .is_empty());
    assert!(part_names(&deck, "ppt/comments/comment").is_empty());

    let rollback_path = temp.path().join("rollback.docx");
    let mut rollback = NativeOfficeEditor::create(&rollback_path).await.unwrap();
    let before = rollback.package().content_sha256();
    let error = rollback
        .apply_batch(&[
            NativeOfficeMutation::AddComment {
                parent: "/body/p[1]".into(),
                comment: NativeOfficeComment::new("Alice", "Transient").unwrap(),
            },
            NativeOfficeMutation::SetComment {
                path: "/comments/comment[1]".into(),
                update: NativeOfficeCommentUpdate::default(),
            },
        ])
        .unwrap_err();
    assert_eq!(error.code, "use.office.comment_update_empty");
    assert_eq!(rollback.package().content_sha256(), before);
    assert!(rollback
        .snapshot()
        .unwrap()
        .query("comment")
        .unwrap()
        .is_empty());
}

#[tokio::test]
async fn presentation_comment_authors_are_reused_and_forked_without_sibling_poisoning() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("authors.pptx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor.add_slide("/", "Authors").unwrap();
    let first = editor
        .add_comment(
            "/slide[1]",
            NativeOfficeComment::new("Alice", "First")
                .unwrap()
                .with_initials("AL"),
        )
        .unwrap();
    let second = editor
        .add_comment(
            "/slide[1]",
            NativeOfficeComment::new("Alice", "Second")
                .unwrap()
                .with_initials("AL"),
        )
        .unwrap();
    let author_part = part_names(&editor, "ppt/commentAuthors")
        .into_iter()
        .next()
        .unwrap();
    let authors = String::from_utf8(editor.package().part(&author_part).unwrap().to_vec()).unwrap();
    assert_eq!(authors.matches("cmAuthor ").count(), 1);
    assert_eq!(
        editor.snapshot().unwrap().get(&first, 0).unwrap().format["index"],
        "1"
    );
    assert_eq!(
        editor.snapshot().unwrap().get(&second, 0).unwrap().format["index"],
        "2"
    );

    editor
        .set_comment(
            &first,
            NativeOfficeCommentUpdate {
                author: Some("Bob".into()),
                initials: Some("BO".into()),
                ..NativeOfficeCommentUpdate::default()
            },
        )
        .unwrap();
    let snapshot = editor.snapshot().unwrap();
    assert_eq!(snapshot.get(&first, 0).unwrap().format["author"], "Bob");
    assert_eq!(snapshot.get(&second, 0).unwrap().format["author"], "Alice");
}

#[tokio::test]
async fn comment_updates_preserve_unknown_attributes_and_extension_nodes() {
    let temp = tempfile::tempdir().unwrap();

    let word_path = temp.path().join("preserve-comments.docx");
    let mut word = NativeOfficeEditor::create(&word_path).await.unwrap();
    let word_comment = word
        .add_comment(
            "/body/p[1]",
            NativeOfficeComment::new("Alice", "Original word comment").unwrap(),
        )
        .unwrap();
    let word_part = part_names(&word, "word/comments")
        .into_iter()
        .next()
        .unwrap();
    let mut package = word.package().clone();
    let word_xml = String::from_utf8(package.part(&word_part).unwrap().to_vec())
        .unwrap()
        .replacen(
            "<w:comments ",
            "<w:comments xmlns:x=\"urn:a3s:unknown\" ",
            1,
        )
        .replacen("<w:comment ", "<w:comment x:keep=\"word\" ", 1)
        .replacen(
            "<w:t>Original word comment</w:t>",
            "<z:t xmlns:z=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">Original word comment</z:t>",
            1,
        )
        .replacen(
            "</w:comment>",
            "<x:opaque x:value=\"word\"><x:t>keep-word</x:t></x:opaque></w:comment>",
            1,
        );
    package.set_part(&word_part, word_xml.into_bytes()).unwrap();
    let mut word = NativeOfficeEditor::from_package(package).unwrap();
    word.set_comment(
        &word_comment,
        NativeOfficeCommentUpdate {
            author: Some("Bob".into()),
            text: Some("Updated word comment".into()),
            ..NativeOfficeCommentUpdate::default()
        },
    )
    .unwrap();
    let word_xml = String::from_utf8(word.package().part(&word_part).unwrap().to_vec()).unwrap();
    assert!(word_xml.contains("x:keep=\"word\""));
    assert!(word_xml.contains(
        "<z:t xmlns:z=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">Updated word comment</z:t>"
    ));
    assert!(word_xml.contains("<x:opaque x:value=\"word\"><x:t>keep-word</x:t></x:opaque>"));

    let sheet_path = temp.path().join("preserve-comments.xlsx");
    let mut sheet = NativeOfficeEditor::create(&sheet_path).await.unwrap();
    let sheet_comment = sheet
        .add_comment(
            "/Sheet1/B2",
            NativeOfficeComment::new("Alice", "Original sheet comment").unwrap(),
        )
        .unwrap();
    let sheet_part = part_names(&sheet, "xl/comments")
        .into_iter()
        .next()
        .unwrap();
    let mut package = sheet.package().clone();
    let sheet_xml = String::from_utf8(package.part(&sheet_part).unwrap().to_vec())
        .unwrap()
        .replacen("<comments ", "<comments xmlns:x=\"urn:a3s:unknown\" ", 1)
        .replacen("<comment ", "<comment x:keep=\"sheet\" ", 1)
        .replacen(
            "</comment>",
            "<x:opaque x:value=\"sheet\"><x:t>keep-sheet</x:t></x:opaque></comment>",
            1,
        );
    package
        .set_part(&sheet_part, sheet_xml.into_bytes())
        .unwrap();
    let mut sheet = NativeOfficeEditor::from_package(package).unwrap();
    sheet
        .set_comment(
            &sheet_comment,
            NativeOfficeCommentUpdate {
                author: Some("Bob".into()),
                text: Some("Updated sheet comment".into()),
                ..NativeOfficeCommentUpdate::default()
            },
        )
        .unwrap();
    let sheet_xml = String::from_utf8(sheet.package().part(&sheet_part).unwrap().to_vec()).unwrap();
    assert!(sheet_xml.contains("x:keep=\"sheet\""));
    assert!(sheet_xml.contains("<x:opaque x:value=\"sheet\"><x:t>keep-sheet</x:t></x:opaque>"));

    let deck_path = temp.path().join("preserve-comments.pptx");
    let mut deck = NativeOfficeEditor::create(&deck_path).await.unwrap();
    deck.add_slide("/", "Review").unwrap();
    let deck_comment = deck
        .add_comment(
            "/slide[1]",
            NativeOfficeComment::new("Alice", "Original slide comment").unwrap(),
        )
        .unwrap();
    let deck_part = part_names(&deck, "ppt/comments/comment")
        .into_iter()
        .next()
        .unwrap();
    let mut package = deck.package().clone();
    let deck_xml = String::from_utf8(package.part(&deck_part).unwrap().to_vec())
        .unwrap()
        .replacen("<p:cmLst ", "<p:cmLst xmlns:x=\"urn:a3s:unknown\" ", 1)
        .replacen("<p:cm ", "<p:cm x:keep=\"deck\" ", 1)
        .replacen(
            "</p:cm>",
            "<x:opaque x:value=\"deck\"><x:text>keep-deck</x:text></x:opaque></p:cm>",
            1,
        );
    package.set_part(&deck_part, deck_xml.into_bytes()).unwrap();
    let mut deck = NativeOfficeEditor::from_package(package).unwrap();
    deck.set_comment(
        &deck_comment,
        NativeOfficeCommentUpdate {
            text: Some("Updated slide comment".into()),
            position: Some(NativeOfficeCommentPosition::new(100, 200)),
            ..NativeOfficeCommentUpdate::default()
        },
    )
    .unwrap();
    let deck_xml = String::from_utf8(deck.package().part(&deck_part).unwrap().to_vec()).unwrap();
    assert!(deck_xml.contains("x:keep=\"deck\""));
    assert!(deck_xml.contains("<x:opaque x:value=\"deck\"><x:text>keep-deck</x:text></x:opaque>"));
}

#[tokio::test]
async fn native_comments_preserve_strict_ooxml_dialects() {
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
    for (extension, parent, owner_part, relationship_part, created_prefix) in [
        (
            "docx",
            "/body/p[1]",
            "word/document.xml",
            "word/_rels/document.xml.rels",
            "word/comments",
        ),
        (
            "xlsx",
            "/Sheet1/A1",
            "xl/worksheets/sheet1.xml",
            "xl/worksheets/_rels/sheet1.xml.rels",
            "xl/comments",
        ),
        (
            "pptx",
            "/slide[1]",
            "ppt/slides/slide1.xml",
            "ppt/slides/_rels/slide1.xml.rels",
            "ppt/comments/comment",
        ),
    ] {
        let path = temp.path().join(format!("strict-comments.{extension}"));
        let mut seed = NativeOfficeEditor::create(&path).await.unwrap();
        if extension == "pptx" {
            seed.add_slide("/", "Strict").unwrap();
        }
        let mut package = seed.package().clone();
        for part_name in package.part_names().map(str::to_string).collect::<Vec<_>>() {
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
            .add_comment(
                parent,
                NativeOfficeComment::new("Alice", "Strict comment").unwrap(),
            )
            .unwrap();
        let owner_xml =
            String::from_utf8(editor.package().part(owner_part).unwrap().to_vec()).unwrap();
        assert!(
            owner_xml.contains("http://purl.oclc.org/ooxml/"),
            "{extension}"
        );
        let relationships =
            String::from_utf8(editor.package().part(relationship_part).unwrap().to_vec()).unwrap();
        assert!(relationships.contains(&format!("{STRICT_RELATIONSHIPS}/comments")));
        let created = part_names(&editor, created_prefix)
            .into_iter()
            .next()
            .unwrap();
        let created_xml =
            String::from_utf8(editor.package().part(&created).unwrap().to_vec()).unwrap();
        assert!(
            created_xml.contains("http://purl.oclc.org/ooxml/"),
            "{extension}"
        );
        assert_eq!(
            editor.snapshot().unwrap().query("comment").unwrap().len(),
            1
        );
    }
}
