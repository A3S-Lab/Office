use crate::{
    DocumentKind, NativeOfficeEditor, NativeOfficeMutation, NativeOfficePackage,
    NativeOfficeTextMatchMode, NativeOfficeTextReplacement, PackageLimits,
};

fn package(kind: DocumentKind) -> NativeOfficePackage {
    NativeOfficePackage::blank_in_memory(kind, PackageLimits::default()).unwrap()
}

fn part_text(package: &NativeOfficePackage, name: &str) -> String {
    String::from_utf8(package.part(name).unwrap().to_vec()).unwrap()
}

#[test]
fn typed_replacement_contract_is_stable_and_send_sync() {
    fn assert_send_sync<T: Send + Sync>() {}
    assert_send_sync::<NativeOfficeTextReplacement>();
    assert_send_sync::<crate::NativeOfficeTextReplacementResult>();

    let mutation = NativeOfficeMutation::ReplaceText {
        path: "/body".into(),
        replacement: NativeOfficeTextReplacement::literal("before", "after").unwrap(),
    };
    assert_eq!(
        serde_json::to_value(&mutation).unwrap(),
        serde_json::json!({
            "operation": "replace-text",
            "path": "/body",
            "replacement": {
                "find": "before",
                "replace": "after",
                "mode": "literal"
            }
        })
    );
    let decoded: NativeOfficeMutation =
        serde_json::from_value(serde_json::to_value(mutation).unwrap()).unwrap();
    assert!(matches!(
        decoded,
        NativeOfficeMutation::ReplaceText {
            replacement: NativeOfficeTextReplacement {
                mode: NativeOfficeTextMatchMode::Literal,
                ..
            },
            ..
        }
    ));
}

#[test]
fn word_replacement_spans_runs_preserves_unknown_xml_and_processes_auxiliary_parts() {
    let mut package = package(DocumentKind::Word);
    package
        .set_part(
            "word/document.xml",
            br#"<w:document xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main" xmlns:x="urn:test"><w:body><w:p x:keep="yes"><w:r><w:rPr><w:b/></w:rPr><w:t>alpha </w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>beta and alpha beta</w:t><x:ext value="keep"><x:t>alpha beta extension</x:t></x:ext></w:r></w:p><w:sectPr/></w:body></w:document>"#.to_vec(),
        )
        .unwrap();
    package
        .set_part(
            "word/header1.xml",
            br#"<w:hdr xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"><w:p><w:r><w:t>header target and root target</w:t></w:r></w:p></w:hdr>"#.to_vec(),
        )
        .unwrap();
    package
        .set_part(
            "word/footer1.xml",
            br#"<w:ftr xmlns:w="http://purl.oclc.org/ooxml/wordprocessingml/main"><w:p><w:r><w:t>footer target</w:t></w:r></w:p></w:ftr>"#.to_vec(),
        )
        .unwrap();
    crate::opc_edit::add_content_type_override(
        &mut package,
        "word/header1.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
    )
    .unwrap();
    crate::opc_edit::add_content_type_override(
        &mut package,
        "word/footer1.xml",
        "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
    )
    .unwrap();
    crate::opc_edit::add_relationship(
        &mut package,
        "word/_rels/document.xml.rels",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
        "header1.xml",
    )
    .unwrap();
    crate::opc_edit::add_relationship(
        &mut package,
        "word/_rels/document.xml.rels",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
        "footer1.xml",
    )
    .unwrap();
    let mut editor = NativeOfficeEditor::from_package(package).unwrap();

    let body = editor
        .replace_text(
            "/body",
            NativeOfficeTextReplacement::literal("alpha beta", "done").unwrap(),
        )
        .unwrap();
    assert_eq!(body.match_count, 2);
    assert!(body.changed);
    assert_eq!(body.changed_parts, ["/word/document.xml"]);
    let document = part_text(editor.package(), "word/document.xml");
    assert!(document.contains("<w:rPr><w:b/></w:rPr><w:t>done</w:t>"));
    assert!(document.contains("<w:rPr><w:i/></w:rPr><w:t xml:space=\"preserve\"> and done</w:t>"));
    assert!(document.contains("x:keep=\"yes\""));
    assert!(document.contains("<x:ext value=\"keep\"><x:t>alpha beta extension</x:t></x:ext>"));
    assert!(part_text(editor.package(), "word/header1.xml").contains("header target"));

    let run = editor
        .replace_text(
            "/body/p[1]/r[1]",
            NativeOfficeTextReplacement::literal("done", "run done").unwrap(),
        )
        .unwrap();
    assert_eq!(run.match_count, 1);
    let document = part_text(editor.package(), "word/document.xml");
    assert!(document.contains("<w:rPr><w:b/></w:rPr><w:t>run done</w:t>"));
    assert!(document.contains("<w:rPr><w:i/></w:rPr><w:t xml:space=\"preserve\"> and done</w:t>"));

    let header = editor
        .replace_text(
            "/header[1]/p[1]/r[1]",
            NativeOfficeTextReplacement::literal("header target", "header done").unwrap(),
        )
        .unwrap();
    assert_eq!(header.match_count, 1);
    assert_eq!(header.changed_parts, ["/word/header1.xml"]);
    assert!(part_text(editor.package(), "word/header1.xml").contains("header done"));

    let footer = editor
        .replace_text(
            "/footer[1]",
            NativeOfficeTextReplacement::literal("footer target", "footer done").unwrap(),
        )
        .unwrap();
    assert_eq!(footer.match_count, 1);
    assert_eq!(footer.changed_parts, ["/word/footer1.xml"]);
    assert!(part_text(editor.package(), "word/footer1.xml").contains("footer done"));

    let root = editor
        .replace_text(
            "/",
            NativeOfficeTextReplacement::literal("root target", "root done").unwrap(),
        )
        .unwrap();
    assert_eq!(root.match_count, 1);
    assert_eq!(root.changed_parts, ["/word/header1.xml"]);
    assert!(part_text(editor.package(), "word/header1.xml").contains("root done"));
}

#[test]
fn spreadsheet_scoped_replacement_clones_shared_strings_without_poisoning_siblings() {
    let mut package = package(DocumentKind::Spreadsheet);
    package
        .set_part(
            "xl/worksheets/sheet1.xml",
            br#"<worksheet xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="inlineStr"><is><r><rPr><b/></rPr><t>alpha </t></r><r><rPr><i/></rPr><t>beta</t></r></is></c><c r="C1"><v>42</v></c></row><row r="2"><c r="A2" t="s"><v>0</v></c></row></sheetData></worksheet>"#.to_vec(),
        )
        .unwrap();
    package
        .set_part(
            "xl/sharedStrings.xml",
            br#"<sst xmlns="http://purl.oclc.org/ooxml/spreadsheetml/main" xmlns:x="urn:test" count="2" uniqueCount="1"><si x:keep="yes"><r><rPr><b/></rPr><t>alpha </t></r><r><rPr><i/></rPr><t>beta</t></r><x:ext value="keep"><x:t>alpha beta extension</x:t></x:ext></si></sst>"#.to_vec(),
        )
        .unwrap();
    let mut editor = NativeOfficeEditor::from_package(package).unwrap();

    let result = editor
        .replace_text(
            "/Sheet1/A1",
            NativeOfficeTextReplacement::literal("alpha beta", "selected").unwrap(),
        )
        .unwrap();
    assert_eq!(result.match_count, 1);
    assert_eq!(
        result.changed_parts,
        ["/xl/sharedStrings.xml", "/xl/worksheets/sheet1.xml"]
    );
    let snapshot = editor.snapshot().unwrap();
    assert_eq!(snapshot.get("/Sheet1/A1", 0).unwrap().text, "selected");
    assert_eq!(snapshot.get("/Sheet1/A2", 0).unwrap().text, "alpha beta");
    let shared = part_text(editor.package(), "xl/sharedStrings.xml");
    assert_eq!(shared.matches("<si").count(), 2);
    assert!(!shared.contains("uniqueCount="));
    assert!(shared.contains("x:keep=\"yes\""));
    assert_eq!(
        shared
            .matches("<x:ext value=\"keep\"><x:t>alpha beta extension</x:t></x:ext>")
            .count(),
        2
    );
    let worksheet = part_text(editor.package(), "xl/worksheets/sheet1.xml");
    assert!(worksheet.contains("<c r=\"A1\" t=\"s\"><v>1</v></c>"));
    assert!(worksheet.contains("<c r=\"A2\" t=\"s\"><v>0</v></c>"));

    let inline = editor
        .replace_text(
            "/Sheet1/B1:C1",
            NativeOfficeTextReplacement::regex(r"alpha\s+beta", "inline").unwrap(),
        )
        .unwrap();
    assert_eq!(inline.match_count, 1);
    assert_eq!(inline.changed_parts, ["/xl/worksheets/sheet1.xml"]);
    assert_eq!(
        editor
            .snapshot()
            .unwrap()
            .get("/Sheet1/B1", 0)
            .unwrap()
            .text,
        "inline"
    );
    assert_eq!(
        editor
            .snapshot()
            .unwrap()
            .get("/Sheet1/C1", 0)
            .unwrap()
            .text,
        "42"
    );
}

#[test]
fn spreadsheet_scoped_replacement_assigns_consecutive_indices_to_multiple_shared_clones() {
    let mut package = package(DocumentKind::Spreadsheet);
    package
        .set_part(
            "xl/worksheets/sheet1.xml",
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="B1" t="s"><v>1</v></c></row><row r="2"><c r="A2" t="s"><v>0</v></c><c r="B2" t="s"><v>1</v></c></row></sheetData></worksheet>"#.to_vec(),
        )
        .unwrap();
    package
        .set_part(
            "xl/sharedStrings.xml",
            br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="4" uniqueCount="2"><si><t>alpha one</t></si><si><t>alpha two</t></si></sst>"#.to_vec(),
        )
        .unwrap();
    let mut editor = NativeOfficeEditor::from_package(package).unwrap();

    let result = editor
        .replace_text(
            "/Sheet1/A1:B1",
            NativeOfficeTextReplacement::literal("alpha", "selected").unwrap(),
        )
        .unwrap();

    assert_eq!(result.match_count, 2);
    let worksheet = part_text(editor.package(), "xl/worksheets/sheet1.xml");
    assert!(worksheet.contains("<c r=\"A1\" t=\"s\"><v>2</v></c>"));
    assert!(worksheet.contains("<c r=\"B1\" t=\"s\"><v>3</v></c>"));
    assert!(worksheet.contains("<c r=\"A2\" t=\"s\"><v>0</v></c>"));
    assert!(worksheet.contains("<c r=\"B2\" t=\"s\"><v>1</v></c>"));

    let shared = part_text(editor.package(), "xl/sharedStrings.xml");
    assert_eq!(shared.matches("<si").count(), 4);
    assert!(shared.contains("<si><t>selected one</t></si>"));
    assert!(shared.contains("<si><t>selected two</t></si>"));
    let snapshot = editor.snapshot().unwrap();
    assert_eq!(snapshot.get("/Sheet1/A1", 0).unwrap().text, "selected one");
    assert_eq!(snapshot.get("/Sheet1/B1", 0).unwrap().text, "selected two");
    assert_eq!(snapshot.get("/Sheet1/A2", 0).unwrap().text, "alpha one");
    assert_eq!(snapshot.get("/Sheet1/B2", 0).unwrap().text, "alpha two");
}

#[test]
fn presentation_replacement_supports_shape_runs_and_related_notes() {
    let mut seed = NativeOfficeEditor::from_package(package(DocumentKind::Presentation)).unwrap();
    seed.add_slide("/", "Seed").unwrap();
    let mut package = seed.package().clone();
    package
        .set_part(
            "ppt/slides/slide1.xml",
            br#"<p:sld xmlns:p="http://purl.oclc.org/ooxml/presentationml/main" xmlns:a="http://purl.oclc.org/ooxml/drawingml/main" xmlns:x="urn:test"><p:cSld><p:spTree><p:nvGrpSpPr/><p:grpSpPr/><p:sp x:keep="yes"><p:nvSpPr><p:cNvPr id="2" name="Text"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr b="1"/><a:t>alpha </a:t></a:r><a:r><a:rPr i="1"/><a:t>beta</a:t><x:ext value="keep"><x:t>alpha beta extension</x:t></x:ext></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#.to_vec(),
        )
        .unwrap();
    package
        .set_part(
            "ppt/notesSlides/notesSlide1.xml",
            br#"<p:notes xmlns:p="http://purl.oclc.org/ooxml/presentationml/main" xmlns:a="http://purl.oclc.org/ooxml/drawingml/main"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>alpha beta note</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>"#.to_vec(),
        )
        .unwrap();
    crate::opc_edit::add_content_type_override(
        &mut package,
        "ppt/notesSlides/notesSlide1.xml",
        "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
    )
    .unwrap();
    crate::opc_edit::add_relationship(
        &mut package,
        "ppt/slides/_rels/slide1.xml.rels",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide",
        "../notesSlides/notesSlide1.xml",
    )
    .unwrap();
    let mut editor = NativeOfficeEditor::from_package(package).unwrap();

    let shape = editor
        .replace_text(
            "/slide[1]/shape[1]",
            NativeOfficeTextReplacement::literal("alpha beta", "slide").unwrap(),
        )
        .unwrap();
    assert_eq!(shape.match_count, 1);
    assert_eq!(shape.changed_parts, ["/ppt/slides/slide1.xml"]);
    let slide = part_text(editor.package(), "ppt/slides/slide1.xml");
    assert!(slide.contains("<a:rPr b=\"1\"/><a:t>slide</a:t>"));
    assert!(slide.contains("<a:rPr i=\"1\"/><a:t></a:t>"));
    assert!(slide.contains("x:keep=\"yes\""));
    assert!(slide.contains("<x:ext value=\"keep\"><x:t>alpha beta extension</x:t></x:ext>"));

    let notes = editor
        .replace_text(
            "/slide[1]/notes",
            NativeOfficeTextReplacement::regex(r"alpha beta (note)", "$1 updated").unwrap(),
        )
        .unwrap();
    assert_eq!(notes.match_count, 1);
    assert_eq!(notes.changed_parts, ["/ppt/notesSlides/notesSlide1.xml"]);
    assert!(
        part_text(editor.package(), "ppt/notesSlides/notesSlide1.xml").contains("note updated")
    );
}

#[test]
fn zero_matches_are_explicit_and_failed_batches_roll_back_prior_replacements() {
    let mut package = package(DocumentKind::Word);
    package
        .set_part(
            "word/document.xml",
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>keep alpha beta</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#.to_vec(),
        )
        .unwrap();
    let mut editor = NativeOfficeEditor::from_package(package).unwrap();
    let before = editor.package().content_sha256();
    let zero = editor
        .replace_text(
            "/body",
            NativeOfficeTextReplacement::literal("missing", "unused").unwrap(),
        )
        .unwrap();
    assert_eq!(zero.match_count, 0);
    assert!(!zero.changed);
    assert!(zero.changed_parts.is_empty());
    assert_eq!(editor.package().content_sha256(), before);

    let construction_error = NativeOfficeTextReplacement::regex(r"a*", "unused").unwrap_err();
    assert_eq!(construction_error.code, "use.office.text_regex_empty_match");

    let runtime_error = editor
        .replace_text(
            "/body",
            NativeOfficeTextReplacement::regex(r"\b", "boundary").unwrap(),
        )
        .unwrap_err();
    assert_eq!(runtime_error.code, "use.office.text_regex_empty_match");
    assert_eq!(editor.package().content_sha256(), before);

    let error = editor
        .apply_batch(&[
            NativeOfficeMutation::ReplaceText {
                path: "/body".into(),
                replacement: NativeOfficeTextReplacement::literal("alpha beta", "changed").unwrap(),
            },
            NativeOfficeMutation::SetText {
                path: "/body/missing[1]".into(),
                text: "fail".into(),
            },
        ])
        .unwrap_err();
    assert_eq!(error.code, "use.office.mutation_path_unsupported");
    assert_eq!(editor.package().content_sha256(), before);
}
