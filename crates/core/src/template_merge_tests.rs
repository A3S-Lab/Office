use crate::{template_merge, DocumentKind, NativeOfficeEditor, NativeOfficePackage, PackageLimits};

fn package(kind: DocumentKind) -> NativeOfficePackage {
    NativeOfficePackage::blank_in_memory(kind, PackageLimits::default()).unwrap()
}

fn part_text(package: &NativeOfficePackage, name: &str) -> String {
    String::from_utf8(package.part(name).unwrap().to_vec()).unwrap()
}

#[test]
fn word_merge_preserves_split_run_formatting_and_processes_auxiliary_parts() {
    let mut package = package(DocumentKind::Word);
    package
        .set_part(
            "word/document.xml",
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Hello {{user.</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>name}}</w:t></w:r></w:p><w:p><w:r><w:t>{{missing}}</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#.to_vec(),
        )
        .unwrap();
    package
        .set_part(
            "word/header1.xml",
            br#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>{{header}} / {{alpha}}</w:t></w:r></w:p></w:hdr>"#.to_vec(),
        )
        .unwrap();
    package
        .set_part(
            "word/footnotes.xml",
            br#"<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:footnote w:id="1"><w:p><w:r><w:t>{{footnote}}</w:t></w:r></w:p></w:footnote></w:footnotes>"#.to_vec(),
        )
        .unwrap();
    package
        .set_part(
            "word/endnotes.xml",
            br#"<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:endnote w:id="1"><w:p><w:r><w:t>{{endnote}}</w:t></w:r></w:p></w:endnote></w:endnotes>"#.to_vec(),
        )
        .unwrap();
    package
        .set_part(
            "word/comments.xml",
            br#"<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="0"><w:p><w:r><w:t>{{comment}}</w:t></w:r></w:p></w:comment></w:comments>"#.to_vec(),
        )
        .unwrap();
    package
        .set_part(
            "word/footer1.xml",
            br#"<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>{{footer}}</w:t></w:r></w:p></w:ftr>"#.to_vec(),
        )
        .unwrap();

    let result = template_merge::merge(
        &mut package,
        &serde_json::json!({
            "user": {"name": "Alice"},
            "header": "Confidential",
            "footer": "Page 1",
            "footnote": "Footnote text",
            "endnote": "Endnote text",
            "comment": "Comment text"
        }),
    )
    .unwrap();

    assert_eq!(result.replaced_count, 6);
    assert_eq!(
        result.used_keys,
        [
            "comment",
            "endnote",
            "footer",
            "footnote",
            "header",
            "user.name"
        ]
    );
    assert_eq!(result.unresolved_placeholders, ["alpha", "missing"]);
    assert_eq!(
        result.changed_parts,
        [
            "/word/comments.xml",
            "/word/document.xml",
            "/word/endnotes.xml",
            "/word/footer1.xml",
            "/word/footnotes.xml",
            "/word/header1.xml"
        ]
    );
    let document = part_text(&package, "word/document.xml");
    assert!(document.contains("<w:rPr><w:b/></w:rPr><w:t>Hello Alice</w:t>"));
    assert!(document.contains("<w:rPr><w:i/></w:rPr><w:t></w:t>"));
    assert!(document.contains("{{missing}}"));
    assert!(part_text(&package, "word/header1.xml").contains("Confidential"));
    assert!(part_text(&package, "word/footer1.xml").contains("Page 1"));
    assert!(part_text(&package, "word/footnotes.xml").contains("Footnote text"));
    assert!(part_text(&package, "word/endnotes.xml").contains("Endnote text"));
    assert!(part_text(&package, "word/comments.xml").contains("Comment text"));
}

#[test]
fn spreadsheet_merge_preserves_rich_strings_and_counts_shared_references() {
    let mut package = package(DocumentKind::Spreadsheet);
    package
        .set_part(
            "xl/worksheets/sheet1.xml",
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="s"><v>0</v></c><c r="A2" t="s"><v>0</v></c><c r="B1" t="inlineStr"><is><r><rPr><b/></rPr><t>{{inline</t></r><r><rPr><i/></rPr><t>}}</t></r><rPh sb="0" eb="1"><t>{{phonetic}}</t></rPh></is></c><c r="C1" t="str"><v>{{direct}}</v></c><c r="D1" t="inlineStr"><is><t>{{later}}</t></is></c><c r="E1"><v>42</v></c></row></sheetData></worksheet>"#.to_vec(),
        )
        .unwrap();
    package
        .set_part(
            "xl/sharedStrings.xml",
            br#"<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="2" uniqueCount="1"><si><r><rPr><b/></rPr><t>{{shared.</t></r><r><rPr><i/></rPr><t>name}}</t></r><rPh sb="0" eb="1"><t>{{phonetic}}</t></rPh></si></sst>"#.to_vec(),
        )
        .unwrap();

    let result = template_merge::merge(
        &mut package,
        &serde_json::json!({
            "shared": {"name": "Revenue"},
            "inline": "North",
            "direct": "Forecast"
        }),
    )
    .unwrap();

    assert_eq!(result.replaced_count, 4);
    assert_eq!(result.used_keys, ["direct", "inline", "shared.name"]);
    assert_eq!(result.unresolved_placeholders, ["later"]);
    assert_eq!(
        result.changed_parts,
        ["/xl/sharedStrings.xml", "/xl/worksheets/sheet1.xml"]
    );
    let shared = part_text(&package, "xl/sharedStrings.xml");
    assert!(shared.contains("<rPr><b/></rPr><t>Revenue</t>"));
    assert!(shared.contains("<rPr><i/></rPr><t></t>"));
    assert!(shared.contains("<rPh sb=\"0\" eb=\"1\"><t>{{phonetic}}</t></rPh>"));
    let worksheet = part_text(&package, "xl/worksheets/sheet1.xml");
    assert!(worksheet.contains("<rPr><b/></rPr><t>North</t>"));
    assert!(worksheet.contains("<rPr><i/></rPr><t></t>"));
    assert!(worksheet.contains("<c r=\"C1\" t=\"str\"><v>Forecast</v></c>"));
    assert!(worksheet.contains("<c r=\"E1\"><v>42</v></c>"));
}

#[test]
fn spreadsheet_merge_fails_closed_for_non_string_cell_values() {
    let mut package = package(DocumentKind::Spreadsheet);
    package
        .set_part(
            "xl/worksheets/sheet1.xml",
            br#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="n"><v>{{amount}}</v></c></row></sheetData></worksheet>"#.to_vec(),
        )
        .unwrap();
    let before = package.content_sha256();

    let error =
        template_merge::merge(&mut package, &serde_json::json!({"amount": 42})).unwrap_err();

    assert_eq!(error.code, "use.office.template_cell_type_unsupported");
    assert_eq!(error.details["cell"], "A1");
    assert_eq!(package.content_sha256(), before);
}

#[test]
fn presentation_merge_handles_split_runs_on_slides_and_notes() {
    let mut package = package(DocumentKind::Presentation);
    package
        .set_part(
            "ppt/slides/slide1.xml",
            br#"<p:sld xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:rPr b="1"/><a:t>{{title.</a:t></a:r><a:r><a:rPr i="1"/><a:t>text}}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>"#.to_vec(),
        )
        .unwrap();
    package
        .set_part(
            "ppt/notesSlides/notesSlide1.xml",
            br#"<p:notes xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:cSld><p:spTree><p:sp><p:txBody><a:p><a:r><a:t>{{note}}</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:notes>"#.to_vec(),
        )
        .unwrap();

    let result = template_merge::merge(
        &mut package,
        &serde_json::json!({"title": {"text": "Q3"}, "note": "Internal"}),
    )
    .unwrap();

    assert_eq!(result.replaced_count, 2);
    assert_eq!(result.used_keys, ["note", "title.text"]);
    assert!(result.unresolved_placeholders.is_empty());
    let slide = part_text(&package, "ppt/slides/slide1.xml");
    assert!(slide.contains("<a:rPr b=\"1\"/><a:t>Q3</a:t>"));
    assert!(slide.contains("<a:rPr i=\"1\"/><a:t></a:t>"));
    assert!(part_text(&package, "ppt/notesSlides/notesSlide1.xml").contains("Internal"));
}

#[test]
fn editor_rolls_back_all_parts_when_a_template_value_is_not_valid_xml_text() {
    let mut package = package(DocumentKind::Word);
    package
        .set_part(
            "word/document.xml",
            br#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:r><w:t>{{safe}}</w:t></w:r></w:p><w:sectPr/></w:body></w:document>"#.to_vec(),
        )
        .unwrap();
    package
        .set_part(
            "word/header1.xml",
            br#"<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>{{bad}}</w:t></w:r></w:p></w:hdr>"#.to_vec(),
        )
        .unwrap();
    let mut editor = NativeOfficeEditor::from_package(package).unwrap();
    let before = editor.package().content_sha256();

    let error = editor
        .merge_template(&serde_json::json!({"safe": "changed", "bad": "\u{0}"}))
        .unwrap_err();

    assert_eq!(error.code, "use.office.template_value_invalid");
    assert_eq!(editor.package().content_sha256(), before);
    assert!(part_text(editor.package(), "word/document.xml").contains("{{safe}}"));
}
