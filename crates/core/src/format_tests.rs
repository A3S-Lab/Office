use super::{
    NativeOfficeEditor, NativeOfficeHighlightColor, NativeOfficeHorizontalAlignment,
    NativeOfficeMutation, NativeOfficeRgbColor, NativeOfficeTextCase, NativeOfficeTextFormat,
    NativeOfficeTextScript, NativeOfficeUnderline,
};

fn rich_format() -> NativeOfficeTextFormat {
    NativeOfficeTextFormat {
        bold: Some(true),
        italic: Some(false),
        underline: Some(NativeOfficeUnderline::Double),
        script: Some(NativeOfficeTextScript::Superscript),
        strikethrough: None,
        double_strikethrough: None,
        text_case: None,
        highlight: None,
        language: None,
        font_family: Some("Aptos".into()),
        font_size_centipoints: Some(1_400),
        text_color: Some(NativeOfficeRgbColor::new(0x12, 0x34, 0x56)),
        alignment: None,
    }
}

fn advanced_word_run_format() -> NativeOfficeTextFormat {
    NativeOfficeTextFormat {
        double_strikethrough: Some(true),
        text_case: Some(NativeOfficeTextCase::SmallCaps),
        highlight: Some(NativeOfficeHighlightColor::Yellow),
        language: Some("en-US".into()),
        ..rich_format()
    }
}

fn centered() -> NativeOfficeTextFormat {
    NativeOfficeTextFormat {
        alignment: Some(NativeOfficeHorizontalAlignment::Center),
        ..NativeOfficeTextFormat::default()
    }
}

#[test]
fn text_format_mutation_has_a_typed_stable_json_contract() {
    let mutation = NativeOfficeMutation::SetTextFormat {
        path: "/body/p[1]/r[1]".into(),
        format: NativeOfficeTextFormat {
            alignment: Some(NativeOfficeHorizontalAlignment::Justify),
            strikethrough: Some(true),
            ..advanced_word_run_format()
        },
    };
    assert_eq!(
        serde_json::to_value(&mutation).unwrap(),
        serde_json::json!({
            "operation": "set-text-format",
            "path": "/body/p[1]/r[1]",
            "format": {
                "bold": true,
                "italic": false,
                "underline": "double",
                "script": "superscript",
                "strikethrough": true,
                "doubleStrikethrough": true,
                "textCase": "small-caps",
                "highlight": "yellow",
                "language": "en-US",
                "fontFamily": "Aptos",
                "fontSizeCentipoints": 1400,
                "textColor": { "red": 18, "green": 52, "blue": 86 },
                "alignment": "justify"
            }
        })
    );
    assert!(
        serde_json::from_value::<NativeOfficeMutation>(serde_json::json!({
            "operation": "set-text-format",
            "path": "/body/p[1]/r[1]",
            "format": { "shadow": true }
        }))
        .is_err()
    );
}

#[test]
fn portable_highlight_palette_has_stable_word_and_rgb_mappings() {
    for (color, word, rgb) in [
        (NativeOfficeHighlightColor::None, "none", None),
        (NativeOfficeHighlightColor::Black, "black", Some("000000")),
        (NativeOfficeHighlightColor::Blue, "blue", Some("0000FF")),
        (NativeOfficeHighlightColor::Cyan, "cyan", Some("00FFFF")),
        (
            NativeOfficeHighlightColor::DarkBlue,
            "darkBlue",
            Some("000080"),
        ),
        (
            NativeOfficeHighlightColor::DarkCyan,
            "darkCyan",
            Some("008080"),
        ),
        (
            NativeOfficeHighlightColor::DarkGray,
            "darkGray",
            Some("808080"),
        ),
        (
            NativeOfficeHighlightColor::DarkGreen,
            "darkGreen",
            Some("008000"),
        ),
        (
            NativeOfficeHighlightColor::DarkMagenta,
            "darkMagenta",
            Some("800080"),
        ),
        (
            NativeOfficeHighlightColor::DarkRed,
            "darkRed",
            Some("800000"),
        ),
        (
            NativeOfficeHighlightColor::DarkYellow,
            "darkYellow",
            Some("808000"),
        ),
        (NativeOfficeHighlightColor::Green, "green", Some("00FF00")),
        (
            NativeOfficeHighlightColor::LightGray,
            "lightGray",
            Some("C0C0C0"),
        ),
        (
            NativeOfficeHighlightColor::Magenta,
            "magenta",
            Some("FF00FF"),
        ),
        (NativeOfficeHighlightColor::Red, "red", Some("FF0000")),
        (NativeOfficeHighlightColor::White, "white", Some("FFFFFF")),
        (NativeOfficeHighlightColor::Yellow, "yellow", Some("FFFF00")),
    ] {
        assert_eq!(color.word_value(), word);
        assert_eq!(color.rgb_hex(), rgb);
    }
}

#[tokio::test]
async fn native_word_writes_run_format_and_paragraph_alignment_losslessly() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("formatted.docx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor.set_text("/body/p[1]", "Native Word").unwrap();

    editor
        .set_text_format(
            "/body/p[1]/r[1]",
            NativeOfficeTextFormat {
                strikethrough: Some(true),
                ..advanced_word_run_format()
            },
        )
        .unwrap();
    editor.set_text_format("/body/p[1]", centered()).unwrap();

    let snapshot = editor.snapshot().unwrap();
    let paragraph = snapshot.get("/body/p[1]", 1).unwrap();
    let run = snapshot.get("/body/p[1]/r[1]", 0).unwrap();
    assert_eq!(paragraph.format["alignment"], "center");
    assert_eq!(run.format["bold"], "true");
    assert_eq!(run.format["italic"], "false");
    assert_eq!(run.format["underline"], "double");
    assert_eq!(run.format["script"], "superscript");
    assert_eq!(run.format["strike"], "true");
    assert_eq!(run.format["doubleStrike"], "true");
    assert_eq!(run.format["textCase"], "small-caps");
    assert_eq!(run.format["highlight"], "yellow");
    assert_eq!(run.format["language"], "en-US");
    assert_eq!(run.format["font"], "Aptos");
    assert_eq!(run.format["size"], "14pt");
    assert_eq!(run.format["color"], "123456");

    let xml =
        String::from_utf8(editor.package().part("word/document.xml").unwrap().to_vec()).unwrap();
    assert!(xml.contains("<w:sectPr>"));
    assert_eq!(xml.matches("<w:rPr").count(), 1);
    assert_eq!(xml.matches("<w:pPr").count(), 1);
    assert!(xml.contains("<w:strike w:val=\"1\"/>"));
    assert!(xml.contains("<w:dstrike w:val=\"1\"/>"));
    assert!(xml.contains("<w:caps w:val=\"0\"/>"));
    assert!(xml.contains("<w:smallCaps w:val=\"1\"/>"));
    assert!(xml.contains("<w:highlight w:val=\"yellow\"/>"));
    assert!(xml.contains("<w:lang w:val=\"en-US\"/>"));
    assert!(xml.contains("<w:u w:val=\"double\"/>"));
    assert!(xml.contains("<w:vertAlign w:val=\"superscript\"/>"));
}

#[tokio::test]
async fn word_explicit_format_overrides_theme_and_complex_script_properties() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("word-theme.docx");
    let mut seed = NativeOfficeEditor::create(&path).await.unwrap();
    seed.set_text("/body/p[1]", "Native Word").unwrap();

    let mut package = seed.package().clone();
    let document_part = "word/document.xml";
    let document = String::from_utf8(package.part(document_part).unwrap().to_vec())
        .unwrap()
        .replace(
            "<w:r>",
            concat!(
                "<w:r><w:rPr>",
                "<w:rFonts w:asciiTheme=\"minorHAnsi\" w:hAnsiTheme=\"minorHAnsi\" ",
                "w:eastAsiaTheme=\"minorEastAsia\" w:cstheme=\"minorBidi\" dataKeep=\"font\"/>",
                "<w:bCs w:val=\"1\"/><w:iCs w:val=\"1\"/>",
                "</w:rPr>"
            ),
        );
    package
        .set_part(document_part, document.into_bytes())
        .unwrap();

    let mut editor = NativeOfficeEditor::from_package(package).unwrap();
    editor
        .set_text_format(
            "/body/p[1]/r[1]",
            NativeOfficeTextFormat {
                bold: Some(false),
                italic: Some(false),
                font_family: Some("Courier New".into()),
                ..NativeOfficeTextFormat::default()
            },
        )
        .unwrap();

    let document =
        String::from_utf8(editor.package().part(document_part).unwrap().to_vec()).unwrap();
    for attribute in ["w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"] {
        assert!(document.contains(&format!("{attribute}=\"Courier New\"")));
    }
    for attribute in [
        "w:asciiTheme",
        "w:hAnsiTheme",
        "w:eastAsiaTheme",
        "w:cstheme",
    ] {
        assert!(!document.contains(attribute));
    }
    assert!(document.contains("dataKeep=\"font\""));
    assert!(document.contains("<w:bCs w:val=\"0\"/>"));
    assert!(document.contains("<w:iCs w:val=\"0\"/>"));
}

#[tokio::test]
async fn native_presentation_writes_run_format_and_paragraph_alignment() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("formatted.pptx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor.add_slide("/", "Native Slides").unwrap();

    editor
        .set_text_format(
            "/slide[1]/shape[1]/paragraph[1]/run[1]",
            NativeOfficeTextFormat {
                double_strikethrough: None,
                ..advanced_word_run_format()
            },
        )
        .unwrap();
    editor
        .set_text_format("/slide[1]/shape[1]/paragraph[1]", centered())
        .unwrap();

    let snapshot = editor.snapshot().unwrap();
    let paragraph = snapshot.get("/slide[1]/shape[1]/paragraph[1]", 1).unwrap();
    let run = snapshot
        .get("/slide[1]/shape[1]/paragraph[1]/run[1]", 0)
        .unwrap();
    assert_eq!(paragraph.format["alignment"], "ctr");
    assert_eq!(run.format["bold"], "1");
    assert_eq!(run.format["italic"], "0");
    assert_eq!(run.format["underline"], "double");
    assert_eq!(run.format["script"], "superscript");
    assert_eq!(run.format["textCase"], "small-caps");
    assert_eq!(run.format["highlight"], "yellow");
    assert_eq!(run.format["language"], "en-US");
    assert_eq!(run.format["font"], "Aptos");
    assert_eq!(run.format["size"], "14pt");
    assert_eq!(run.format["color"], "123456");

    let xml = String::from_utf8(
        editor
            .package()
            .part("ppt/slides/slide1.xml")
            .unwrap()
            .to_vec(),
    )
    .unwrap();
    assert!(xml.contains("u=\"dbl\""));
    assert!(xml.contains("baseline=\"30000\""));
    assert!(xml.contains("cap=\"small\""));
    assert!(xml.contains("lang=\"en-US\""));
    assert!(xml.contains("<a:highlight><a:srgbClr val=\"FFFF00\"/></a:highlight>"));
}

#[tokio::test]
async fn presentation_color_update_preserves_existing_attributes_and_transforms() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("preserve-color.pptx");
    let mut seed = NativeOfficeEditor::create(&path).await.unwrap();
    seed.add_slide("/", "Native Slides").unwrap();
    let run_path = "/slide[1]/shape[1]/paragraph[1]/run[1]";
    seed.set_text_format(
        run_path,
        NativeOfficeTextFormat {
            text_color: Some(NativeOfficeRgbColor::new(0, 0, 0)),
            ..NativeOfficeTextFormat::default()
        },
    )
    .unwrap();

    let mut package = seed.package().clone();
    let slide_part = "ppt/slides/slide1.xml";
    let slide = String::from_utf8(package.part(slide_part).unwrap().to_vec())
        .unwrap()
        .replace(
            "<a:srgbClr val=\"000000\"/>",
            "<a:srgbClr val=\"000000\" dataKeep=\"color\"><a:alpha val=\"50000\"/></a:srgbClr>",
        );
    package.set_part(slide_part, slide.into_bytes()).unwrap();

    let mut editor = NativeOfficeEditor::from_package(package).unwrap();
    editor
        .set_text_format(
            run_path,
            NativeOfficeTextFormat {
                text_color: Some(NativeOfficeRgbColor::new(0x12, 0x34, 0x56)),
                ..NativeOfficeTextFormat::default()
            },
        )
        .unwrap();

    let slide = String::from_utf8(editor.package().part(slide_part).unwrap().to_vec()).unwrap();
    let color_start = slide.find("<a:srgbClr").unwrap();
    let color_end = slide[color_start..].find("</a:srgbClr>").unwrap() + color_start;
    let color = &slide[color_start..color_end];
    assert!(color.contains("val=\"123456\""));
    assert!(color.contains("dataKeep=\"color\""));
    assert!(color.contains("<a:alpha val=\"50000\"/>"));
}

#[tokio::test]
async fn presentation_highlight_update_preserves_existing_attributes_and_transforms() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("preserve-highlight.pptx");
    let mut seed = NativeOfficeEditor::create(&path).await.unwrap();
    seed.add_slide("/", "Native Slides").unwrap();
    let run_path = "/slide[1]/shape[1]/paragraph[1]/run[1]";
    seed.set_text_format(
        run_path,
        NativeOfficeTextFormat {
            highlight: Some(NativeOfficeHighlightColor::Black),
            ..NativeOfficeTextFormat::default()
        },
    )
    .unwrap();

    let mut package = seed.package().clone();
    let slide_part = "ppt/slides/slide1.xml";
    let slide = String::from_utf8(package.part(slide_part).unwrap().to_vec())
        .unwrap()
        .replace(
            "<a:highlight><a:srgbClr val=\"000000\"/></a:highlight>",
            concat!(
                "<a:highlight dataKeep=\"highlight\">",
                "<a:srgbClr val=\"000000\" dataKeep=\"color\">",
                "<a:alpha val=\"50000\"/>",
                "</a:srgbClr></a:highlight>"
            ),
        );
    package.set_part(slide_part, slide.into_bytes()).unwrap();

    let mut editor = NativeOfficeEditor::from_package(package).unwrap();
    editor
        .set_text_format(
            run_path,
            NativeOfficeTextFormat {
                highlight: Some(NativeOfficeHighlightColor::Cyan),
                ..NativeOfficeTextFormat::default()
            },
        )
        .unwrap();

    let slide = String::from_utf8(editor.package().part(slide_part).unwrap().to_vec()).unwrap();
    assert!(slide.contains("<a:highlight dataKeep=\"highlight\">"));
    let highlight_start = slide.find("<a:highlight").unwrap();
    let highlight_end = slide[highlight_start..].find("</a:highlight>").unwrap() + highlight_start;
    let highlight = &slide[highlight_start..highlight_end];
    assert!(highlight.contains("val=\"00FFFF\""));
    assert!(highlight.contains("dataKeep=\"color\""));
    assert!(highlight.contains("<a:alpha val=\"50000\"/>"));
    let run = editor.snapshot().unwrap().get(run_path, 0).unwrap();
    assert_eq!(run.format["highlight"], "cyan");
}

#[tokio::test]
async fn native_spreadsheet_creates_and_deduplicates_cell_styles() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("formatted.xlsx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    assert!(!editor.package().contains_part("xl/styles.xml"));

    editor
        .set_text_format(
            "/Sheet1/A1:B2",
            NativeOfficeTextFormat {
                alignment: Some(NativeOfficeHorizontalAlignment::Center),
                strikethrough: Some(false),
                ..rich_format()
            },
        )
        .unwrap();

    for cell_path in ["/Sheet1/A1", "/Sheet1/B2"] {
        let cell = editor.snapshot().unwrap().get(cell_path, 0).unwrap();
        assert_eq!(cell.format["bold"], "true");
        assert_eq!(cell.format["italic"], "false");
        assert_eq!(cell.format["underline"], "double");
        assert_eq!(cell.format["script"], "superscript");
        assert_eq!(cell.format["strike"], "false");
        assert_eq!(cell.format["font"], "Aptos");
        assert_eq!(cell.format["size"], "14pt");
        assert_eq!(cell.format["color"], "123456");
        assert_eq!(cell.format["alignment"], "center");
    }
    assert!(editor.package().contains_part("xl/styles.xml"));
    let first_styles = editor.package().part("xl/styles.xml").unwrap().to_vec();
    let styles = String::from_utf8(first_styles.clone()).unwrap();
    assert_eq!(styles.matches("<scheme val=\"minor\"/>").count(), 1);

    editor
        .set_text_format(
            "/Sheet1/A1:B2",
            NativeOfficeTextFormat {
                alignment: Some(NativeOfficeHorizontalAlignment::Center),
                strikethrough: Some(false),
                ..rich_format()
            },
        )
        .unwrap();
    assert_eq!(
        editor.package().part("xl/styles.xml").unwrap(),
        first_styles
    );
}

#[tokio::test]
async fn native_formats_explicitly_clear_underline_and_vertical_script() {
    let temp = tempfile::tempdir().unwrap();

    let word_path = temp.path().join("clear.docx");
    let mut word = NativeOfficeEditor::create(&word_path).await.unwrap();
    word.set_text("/body/p[1]", "Word").unwrap();
    word.set_text_format(
        "/body/p[1]/r[1]",
        NativeOfficeTextFormat {
            underline: Some(NativeOfficeUnderline::None),
            script: Some(NativeOfficeTextScript::Baseline),
            strikethrough: Some(false),
            ..NativeOfficeTextFormat::default()
        },
    )
    .unwrap();
    let run = word.snapshot().unwrap().get("/body/p[1]/r[1]", 0).unwrap();
    assert_eq!(run.format["underline"], "none");
    assert_eq!(run.format["script"], "baseline");
    assert_eq!(run.format["strike"], "false");

    let spreadsheet_path = temp.path().join("clear.xlsx");
    let mut spreadsheet = NativeOfficeEditor::create(&spreadsheet_path).await.unwrap();
    spreadsheet
        .set_text_format(
            "/Sheet1/A1",
            NativeOfficeTextFormat {
                underline: Some(NativeOfficeUnderline::None),
                script: Some(NativeOfficeTextScript::Baseline),
                strikethrough: Some(false),
                ..NativeOfficeTextFormat::default()
            },
        )
        .unwrap();
    let cell = spreadsheet
        .snapshot()
        .unwrap()
        .get("/Sheet1/A1", 0)
        .unwrap();
    assert_eq!(cell.format["underline"], "none");
    assert_eq!(cell.format["script"], "baseline");
    assert_eq!(cell.format["strike"], "false");

    let presentation_path = temp.path().join("clear.pptx");
    let mut presentation = NativeOfficeEditor::create(&presentation_path)
        .await
        .unwrap();
    presentation.add_slide("/", "Presentation").unwrap();
    presentation
        .set_text_format(
            "/slide[1]/shape[1]/paragraph[1]/run[1]",
            NativeOfficeTextFormat {
                underline: Some(NativeOfficeUnderline::None),
                script: Some(NativeOfficeTextScript::Baseline),
                ..NativeOfficeTextFormat::default()
            },
        )
        .unwrap();
    let run = presentation
        .snapshot()
        .unwrap()
        .get("/slide[1]/shape[1]/paragraph[1]/run[1]", 0)
        .unwrap();
    assert_eq!(run.format["underline"], "none");
    assert_eq!(run.format["script"], "baseline");
}

#[tokio::test]
async fn native_word_and_presentation_clear_highlight_and_text_case() {
    let temp = tempfile::tempdir().unwrap();

    let word_path = temp.path().join("clear-highlight.docx");
    let mut word = NativeOfficeEditor::create(&word_path).await.unwrap();
    word.set_text("/body/p[1]", "Word").unwrap();
    word.set_text_format("/body/p[1]/r[1]", advanced_word_run_format())
        .unwrap();
    word.set_text_format(
        "/body/p[1]/r[1]",
        NativeOfficeTextFormat {
            text_case: Some(NativeOfficeTextCase::None),
            highlight: Some(NativeOfficeHighlightColor::None),
            language: Some("zh-Hant-CN".into()),
            ..NativeOfficeTextFormat::default()
        },
    )
    .unwrap();
    let run = word.snapshot().unwrap().get("/body/p[1]/r[1]", 0).unwrap();
    assert_eq!(run.format["textCase"], "none");
    assert_eq!(run.format["highlight"], "none");
    assert_eq!(run.format["language"], "zh-Hant-CN");

    let presentation_path = temp.path().join("clear-highlight.pptx");
    let mut presentation = NativeOfficeEditor::create(&presentation_path)
        .await
        .unwrap();
    presentation.add_slide("/", "Presentation").unwrap();
    let run_path = "/slide[1]/shape[1]/paragraph[1]/run[1]";
    presentation
        .set_text_format(
            run_path,
            NativeOfficeTextFormat {
                double_strikethrough: None,
                ..advanced_word_run_format()
            },
        )
        .unwrap();
    presentation
        .set_text_format(
            run_path,
            NativeOfficeTextFormat {
                text_case: Some(NativeOfficeTextCase::None),
                highlight: Some(NativeOfficeHighlightColor::None),
                language: Some("zh-Hant-CN".into()),
                ..NativeOfficeTextFormat::default()
            },
        )
        .unwrap();
    let run = presentation.snapshot().unwrap().get(run_path, 0).unwrap();
    assert_eq!(run.format["textCase"], "none");
    assert!(!run.format.contains_key("highlight"));
    assert_eq!(run.format["language"], "zh-Hant-CN");
    let slide = String::from_utf8(
        presentation
            .package()
            .part("ppt/slides/slide1.xml")
            .unwrap()
            .to_vec(),
    )
    .unwrap();
    assert!(!slide.contains("<a:highlight>"));
}

#[tokio::test]
async fn invalid_language_and_spreadsheet_run_only_formatting_are_atomic() {
    let temp = tempfile::tempdir().unwrap();

    let word_path = temp.path().join("invalid-language.docx");
    let mut word = NativeOfficeEditor::create(&word_path).await.unwrap();
    word.set_text("/body/p[1]", "Before").unwrap();
    let original = word.package().part("word/document.xml").unwrap().to_vec();
    let error = word
        .apply_batch(&[
            NativeOfficeMutation::SetText {
                path: "/body/p[1]/r[1]".into(),
                text: "After".into(),
            },
            NativeOfficeMutation::SetTextFormat {
                path: "/body/p[1]/r[1]".into(),
                format: NativeOfficeTextFormat {
                    language: Some("not_a_language".into()),
                    ..NativeOfficeTextFormat::default()
                },
            },
        ])
        .unwrap_err();
    assert_eq!(error.code, "use.office.language_invalid");
    assert_eq!(word.package().part("word/document.xml").unwrap(), original);

    let spreadsheet_path = temp.path().join("unsupported-run-format.xlsx");
    let mut spreadsheet = NativeOfficeEditor::create(&spreadsheet_path).await.unwrap();
    let original = spreadsheet
        .package()
        .part("xl/worksheets/sheet1.xml")
        .unwrap()
        .to_vec();
    let error = spreadsheet
        .apply_batch(&[
            NativeOfficeMutation::SetText {
                path: "/Sheet1/A1".into(),
                text: "After".into(),
            },
            NativeOfficeMutation::SetTextFormat {
                path: "/Sheet1/A1".into(),
                format: NativeOfficeTextFormat {
                    highlight: Some(NativeOfficeHighlightColor::Yellow),
                    ..NativeOfficeTextFormat::default()
                },
            },
        ])
        .unwrap_err();
    assert_eq!(error.code, "use.office.spreadsheet_run_format_unsupported");
    assert_eq!(
        spreadsheet
            .package()
            .part("xl/worksheets/sheet1.xml")
            .unwrap(),
        original
    );
}

#[tokio::test]
async fn native_spreadsheet_formatting_preserves_the_strict_ooxml_dialect() {
    const TRANSITIONAL_SPREADSHEET: &str =
        "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const STRICT_SPREADSHEET: &str = "http://purl.oclc.org/ooxml/spreadsheetml/main";
    const TRANSITIONAL_RELATIONSHIPS: &str =
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const STRICT_RELATIONSHIPS: &str = "http://purl.oclc.org/ooxml/officeDocument/relationships";

    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("strict.xlsx");
    let seed = NativeOfficeEditor::create(&path).await.unwrap();
    let mut package = seed.package().clone();
    for part_name in [
        "_rels/.rels",
        "xl/workbook.xml",
        "xl/_rels/workbook.xml.rels",
        "xl/worksheets/sheet1.xml",
    ] {
        let xml = String::from_utf8(package.part(part_name).unwrap().to_vec())
            .unwrap()
            .replace(TRANSITIONAL_SPREADSHEET, STRICT_SPREADSHEET)
            .replace(TRANSITIONAL_RELATIONSHIPS, STRICT_RELATIONSHIPS);
        package.set_part(part_name, xml.into_bytes()).unwrap();
    }
    let mut editor = NativeOfficeEditor::from_package(package).unwrap();
    editor.set_text_format("/Sheet1/A1", rich_format()).unwrap();
    editor
        .set_cell_format(
            "/Sheet1/A1",
            super::NativeSpreadsheetCellFormat {
                number_format: Some("scientific".into()),
                fill: Some(super::NativeSpreadsheetFill::Solid {
                    color: NativeOfficeRgbColor::new(0x12, 0x34, 0x56),
                }),
                border: Some(super::NativeSpreadsheetBorder {
                    top: Some(super::NativeSpreadsheetBorderLine::Line {
                        style: super::NativeSpreadsheetBorderStyle::Double,
                        color: Some(NativeOfficeRgbColor::new(0x65, 0x43, 0x21)),
                    }),
                    ..super::NativeSpreadsheetBorder::default()
                }),
                vertical_alignment: Some(super::NativeSpreadsheetVerticalAlignment::Distributed),
                ..super::NativeSpreadsheetCellFormat::default()
            },
        )
        .unwrap();

    let styles =
        String::from_utf8(editor.package().part("xl/styles.xml").unwrap().to_vec()).unwrap();
    let relationships = String::from_utf8(
        editor
            .package()
            .part("xl/_rels/workbook.xml.rels")
            .unwrap()
            .to_vec(),
    )
    .unwrap();
    assert!(styles.contains(STRICT_SPREADSHEET));
    assert!(!styles.contains(TRANSITIONAL_SPREADSHEET));
    assert!(styles.contains("rgb=\"FF123456\""));
    assert!(styles.contains("<top style=\"double\"><color rgb=\"FF654321\"/></top>"));
    assert!(styles.contains("vertical=\"distributed\""));
    assert!(relationships.contains(&format!("{STRICT_RELATIONSHIPS}/styles")));
}

#[tokio::test]
async fn spreadsheet_style_cloning_preserves_unknown_font_and_xf_data() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("preserve.xlsx");
    let mut seed = NativeOfficeEditor::create(&path).await.unwrap();
    seed.set_text_format("/Sheet1/A1", rich_format()).unwrap();

    let mut package = seed.package().clone();
    let mut styles = String::from_utf8(package.part("xl/styles.xml").unwrap().to_vec()).unwrap();
    let font_start = styles.rfind("<font>").unwrap() + "<font".len();
    styles.insert_str(font_start, " dataKeep=\"font\"");
    styles = styles.replace(
        "<u val=\"double\"/>",
        "<u val=\"double\" dataKeep=\"underline\"/>",
    );
    styles = styles.replace(
        "<vertAlign val=\"superscript\"/>",
        "<vertAlign val=\"superscript\" dataKeep=\"script\"/>",
    );
    let cell_xfs_end = styles.find("</cellXfs>").unwrap();
    let xf_start = styles[..cell_xfs_end].rfind("<xf ").unwrap();
    let xf_end = xf_start + styles[xf_start..].find('>').unwrap();
    let xf_attribute_position = if styles.as_bytes()[xf_end - 1] == b'/' {
        xf_end - 1
    } else {
        xf_end
    };
    styles.insert_str(xf_attribute_position, " quotePrefix=\"1\" dataKeep=\"xf\"");
    package
        .set_part("xl/styles.xml", styles.into_bytes())
        .unwrap();

    let mut editor = NativeOfficeEditor::from_package(package).unwrap();
    editor
        .set_text_format(
            "/Sheet1/A1",
            NativeOfficeTextFormat {
                italic: Some(true),
                underline: Some(NativeOfficeUnderline::Single),
                ..NativeOfficeTextFormat::default()
            },
        )
        .unwrap();
    let styles =
        String::from_utf8(editor.package().part("xl/styles.xml").unwrap().to_vec()).unwrap();
    assert_eq!(styles.matches("dataKeep=\"font\"").count(), 2);
    assert_eq!(styles.matches("dataKeep=\"underline\"").count(), 2);
    assert!(styles.split("<u ").skip(1).any(|fragment| {
        fragment.split("/>").next().is_some_and(|tag| {
            tag.contains("val=\"single\"") && tag.contains("dataKeep=\"underline\"")
        })
    }));
    assert_eq!(styles.matches("dataKeep=\"script\"").count(), 2);
    assert_eq!(styles.matches("val=\"superscript\"").count(), 2);
    assert_eq!(styles.matches("dataKeep=\"xf\"").count(), 2);
    assert_eq!(styles.matches("quotePrefix=\"1\"").count(), 2);
}

#[tokio::test]
async fn unsupported_presentation_strikethrough_rolls_back_an_entire_native_batch() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("atomic.pptx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor.add_slide("/", "Before").unwrap();
    let slide_part = "ppt/slides/slide1.xml";
    let original = editor.package().part(slide_part).unwrap().to_vec();
    let run_path = "/slide[1]/shape[1]/paragraph[1]/run[1]";

    let error = editor
        .apply_batch(&[
            NativeOfficeMutation::SetText {
                path: run_path.into(),
                text: "After".into(),
            },
            NativeOfficeMutation::SetTextFormat {
                path: run_path.into(),
                format: NativeOfficeTextFormat {
                    strikethrough: Some(true),
                    ..NativeOfficeTextFormat::default()
                },
            },
        ])
        .unwrap_err();
    assert_eq!(
        error.code,
        "use.office.presentation_strikethrough_unsupported"
    );
    assert_eq!(editor.package().part(slide_part).unwrap(), original);

    let error = editor
        .set_text_format(
            run_path,
            NativeOfficeTextFormat {
                double_strikethrough: Some(true),
                ..NativeOfficeTextFormat::default()
            },
        )
        .unwrap_err();
    assert_eq!(
        error.code,
        "use.office.presentation_double_strikethrough_unsupported"
    );
    assert_eq!(editor.package().part(slide_part).unwrap(), original);
}

#[tokio::test]
async fn invalid_format_rolls_back_an_entire_native_batch() {
    let temp = tempfile::tempdir().unwrap();
    let path = temp.path().join("atomic.docx");
    let mut editor = NativeOfficeEditor::create(&path).await.unwrap();
    editor.set_text("/body/p[1]", "Before").unwrap();
    let original = editor.package().part("word/document.xml").unwrap().to_vec();

    let error = editor
        .apply_batch(&[
            NativeOfficeMutation::SetText {
                path: "/body/p[1]".into(),
                text: "After".into(),
            },
            NativeOfficeMutation::SetTextFormat {
                path: "/body/p[1]/r[1]".into(),
                format: NativeOfficeTextFormat {
                    font_size_centipoints: Some(1_125),
                    ..NativeOfficeTextFormat::default()
                },
            },
        ])
        .unwrap_err();
    assert_eq!(error.code, "use.office.font_size_unsupported");
    assert_eq!(
        editor.package().part("word/document.xml").unwrap(),
        original
    );
}
