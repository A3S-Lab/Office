use super::*;

#[test]
fn wraps_cjk_at_unicode_break_opportunities() {
    let text = "甲乙丙丁";
    let ranges = break_line_ranges(text, 2.0, 2.0, TextWhiteSpace::Normal, |start, end| {
        text[start..end].graphemes(true).count() as f64
    });

    assert_eq!(
        ranges,
        vec![
            LineRange {
                start: 0,
                visible_end: 6,
                next_start: 6,
                hard_break: false,
            },
            LineRange {
                start: 6,
                visible_end: 12,
                next_start: 12,
                hard_break: false,
            },
        ]
    );
}

#[test]
fn uses_grapheme_boundaries_for_emergency_wraps() {
    let text = "ab😀cd";
    let ranges = break_line_ranges(text, 3.0, 3.0, TextWhiteSpace::Normal, |start, end| {
        text[start..end].graphemes(true).count() as f64
    });
    let offsets = Utf16Offsets::new(text);

    assert_eq!(offsets.at(ranges[0].visible_end).expect("UTF-16"), 4);
    assert_eq!(&text[ranges[0].start..ranges[0].visible_end], "ab😀");
    assert_eq!(&text[ranges[1].start..ranges[1].visible_end], "cd");
}

#[test]
fn preserves_spaces_at_breaks_in_break_spaces_mode() {
    let text = "ab  cd";
    let ranges = break_line_ranges(text, 3.0, 3.0, TextWhiteSpace::BreakSpaces, |start, end| {
        text[start..end].chars().count() as f64
    });

    assert_eq!(&text[ranges[0].start..ranges[0].visible_end], "ab ");
    assert_eq!(&text[ranges[1].start..ranges[1].visible_end], " cd");
}

#[test]
fn preserves_explicit_line_breaks() {
    let ranges = break_line_ranges(
        "first\nsecond",
        80.0,
        80.0,
        TextWhiteSpace::Normal,
        |start, end| (end - start) as f64,
    );

    assert_eq!(ranges.len(), 2);
    assert_eq!(
        &"first\nsecond"[ranges[0].start..ranges[0].visible_end],
        "first"
    );
    assert!(ranges[0].hard_break);
    assert_eq!(
        &"first\nsecond"[ranges[1].start..ranges[1].visible_end],
        "second"
    );
}

#[test]
fn resolves_mixed_direction_text_into_line_scoped_bidi_runs() {
    let text = "A3S שלום 2026";
    let bidi = ResolvedBidiText::new(text, TextDirection::Ltr);
    let runs = bidi.directional_runs(0, text.len()).collect::<Vec<_>>();

    assert!(runs.iter().any(|run| run.direction == TextDirection::Ltr));
    assert!(runs.iter().any(|run| run.direction == TextDirection::Rtl));

    let mut logical_runs = runs;
    logical_runs.sort_by_key(|run| run.start);
    assert_eq!(logical_runs.first().map(|run| run.start), Some(0));
    assert_eq!(logical_runs.last().map(|run| run.end), Some(text.len()));
    assert!(logical_runs
        .windows(2)
        .all(|pair| pair[0].end == pair[1].start));
    assert!(logical_runs
        .iter()
        .all(|run| { text.is_char_boundary(run.start) && text.is_char_boundary(run.end) }));
}

#[test]
fn rejects_structured_tabs_when_the_resolved_line_contains_rtl_text() {
    let ltr = ResolvedBidiText::new("A3S\tOffice", TextDirection::Ltr);
    let mixed = ResolvedBidiText::new("A3S\tשלום", TextDirection::Ltr);

    assert!(text_tabs_support_bidi(&ltr));
    assert!(!text_tabs_support_bidi(&mixed));
}

#[test]
fn rejects_duplicate_paragraph_ids() {
    let paragraph = TextLayoutParagraph {
        id: "same".into(),
        text: "A3S".into(),
        runs: vec![TextLayoutRun {
            start_utf16: 0,
            end_utf16: 3,
            font_id: "font".into(),
            fallback_font_ids: Vec::new(),
            font_size: 14.0,
            line_height: 21.0,
            letter_spacing: 0.0,
            ligatures: true,
            kerning: true,
        }],
        max_width: 100.0,
        first_line_max_width: None,
        direction: TextDirection::Auto,
        white_space: TextWhiteSpace::Normal,
        tab_layout: None,
    };
    let request = TextLayoutRequest {
        protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
        kind: "textLayout".into(),
        request_id: 1,
        revision: 1,
        document_revision: 1,
        paragraphs: vec![paragraph.clone(), paragraph],
    };

    assert_eq!(
        validate_text_layout_request(&request)
            .expect_err("duplicate ID")
            .code,
        "office.kernel.text_paragraph_id_duplicate"
    );
}

#[test]
fn rejects_gaps_between_text_runs() {
    let request = TextLayoutRequest {
        protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
        kind: "textLayout".into(),
        request_id: 1,
        revision: 1,
        document_revision: 1,
        paragraphs: vec![TextLayoutParagraph {
            id: "paragraph".into(),
            text: "A3S 文档".into(),
            runs: vec![
                TextLayoutRun {
                    start_utf16: 0,
                    end_utf16: 3,
                    font_id: "regular".into(),
                    fallback_font_ids: Vec::new(),
                    font_size: 14.0,
                    line_height: 21.0,
                    letter_spacing: 0.0,
                    ligatures: true,
                    kerning: true,
                },
                TextLayoutRun {
                    start_utf16: 4,
                    end_utf16: 6,
                    font_id: "regular".into(),
                    fallback_font_ids: Vec::new(),
                    font_size: 18.0,
                    line_height: 27.0,
                    letter_spacing: 0.0,
                    ligatures: true,
                    kerning: true,
                },
            ],
            max_width: 100.0,
            first_line_max_width: None,
            direction: TextDirection::Auto,
            white_space: TextWhiteSpace::Normal,
            tab_layout: None,
        }],
    };

    assert_eq!(
        validate_text_layout_request(&request)
            .expect_err("run gap")
            .code,
        "office.kernel.text_run_range_invalid"
    );
}

#[test]
fn rejects_utf16_offsets_inside_a_surrogate_pair() {
    let request = TextLayoutRequest {
        protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
        kind: "textLayout".into(),
        request_id: 1,
        revision: 1,
        document_revision: 1,
        paragraphs: vec![TextLayoutParagraph {
            id: "paragraph".into(),
            text: "A😀B".into(),
            runs: vec![
                TextLayoutRun {
                    start_utf16: 0,
                    end_utf16: 2,
                    font_id: "regular".into(),
                    fallback_font_ids: Vec::new(),
                    font_size: 14.0,
                    line_height: 21.0,
                    letter_spacing: 0.0,
                    ligatures: true,
                    kerning: true,
                },
                TextLayoutRun {
                    start_utf16: 2,
                    end_utf16: 4,
                    font_id: "regular".into(),
                    fallback_font_ids: Vec::new(),
                    font_size: 14.0,
                    line_height: 21.0,
                    letter_spacing: 0.0,
                    ligatures: true,
                    kerning: true,
                },
            ],
            max_width: 100.0,
            first_line_max_width: None,
            direction: TextDirection::Auto,
            white_space: TextWhiteSpace::Normal,
            tab_layout: None,
        }],
    };

    assert_eq!(
        validate_text_layout_request(&request)
            .expect_err("surrogate split")
            .code,
        "office.kernel.text_offset_invalid"
    );
}

#[test]
fn rejects_duplicate_primary_and_fallback_font_ids() {
    let paragraph = TextLayoutParagraph {
        id: "fallback".into(),
        text: "A3S".into(),
        runs: vec![TextLayoutRun {
            start_utf16: 0,
            end_utf16: 3,
            font_id: "regular".into(),
            fallback_font_ids: vec!["cjk".into(), "regular".into()],
            font_size: 14.0,
            line_height: 21.0,
            letter_spacing: 0.0,
            ligatures: true,
            kerning: true,
        }],
        max_width: 100.0,
        first_line_max_width: None,
        direction: TextDirection::Ltr,
        white_space: TextWhiteSpace::Normal,
        tab_layout: None,
    };

    assert_eq!(
        validate_text_layout_runs(&paragraph)
            .expect_err("duplicate fallback font")
            .code,
        "office.kernel.text_fallback_font_invalid"
    );
}

#[test]
fn chooses_custom_tab_stops_without_skipping_earlier_default_stops() {
    let layout = TextTabLayout {
        origin: 0.0,
        first_line_indent: 0.0,
        default_interval: 48.0,
        stops: vec![
            TextTabStop {
                position: 96.0,
                alignment: TextTabAlignment::Right,
            },
            TextTabStop {
                position: 144.0,
                alignment: TextTabAlignment::Decimal,
            },
        ],
    };

    assert_eq!(
        next_text_tab_stop(&layout, 20.0),
        EffectiveTextTabStop {
            position: 48.0,
            alignment: TextTabAlignment::Left,
        }
    );
    assert_eq!(
        next_text_tab_stop(&layout, 60.0),
        EffectiveTextTabStop {
            position: 96.0,
            alignment: TextTabAlignment::Right,
        }
    );
    assert_eq!(
        next_text_tab_stop(&layout, 100.0),
        EffectiveTextTabStop {
            position: 144.0,
            alignment: TextTabAlignment::Decimal,
        }
    );
}

#[test]
fn aligns_tab_segments_by_left_center_right_and_decimal_metrics() {
    let segment = TextTabSegment {
        width: 80.0,
        decimal_width: Some(30.0),
    };

    assert_eq!(
        text_tab_alignment_offset(TextTabAlignment::Left, segment),
        0.0
    );
    assert_eq!(
        text_tab_alignment_offset(TextTabAlignment::Center, segment),
        40.0
    );
    assert_eq!(
        text_tab_alignment_offset(TextTabAlignment::Right, segment),
        80.0
    );
    assert_eq!(
        text_tab_alignment_offset(TextTabAlignment::Decimal, segment),
        30.0
    );
    assert_eq!(
        text_tab_alignment_offset(
            TextTabAlignment::Decimal,
            TextTabSegment {
                width: 80.0,
                decimal_width: None,
            },
        ),
        80.0
    );
}

#[test]
fn requires_layout_metadata_for_structured_tabs() {
    let paragraph = TextLayoutParagraph {
        id: "tabs".into(),
        text: "A\tB".into(),
        runs: vec![TextLayoutRun {
            start_utf16: 0,
            end_utf16: 3,
            font_id: "font".into(),
            fallback_font_ids: Vec::new(),
            font_size: 14.0,
            line_height: 21.0,
            letter_spacing: 0.0,
            ligatures: true,
            kerning: true,
        }],
        max_width: 100.0,
        first_line_max_width: None,
        direction: TextDirection::Ltr,
        white_space: TextWhiteSpace::Normal,
        tab_layout: None,
    };

    assert_eq!(
        validate_text_tab_layout(&paragraph)
            .expect_err("missing tab layout")
            .code,
        "office.kernel.text_tab_layout_missing"
    );
}
