use super::*;

pub(super) fn validate_text_layout_request(request: &TextLayoutRequest) -> Result<(), KernelError> {
    if request.protocol != OFFICE_KERNEL_PROTOCOL_VERSION {
        return Err(KernelError::invalid(
            "office.kernel.protocol_unsupported",
            format!(
                "Office kernel protocol {} is unsupported; expected {}.",
                request.protocol, OFFICE_KERNEL_PROTOCOL_VERSION
            ),
        ));
    }
    if request.kind != "textLayout" {
        return Err(KernelError::invalid(
            "office.kernel.request_kind_invalid",
            "The Office text kernel only accepts textLayout requests.",
        ));
    }
    if request.paragraphs.len() > MAX_TEXT_LAYOUT_PARAGRAPHS {
        return Err(KernelError::invalid(
            "office.kernel.text_paragraph_limit_exceeded",
            format!(
                "A text layout request may contain at most {MAX_TEXT_LAYOUT_PARAGRAPHS} paragraphs."
            ),
        ));
    }
    let total_bytes = request
        .paragraphs
        .iter()
        .try_fold(0_usize, |total, paragraph| {
            total.checked_add(paragraph.text.len())
        })
        .ok_or_else(|| {
            KernelError::invalid(
                "office.kernel.text_limit_exceeded",
                "The text layout request is too large.",
            )
        })?;
    if total_bytes > MAX_TEXT_LAYOUT_BYTES {
        return Err(KernelError::invalid(
            "office.kernel.text_limit_exceeded",
            format!(
                "A text layout request may contain at most {MAX_TEXT_LAYOUT_BYTES} UTF-8 bytes."
            ),
        ));
    }

    let mut ids = HashSet::with_capacity(request.paragraphs.len());
    let mut total_runs = 0_usize;
    for paragraph in &request.paragraphs {
        validate_identifier("paragraph", &paragraph.id, MAX_PARAGRAPH_ID_BYTES)?;
        if !ids.insert(paragraph.id.as_str()) {
            return Err(KernelError::invalid(
                "office.kernel.text_paragraph_id_duplicate",
                format!("Text layout paragraph ID '{}' is duplicated.", paragraph.id),
            ));
        }
        validate_positive_extent("maxWidth", paragraph.max_width, MAX_LINE_WIDTH)?;
        if let Some(first_line_max_width) = paragraph.first_line_max_width {
            validate_positive_extent("firstLineMaxWidth", first_line_max_width, MAX_LINE_WIDTH)?;
        }
        if paragraph.runs.is_empty() || paragraph.runs.len() > MAX_TEXT_LAYOUT_RUNS_PER_PARAGRAPH {
            return Err(KernelError::invalid(
                "office.kernel.text_run_limit_exceeded",
                format!(
                    "A paragraph requires between 1 and {MAX_TEXT_LAYOUT_RUNS_PER_PARAGRAPH} text runs."
                ),
            ));
        }
        total_runs = total_runs
            .checked_add(paragraph.runs.len())
            .ok_or_else(|| {
                KernelError::invalid(
                    "office.kernel.text_run_limit_exceeded",
                    "The text layout request contains too many text runs.",
                )
            })?;
        if total_runs > MAX_TEXT_LAYOUT_RUNS {
            return Err(KernelError::invalid(
                "office.kernel.text_run_limit_exceeded",
                format!(
                    "A text layout request may contain at most {MAX_TEXT_LAYOUT_RUNS} text runs."
                ),
            ));
        }
        validate_text_layout_runs(paragraph)?;
        validate_text_tab_layout(paragraph)?;
    }
    Ok(())
}

pub(super) fn validate_text_layout_runs(
    paragraph: &TextLayoutParagraph,
) -> Result<(), KernelError> {
    let offsets = Utf16Offsets::new(&paragraph.text);
    let text_len = offsets.len();
    let mut previous_end = 0_u32;
    for (index, run) in paragraph.runs.iter().enumerate() {
        validate_identifier("font", &run.font_id, MAX_FONT_ID_BYTES)?;
        if run.fallback_font_ids.len() + 1 > MAX_TEXT_LAYOUT_FONTS_PER_RUN {
            return Err(KernelError::invalid(
                "office.kernel.text_fallback_font_limit_exceeded",
                format!(
                    "Text layout run {index} in paragraph '{}' may use at most {MAX_TEXT_LAYOUT_FONTS_PER_RUN} font faces.",
                    paragraph.id
                ),
            ));
        }
        let mut font_ids = HashSet::with_capacity(run.fallback_font_ids.len() + 1);
        font_ids.insert(run.font_id.as_str());
        for font_id in &run.fallback_font_ids {
            validate_identifier("font", font_id, MAX_FONT_ID_BYTES)?;
            if !font_ids.insert(font_id.as_str()) {
                return Err(KernelError::invalid(
                    "office.kernel.text_fallback_font_invalid",
                    format!(
                        "Text layout run {index} in paragraph '{}' contains a duplicate font ID.",
                        paragraph.id
                    ),
                ));
            }
        }
        validate_positive_extent("fontSize", run.font_size, MAX_FONT_SIZE)?;
        validate_positive_extent("lineHeight", run.line_height, MAX_LINE_HEIGHT)?;
        if !run.letter_spacing.is_finite()
            || !(-MAX_LETTER_SPACING..=MAX_LETTER_SPACING).contains(&run.letter_spacing)
        {
            return Err(KernelError::invalid(
                "office.kernel.letter_spacing_invalid",
                format!(
                    "letterSpacing must be between -{MAX_LETTER_SPACING} and {MAX_LETTER_SPACING}."
                ),
            ));
        }
        let empty_text_run = text_len == 0 && run.start_utf16 == 0 && run.end_utf16 == 0;
        if run.start_utf16 != previous_end
            || run.end_utf16 > text_len
            || (run.end_utf16 <= run.start_utf16 && !empty_text_run)
        {
            return Err(KernelError::invalid(
                "office.kernel.text_run_range_invalid",
                format!(
                    "Text layout run {index} in paragraph '{}' must cover the next contiguous UTF-16 range.",
                    paragraph.id
                ),
            ));
        }
        offsets.byte_at(run.start_utf16)?;
        offsets.byte_at(run.end_utf16)?;
        previous_end = run.end_utf16;
    }
    if previous_end != text_len {
        return Err(KernelError::invalid(
            "office.kernel.text_run_range_invalid",
            format!(
                "Text layout runs in paragraph '{}' must cover the complete paragraph.",
                paragraph.id
            ),
        ));
    }
    Ok(())
}

pub(super) fn validate_text_tab_layout(paragraph: &TextLayoutParagraph) -> Result<(), KernelError> {
    let tab_count = paragraph.text.matches('\t').count();
    if tab_count > MAX_TEXT_LAYOUT_TABS_PER_PARAGRAPH {
        return Err(KernelError::invalid(
            "office.kernel.text_tab_limit_exceeded",
            format!(
                "A paragraph may contain at most {MAX_TEXT_LAYOUT_TABS_PER_PARAGRAPH} structured tabs."
            ),
        ));
    }
    if tab_count > 0 && paragraph.tab_layout.is_none() {
        return Err(KernelError::invalid(
            "office.kernel.text_tab_layout_missing",
            format!(
                "Paragraph '{}' contains structured tabs but has no tab layout.",
                paragraph.id
            ),
        ));
    }
    let Some(layout) = paragraph.tab_layout.as_ref() else {
        return Ok(());
    };
    if !layout.origin.is_finite() || !(0.0..=MAX_LINE_WIDTH).contains(&layout.origin) {
        return Err(KernelError::invalid(
            "office.kernel.text_tab_origin_invalid",
            format!("Tab origin must be between 0 and {MAX_LINE_WIDTH}."),
        ));
    }
    if !layout.first_line_indent.is_finite()
        || !(-MAX_LINE_WIDTH..=MAX_LINE_WIDTH).contains(&layout.first_line_indent)
    {
        return Err(KernelError::invalid(
            "office.kernel.text_tab_indent_invalid",
            format!(
                "Tab first-line indent must be between -{MAX_LINE_WIDTH} and {MAX_LINE_WIDTH}."
            ),
        ));
    }
    validate_positive_extent(
        "tabLayout.defaultInterval",
        layout.default_interval,
        MAX_LINE_WIDTH,
    )?;
    if layout.stops.len() > MAX_TEXT_LAYOUT_TAB_STOPS {
        return Err(KernelError::invalid(
            "office.kernel.text_tab_stop_limit_exceeded",
            format!("A paragraph may define at most {MAX_TEXT_LAYOUT_TAB_STOPS} tab stops."),
        ));
    }
    let mut previous_position = 0.0;
    for (index, stop) in layout.stops.iter().enumerate() {
        if !stop.position.is_finite()
            || !(0.0 < stop.position && stop.position <= MAX_LINE_WIDTH)
            || stop.position <= previous_position
        {
            return Err(KernelError::invalid(
                "office.kernel.text_tab_stop_invalid",
                format!(
                    "Tab stop {index} in paragraph '{}' must have a unique ascending position between 0 and {MAX_LINE_WIDTH}.",
                    paragraph.id
                ),
            ));
        }
        previous_position = stop.position;
    }
    Ok(())
}

pub(super) fn validate_identifier(
    kind: &str,
    id: &str,
    max_bytes: usize,
) -> Result<(), KernelError> {
    if id.trim().is_empty() || id.len() > max_bytes {
        return Err(KernelError::invalid(
            format!("office.kernel.{kind}_id_invalid"),
            format!("Every {kind} requires a non-empty ID of at most {max_bytes} bytes."),
        ));
    }
    Ok(())
}

fn validate_positive_extent(name: &str, value: f64, maximum: f64) -> Result<(), KernelError> {
    if !value.is_finite() || !(0.0 < value && value <= maximum) {
        return Err(KernelError::invalid(
            "office.kernel.text_extent_invalid",
            format!("{name} must be a finite positive number no greater than {maximum}."),
        ));
    }
    Ok(())
}
