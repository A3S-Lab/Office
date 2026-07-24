use std::collections::{HashMap, HashSet};

use rustybuzz::ttf_parser::Tag;
use rustybuzz::{shape, Direction, Face, Feature, UnicodeBuffer};
use serde::{Deserialize, Serialize};
use unicode_bidi::{BidiInfo, Level};
use unicode_linebreak::{linebreaks, BreakOpportunity};
use unicode_segmentation::UnicodeSegmentation;

use crate::{KernelError, OFFICE_KERNEL_PROTOCOL_VERSION};

mod utf16_offsets;
mod validation;

use utf16_offsets::Utf16Offsets;
use validation::{validate_identifier, validate_text_layout_request};
#[cfg(test)]
use validation::{validate_text_layout_runs, validate_text_tab_layout};

const MAX_FONT_ID_BYTES: usize = 128;
const MAX_PARAGRAPH_ID_BYTES: usize = 256;
const MAX_TEXT_LAYOUT_PARAGRAPHS: usize = 1_024;
const MAX_TEXT_LAYOUT_RUNS: usize = 16_384;
const MAX_TEXT_LAYOUT_RUNS_PER_PARAGRAPH: usize = 4_096;
const MAX_TEXT_LAYOUT_FONTS_PER_RUN: usize = 8;
const MAX_TEXT_LAYOUT_TABS_PER_PARAGRAPH: usize = 4_096;
const MAX_TEXT_LAYOUT_TAB_STOPS: usize = 64;
const MAX_TEXT_LAYOUT_BYTES: usize = 1_048_576;
const MAX_FONT_SIZE: f64 = 512.0;
const MAX_LINE_HEIGHT: f64 = 2_048.0;
const MAX_LINE_WIDTH: f64 = 1_000_000.0;
const MAX_LETTER_SPACING: f64 = 100.0;
const MIN_TAB_ADVANCE: f64 = 1.0;

pub type FontRegistry = HashMap<String, Vec<u8>>;
type ResolvedFontRegistry<'font> = HashMap<&'font str, Face<'font>>;

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum TextDirection {
    #[default]
    Auto,
    Ltr,
    Rtl,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum TextWhiteSpace {
    #[default]
    Normal,
    BreakSpaces,
}

#[derive(Debug, Clone, Copy, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum TextTabAlignment {
    Center,
    Decimal,
    #[default]
    Left,
    Right,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextTabStop {
    pub position: f64,
    pub alignment: TextTabAlignment,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextTabLayout {
    pub origin: f64,
    pub first_line_indent: f64,
    pub default_interval: f64,
    pub stops: Vec<TextTabStop>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextLayoutRun {
    pub start_utf16: u32,
    pub end_utf16: u32,
    pub font_id: String,
    #[serde(default)]
    pub fallback_font_ids: Vec<String>,
    pub font_size: f64,
    pub line_height: f64,
    #[serde(default)]
    pub letter_spacing: f64,
    #[serde(default = "default_ligatures")]
    pub ligatures: bool,
    #[serde(default = "default_kerning")]
    pub kerning: bool,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextLayoutParagraph {
    pub id: String,
    pub text: String,
    pub runs: Vec<TextLayoutRun>,
    pub max_width: f64,
    #[serde(default)]
    pub first_line_max_width: Option<f64>,
    #[serde(default)]
    pub direction: TextDirection,
    #[serde(default)]
    pub white_space: TextWhiteSpace,
    #[serde(default)]
    pub tab_layout: Option<TextTabLayout>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextLayoutRequest {
    pub protocol: u32,
    pub kind: String,
    pub request_id: u32,
    pub revision: u32,
    pub document_revision: u64,
    pub paragraphs: Vec<TextLayoutParagraph>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextLayoutLine {
    pub start_utf16: u32,
    pub end_utf16: u32,
    pub width: f64,
    pub ascent: f64,
    pub descent: f64,
    pub height: f64,
    pub hard_break: bool,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextLayoutParagraphResult {
    pub id: String,
    pub glyph_count: u32,
    pub fallback_glyph_count: u32,
    pub missing_glyph_count: u32,
    pub lines: Vec<TextLayoutLine>,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextLayoutResult {
    pub protocol: u32,
    pub kind: String,
    pub request_id: u32,
    pub revision: u32,
    pub document_revision: u64,
    pub engine: String,
    pub layouts: Vec<TextLayoutParagraphResult>,
    pub unsupported_paragraph_ids: Vec<String>,
}

pub fn validate_font(id: &str, data: &[u8]) -> Result<(), KernelError> {
    validate_identifier("font", id, MAX_FONT_ID_BYTES)?;
    Face::from_slice(data, 0).ok_or_else(|| {
        KernelError::invalid(
            "office.kernel.font_invalid",
            format!("Font '{id}' is not a supported OpenType font."),
        )
    })?;
    Ok(())
}

pub fn layout_text(
    request: &TextLayoutRequest,
    fonts: &FontRegistry,
) -> Result<TextLayoutResult, KernelError> {
    validate_text_layout_request(request)?;
    let resolved_fonts = resolve_font_registry(fonts)?;
    let mut layouts = Vec::with_capacity(request.paragraphs.len());
    let mut unsupported_paragraph_ids = Vec::new();

    for paragraph in &request.paragraphs {
        match layout_paragraph(paragraph, &resolved_fonts)? {
            Some(layout) => layouts.push(layout),
            None => unsupported_paragraph_ids.push(paragraph.id.clone()),
        }
    }

    Ok(TextLayoutResult {
        protocol: OFFICE_KERNEL_PROTOCOL_VERSION,
        kind: "textLayoutResult".into(),
        request_id: request.request_id,
        revision: request.revision,
        document_revision: request.document_revision,
        engine: "wasm".into(),
        layouts,
        unsupported_paragraph_ids,
    })
}

fn resolve_font_registry(fonts: &FontRegistry) -> Result<ResolvedFontRegistry<'_>, KernelError> {
    let mut resolved = HashMap::with_capacity(fonts.len());
    for (id, data) in fonts {
        let face = Face::from_slice(data, 0).ok_or_else(|| {
            KernelError::invalid(
                "office.kernel.font_invalid",
                format!("Font '{id}' is not a supported OpenType font."),
            )
        })?;
        if face.units_per_em() == 0 {
            return Err(KernelError::invalid(
                "office.kernel.font_metrics_invalid",
                format!("Font '{id}' has invalid units-per-em metrics."),
            ));
        }
        resolved.insert(id.as_str(), face);
    }
    Ok(resolved)
}

fn layout_paragraph(
    paragraph: &TextLayoutParagraph,
    fonts: &ResolvedFontRegistry<'_>,
) -> Result<Option<TextLayoutParagraphResult>, KernelError> {
    let bidi = ResolvedBidiText::new(&paragraph.text, paragraph.direction);
    if paragraph.tab_layout.is_some() && !text_tabs_support_bidi(&bidi) {
        return Ok(None);
    }
    let utf16 = Utf16Offsets::new(&paragraph.text);
    let mut runs = Vec::with_capacity(paragraph.runs.len());
    for run in &paragraph.runs {
        let start = utf16.byte_at(run.start_utf16)?;
        let end = utf16.byte_at(run.end_utf16)?;
        let mut faces = Vec::with_capacity(run.fallback_font_ids.len() + 1);
        for font_id in std::iter::once(&run.font_id).chain(&run.fallback_font_ids) {
            let Some(face) = fonts.get(font_id.as_str()) else {
                return Ok(None);
            };
            faces.push(face.clone());
        }
        let font_segments = resolve_font_segments(&paragraph.text, start, end, &faces);
        runs.push(ResolvedTextLayoutRun {
            start,
            end,
            run,
            faces,
            font_segments,
        });
    }
    let resolved_tabs = paragraph
        .tab_layout
        .as_ref()
        .map(|layout| resolve_text_tabs(&paragraph.text, &runs, &bidi, layout));
    let ranges = break_line_ranges(
        &paragraph.text,
        paragraph.max_width,
        paragraph
            .first_line_max_width
            .unwrap_or(paragraph.max_width),
        paragraph.white_space,
        |start, end| {
            shape_text_range(
                &paragraph.text,
                start,
                end,
                &runs,
                &bidi,
                resolved_tabs.as_ref(),
                start == 0,
            )
            .width
        },
    );
    let mut glyph_count = 0_u32;
    let mut fallback_glyph_count = 0_u32;
    let mut missing_glyph_count = 0_u32;
    let mut lines = Vec::with_capacity(ranges.len());

    for range in ranges {
        let shaped = shape_text_range(
            &paragraph.text,
            range.start,
            range.visible_end,
            &runs,
            &bidi,
            resolved_tabs.as_ref(),
            range.start == 0,
        );
        let metrics = text_line_metrics(range.start, range.next_start, &runs);
        glyph_count = glyph_count.saturating_add(shaped.glyph_count);
        fallback_glyph_count = fallback_glyph_count.saturating_add(shaped.fallback_glyph_count);
        missing_glyph_count = missing_glyph_count.saturating_add(shaped.missing_glyph_count);
        lines.push(TextLayoutLine {
            start_utf16: utf16.at(range.start)?,
            end_utf16: utf16.at(range.visible_end)?,
            width: shaped.width,
            ascent: metrics.ascent,
            descent: metrics.descent,
            height: metrics.height,
            hard_break: range.hard_break,
        });
    }

    Ok(Some(TextLayoutParagraphResult {
        id: paragraph.id.clone(),
        glyph_count,
        fallback_glyph_count,
        missing_glyph_count,
        lines,
    }))
}

struct ResolvedTextLayoutRun<'font, 'run> {
    start: usize,
    end: usize,
    run: &'run TextLayoutRun,
    faces: Vec<Face<'font>>,
    font_segments: Vec<ResolvedFontSegment>,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct ResolvedFontSegment {
    start: usize,
    end: usize,
    face_index: usize,
}

struct ResolvedBidiText {
    runs: Vec<ResolvedBidiRun>,
    has_rtl: bool,
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct ResolvedBidiRun {
    start: usize,
    end: usize,
    direction: TextDirection,
}

impl ResolvedBidiText {
    fn new(text: &str, direction: TextDirection) -> Self {
        let base_level = match direction {
            TextDirection::Auto => None,
            TextDirection::Ltr => Some(Level::ltr()),
            TextDirection::Rtl => Some(Level::rtl()),
        };
        let info = BidiInfo::new(text, base_level);
        let mut runs: Vec<ResolvedBidiRun> = Vec::new();

        for (start, character) in text.char_indices() {
            let direction = if info.levels[start].is_rtl() {
                TextDirection::Rtl
            } else {
                TextDirection::Ltr
            };
            let end = start + character.len_utf8();
            if let Some(previous) = runs.last_mut() {
                if previous.direction == direction {
                    previous.end = end;
                    continue;
                }
            }
            runs.push(ResolvedBidiRun {
                start,
                end,
                direction,
            });
        }

        Self {
            has_rtl: runs.iter().any(|run| run.direction == TextDirection::Rtl),
            runs,
        }
    }

    fn directional_runs(
        &self,
        start: usize,
        end: usize,
    ) -> impl Iterator<Item = ResolvedBidiRun> + '_ {
        self.runs.iter().filter_map(move |run| {
            let run_start = start.max(run.start);
            let run_end = end.min(run.end);
            (run_start < run_end).then_some(ResolvedBidiRun {
                start: run_start,
                end: run_end,
                direction: run.direction,
            })
        })
    }
}

fn text_tabs_support_bidi(bidi: &ResolvedBidiText) -> bool {
    !bidi.has_rtl
}

fn resolve_font_segments(
    text: &str,
    start: usize,
    end: usize,
    faces: &[Face<'_>],
) -> Vec<ResolvedFontSegment> {
    let mut segments: Vec<ResolvedFontSegment> = Vec::new();
    for (relative, grapheme) in text[start..end].grapheme_indices(true) {
        let segment_start = start + relative;
        let segment_end = segment_start + grapheme.len();
        let face_index = faces
            .iter()
            .position(|face| font_covers_grapheme(face, grapheme))
            .unwrap_or(0);
        if let Some(previous) = segments.last_mut() {
            if previous.end == segment_start && previous.face_index == face_index {
                previous.end = segment_end;
                continue;
            }
        }
        segments.push(ResolvedFontSegment {
            start: segment_start,
            end: segment_end,
            face_index,
        });
    }
    if segments.is_empty() {
        segments.push(ResolvedFontSegment {
            start,
            end,
            face_index: 0,
        });
    }
    segments
}

fn font_covers_grapheme(face: &Face<'_>, grapheme: &str) -> bool {
    grapheme
        .chars()
        .filter(|character| character_requires_glyph(*character))
        .all(|character| face.glyph_index(character).is_some())
}

fn character_requires_glyph(character: char) -> bool {
    if character.is_control() {
        return false;
    }
    !matches!(
        character,
        '\u{200B}'
            | '\u{200C}'
            | '\u{200D}'
            | '\u{200E}'
            | '\u{200F}'
            | '\u{202A}'..='\u{202E}'
            | '\u{2060}'..='\u{206F}'
            | '\u{FE00}'..='\u{FE0F}'
            | '\u{FEFF}'
            | '\u{E0100}'..='\u{E01EF}'
    )
}

struct ResolvedTextTabs<'layout> {
    layout: &'layout TextTabLayout,
    segments: HashMap<usize, TextTabSegment>,
}

#[derive(Debug, Clone, Copy, Default, PartialEq)]
struct TextTabSegment {
    width: f64,
    decimal_width: Option<f64>,
}

#[derive(Debug, Clone, Copy, PartialEq)]
struct EffectiveTextTabStop {
    position: f64,
    alignment: TextTabAlignment,
}

fn resolve_text_tabs<'layout>(
    text: &str,
    runs: &[ResolvedTextLayoutRun<'_, '_>],
    bidi: &ResolvedBidiText,
    layout: &'layout TextTabLayout,
) -> ResolvedTextTabs<'layout> {
    let mut segments = HashMap::new();
    for (tab_start, _) in text.match_indices('\t') {
        let segment_start = tab_start + '\t'.len_utf8();
        let segment_end = text[segment_start..]
            .char_indices()
            .find_map(|(offset, character)| {
                matches!(character, '\t' | '\n' | '\r' | '\u{2028}' | '\u{2029}')
                    .then_some(segment_start + offset)
            })
            .unwrap_or(text.len());
        let width = shape_plain_text_range(text, segment_start, segment_end, runs, bidi).width;
        let decimal_end =
            text[segment_start..segment_end]
                .char_indices()
                .find_map(|(offset, character)| {
                    is_text_tab_decimal_separator(character).then_some(segment_start + offset)
                });
        let decimal_width = decimal_end
            .map(|end| shape_plain_text_range(text, segment_start, end, runs, bidi).width);
        segments.insert(
            tab_start,
            TextTabSegment {
                width,
                decimal_width,
            },
        );
    }
    ResolvedTextTabs { layout, segments }
}

fn next_text_tab_stop(layout: &TextTabLayout, current_position: f64) -> EffectiveTextTabStop {
    let current = current_position.max(0.0);
    let custom = layout
        .stops
        .iter()
        .find(|stop| stop.position > current + 0.5);
    let default_position = (current / layout.default_interval)
        .floor()
        .mul_add(layout.default_interval, layout.default_interval);
    if let Some(stop) = custom.filter(|stop| stop.position <= default_position + 0.5) {
        return EffectiveTextTabStop {
            position: stop.position,
            alignment: stop.alignment,
        };
    }
    EffectiveTextTabStop {
        position: default_position,
        alignment: TextTabAlignment::Left,
    }
}

fn is_text_tab_decimal_separator(character: char) -> bool {
    matches!(character, '.' | ',' | '，' | '。')
}

#[derive(Debug, Clone, Copy, Default, PartialEq)]
struct TextLineMetrics {
    ascent: f64,
    descent: f64,
    height: f64,
}

#[derive(Debug, Clone, Copy, Default, PartialEq)]
struct ShapedRun {
    width: f64,
    glyph_count: u32,
    fallback_glyph_count: u32,
    missing_glyph_count: u32,
}

fn shape_text_range(
    text: &str,
    start: usize,
    end: usize,
    runs: &[ResolvedTextLayoutRun<'_, '_>],
    bidi: &ResolvedBidiText,
    tabs: Option<&ResolvedTextTabs<'_>>,
    first_line: bool,
) -> ShapedRun {
    let Some(tabs) = tabs else {
        return shape_plain_text_range(text, start, end, runs, bidi);
    };
    let mut shaped = ShapedRun::default();
    let mut cursor = start;
    for (relative, character) in text[start..end].char_indices() {
        if character != '\t' {
            continue;
        }
        let tab_start = start + relative;
        append_shaped_run(
            &mut shaped,
            shape_plain_text_range(text, cursor, tab_start, runs, bidi),
        );
        let current_position = tabs.layout.origin
            + if first_line {
                tabs.layout.first_line_indent
            } else {
                0.0
            }
            + shaped.width;
        let stop = next_text_tab_stop(tabs.layout, current_position);
        let segment = tabs.segments.get(&tab_start).copied().unwrap_or_default();
        let alignment_offset = text_tab_alignment_offset(stop.alignment, segment);
        shaped.width += (stop.position - current_position - alignment_offset).max(MIN_TAB_ADVANCE);
        cursor = tab_start + character.len_utf8();
    }
    append_shaped_run(
        &mut shaped,
        shape_plain_text_range(text, cursor, end, runs, bidi),
    );
    shaped.width = round_layout_value(shaped.width.max(0.0));
    shaped
}

fn shape_plain_text_range(
    text: &str,
    start: usize,
    end: usize,
    runs: &[ResolvedTextLayoutRun<'_, '_>],
    bidi: &ResolvedBidiText,
) -> ShapedRun {
    let mut shaped = ShapedRun {
        width: 0.0,
        glyph_count: 0,
        fallback_glyph_count: 0,
        missing_glyph_count: 0,
    };
    for resolved in runs {
        let segment_start = start.max(resolved.start);
        let segment_end = end.min(resolved.end);
        if segment_start >= segment_end {
            continue;
        }
        let mut run_shaped = ShapedRun::default();
        for bidi_run in bidi.directional_runs(segment_start, segment_end) {
            for font_segment in &resolved.font_segments {
                let font_start = bidi_run.start.max(font_segment.start);
                let font_end = bidi_run.end.min(font_segment.end);
                if font_start >= font_end {
                    continue;
                }
                append_shaped_run(
                    &mut run_shaped,
                    shape_run(
                        &resolved.faces[font_segment.face_index],
                        resolved.run,
                        bidi_run.direction,
                        &text[font_start..font_end],
                    )
                    .with_fallback(font_segment.face_index > 0),
                );
            }
        }
        run_shaped.width +=
            resolved.run.letter_spacing * f64::from(run_shaped.glyph_count.saturating_sub(1));
        run_shaped.width = round_layout_value(run_shaped.width.max(0.0));
        append_shaped_run(&mut shaped, run_shaped);
    }
    shaped.width = round_layout_value(shaped.width.max(0.0));
    shaped
}

fn append_shaped_run(target: &mut ShapedRun, source: ShapedRun) {
    target.width += source.width;
    target.glyph_count = target.glyph_count.saturating_add(source.glyph_count);
    target.fallback_glyph_count = target
        .fallback_glyph_count
        .saturating_add(source.fallback_glyph_count);
    target.missing_glyph_count = target
        .missing_glyph_count
        .saturating_add(source.missing_glyph_count);
}

impl ShapedRun {
    fn with_fallback(mut self, fallback: bool) -> Self {
        if fallback {
            self.fallback_glyph_count = self.glyph_count;
        }
        self
    }
}

fn text_tab_alignment_offset(alignment: TextTabAlignment, segment: TextTabSegment) -> f64 {
    match alignment {
        TextTabAlignment::Left => 0.0,
        TextTabAlignment::Center => segment.width / 2.0,
        TextTabAlignment::Right => segment.width,
        TextTabAlignment::Decimal => segment.decimal_width.unwrap_or(segment.width),
    }
}

fn text_line_metrics(
    start: usize,
    end: usize,
    runs: &[ResolvedTextLayoutRun<'_, '_>],
) -> TextLineMetrics {
    let mut metrics = TextLineMetrics::default();
    let mut matched = false;
    for resolved in runs {
        let run_matches = if end > start {
            resolved.end > start && resolved.start < end
        } else {
            resolved.start <= start && start <= resolved.end
        };
        if !run_matches {
            continue;
        }
        matched = true;
        merge_text_line_metrics(
            &mut metrics,
            text_run_metrics(&resolved.faces[0], resolved.run),
        );
        for segment in &resolved.font_segments {
            if end > start {
                if segment.end <= start || segment.start >= end {
                    continue;
                }
            } else if !(segment.start <= start && start <= segment.end) {
                continue;
            }
            merge_text_line_metrics(
                &mut metrics,
                text_run_metrics(&resolved.faces[segment.face_index], resolved.run),
            );
        }
    }
    if !matched {
        if let Some(resolved) = runs.first() {
            metrics = text_run_metrics(&resolved.faces[0], resolved.run);
        }
    }
    metrics
}

fn merge_text_line_metrics(target: &mut TextLineMetrics, source: TextLineMetrics) {
    target.ascent = target.ascent.max(source.ascent);
    target.descent = target.descent.max(source.descent);
    target.height = target.height.max(source.height);
}

fn text_run_metrics(face: &Face<'_>, run: &TextLayoutRun) -> TextLineMetrics {
    let units_per_em = f64::from(face.units_per_em()).max(1.0);
    let scale = run.font_size / units_per_em;
    TextLineMetrics {
        ascent: round_layout_value(f64::from(face.ascender()).max(0.0) * scale),
        descent: round_layout_value(f64::from(face.descender()).min(0.0).abs() * scale),
        height: round_layout_value(run.line_height),
    }
}

fn shape_run(
    face: &Face<'_>,
    run: &TextLayoutRun,
    direction: TextDirection,
    text: &str,
) -> ShapedRun {
    if text.is_empty() {
        return ShapedRun {
            width: 0.0,
            glyph_count: 0,
            fallback_glyph_count: 0,
            missing_glyph_count: 0,
        };
    }
    let mut buffer = UnicodeBuffer::new();
    buffer.push_str(text);
    match direction {
        TextDirection::Auto => buffer.guess_segment_properties(),
        TextDirection::Ltr => buffer.set_direction(Direction::LeftToRight),
        TextDirection::Rtl => buffer.set_direction(Direction::RightToLeft),
    }
    let mut features = Vec::new();
    if !run.ligatures {
        features.extend(
            [b"liga", b"clig", b"dlig", b"hlig"]
                .map(|tag| Feature::new(Tag::from_bytes(tag), 0, ..)),
        );
    }
    if !run.kerning {
        features.push(Feature::new(Tag::from_bytes(b"kern"), 0, ..));
    }
    let glyphs = shape(face, &features, buffer);
    let glyph_count = u32::try_from(glyphs.len()).unwrap_or(u32::MAX);
    let missing_glyph_count = u32::try_from(
        glyphs
            .glyph_infos()
            .iter()
            .filter(|glyph| glyph.glyph_id == 0)
            .count(),
    )
    .unwrap_or(u32::MAX);
    let units_per_em = f64::from(face.units_per_em()).max(1.0);
    let advance = glyphs
        .glyph_positions()
        .iter()
        .map(|position| f64::from(position.x_advance))
        .sum::<f64>()
        .abs()
        * run.font_size
        / units_per_em;
    ShapedRun {
        width: round_layout_value(advance.max(0.0)),
        glyph_count,
        fallback_glyph_count: 0,
        missing_glyph_count,
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
struct LineRange {
    start: usize,
    visible_end: usize,
    next_start: usize,
    hard_break: bool,
}

fn break_line_ranges(
    text: &str,
    max_width: f64,
    first_line_max_width: f64,
    white_space: TextWhiteSpace,
    mut measure: impl FnMut(usize, usize) -> f64,
) -> Vec<LineRange> {
    if text.is_empty() {
        return vec![LineRange {
            start: 0,
            visible_end: 0,
            next_start: 0,
            hard_break: false,
        }];
    }

    let opportunities = text_break_opportunities(text, white_space);
    let mut lines = Vec::new();
    let mut line_start = 0;
    let mut last_fitting_break = None;
    let mut cursor = 0;

    while cursor < opportunities.len() {
        let (candidate_end, opportunity) = opportunities[cursor];
        if candidate_end <= line_start {
            cursor += 1;
            continue;
        }
        let visible_end = visible_line_end(text, line_start, candidate_end, white_space);
        let width_limit = if lines.is_empty() {
            first_line_max_width
        } else {
            max_width
        };
        let width = measure(line_start, visible_end);

        if width <= width_limit || visible_end == line_start {
            if opportunity == BreakOpportunity::Mandatory {
                lines.push(LineRange {
                    start: line_start,
                    visible_end,
                    next_start: candidate_end,
                    hard_break: has_explicit_line_break(text, candidate_end),
                });
                line_start = candidate_end;
                last_fitting_break = None;
            } else {
                last_fitting_break = Some(candidate_end);
            }
            cursor += 1;
            continue;
        }

        if let Some(line_end) = last_fitting_break.filter(|end| *end > line_start) {
            lines.push(LineRange {
                start: line_start,
                visible_end: visible_line_end(text, line_start, line_end, white_space),
                next_start: line_end,
                hard_break: false,
            });
            line_start = line_end;
            last_fitting_break = None;
            continue;
        }

        let emergency_end = emergency_break(
            text,
            line_start,
            visible_end.max(candidate_end),
            width_limit,
            &mut measure,
        );
        lines.push(LineRange {
            start: line_start,
            visible_end: emergency_end,
            next_start: emergency_end,
            hard_break: false,
        });
        line_start = emergency_end;
        last_fitting_break = None;
    }

    if line_start < text.len() {
        lines.push(LineRange {
            start: line_start,
            visible_end: visible_line_end(text, line_start, text.len(), white_space),
            next_start: text.len(),
            hard_break: false,
        });
    }
    lines
}

fn text_break_opportunities(
    text: &str,
    white_space: TextWhiteSpace,
) -> Vec<(usize, BreakOpportunity)> {
    let mut opportunities = linebreaks(text).collect::<Vec<_>>();
    if white_space == TextWhiteSpace::BreakSpaces {
        opportunities.extend(text.char_indices().filter_map(|(index, character)| {
            matches!(character, ' ' | '\t')
                .then_some((index + character.len_utf8(), BreakOpportunity::Allowed))
        }));
        opportunities.sort_by_key(|(index, opportunity)| {
            (
                *index,
                if *opportunity == BreakOpportunity::Mandatory {
                    0
                } else {
                    1
                },
            )
        });
        opportunities.dedup_by_key(|(index, _)| *index);
    }
    opportunities
}

fn emergency_break(
    text: &str,
    start: usize,
    end: usize,
    width_limit: f64,
    measure: &mut impl FnMut(usize, usize) -> f64,
) -> usize {
    let mut first_end = None;
    let mut fitting_end = None;
    for (offset, grapheme) in text[start..end].grapheme_indices(true) {
        let grapheme_end = start + offset + grapheme.len();
        first_end.get_or_insert(grapheme_end);
        if measure(start, grapheme_end) <= width_limit {
            fitting_end = Some(grapheme_end);
        } else {
            break;
        }
    }
    fitting_end.or(first_end).unwrap_or(end)
}

fn visible_line_end(text: &str, start: usize, end: usize, white_space: TextWhiteSpace) -> usize {
    let mut visible_end = end;
    while visible_end > start {
        let Some(character) = text[..visible_end].chars().next_back() else {
            break;
        };
        let should_trim = match white_space {
            TextWhiteSpace::Normal => character.is_whitespace(),
            TextWhiteSpace::BreakSpaces => {
                matches!(character, '\n' | '\r' | '\u{2028}' | '\u{2029}')
            }
        };
        if !should_trim {
            break;
        }
        visible_end -= character.len_utf8();
    }
    visible_end
}

fn has_explicit_line_break(text: &str, end: usize) -> bool {
    text[..end]
        .chars()
        .next_back()
        .is_some_and(|character| matches!(character, '\n' | '\r' | '\u{2028}' | '\u{2029}'))
}

fn round_layout_value(value: f64) -> f64 {
    (value * 1_000.0).round() / 1_000.0
}

const fn default_ligatures() -> bool {
    true
}

const fn default_kerning() -> bool {
    true
}

#[cfg(test)]
mod tests;
