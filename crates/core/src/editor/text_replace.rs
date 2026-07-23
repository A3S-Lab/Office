use std::collections::BTreeSet;

use a3s_use_core::{UseError, UseResult};
use regex::Regex;

use super::{
    editor_error, node_not_found, parse_segments, NativeOfficeTextMatchMode,
    NativeOfficeTextReplacement, NativeOfficeTextReplacementResult, MAX_NATIVE_OFFICE_TEXT_MATCHES,
    MAX_NATIVE_OFFICE_TEXT_REPLACEMENT_OUTPUT_BYTES,
};
use crate::semantic::{NativeOfficeDocument, OfficeNodeType};
use crate::xml_edit::{
    apply_patches, decoded_element_text, index_xml, replace_element_text_patch, IndexedXmlElement,
    XmlPatch,
};
use crate::{DocumentKind, LosslessXmlPart, NativeOfficePackage};

mod spreadsheet;

const WORD_NAMESPACE: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const STRICT_WORD_NAMESPACE: &str = "http://purl.oclc.org/ooxml/wordprocessingml/main";
const DRAWING_NAMESPACE: &str = "http://schemas.openxmlformats.org/drawingml/2006/main";
const STRICT_DRAWING_NAMESPACE: &str = "http://purl.oclc.org/ooxml/drawingml/main";

#[derive(Debug, Clone, Copy)]
enum TextDialect {
    Word,
    Drawing,
}

impl TextDialect {
    fn accepts(self, element: &IndexedXmlElement) -> bool {
        matches!(
            (self, element.namespace.as_deref()),
            (Self::Word, Some(WORD_NAMESPACE | STRICT_WORD_NAMESPACE))
                | (
                    Self::Drawing,
                    Some(DRAWING_NAMESPACE | STRICT_DRAWING_NAMESPACE)
                )
        )
    }
}

pub(super) fn replace(
    package: &mut NativeOfficePackage,
    path: &str,
    replacement: &NativeOfficeTextReplacement,
) -> UseResult<NativeOfficeTextReplacementResult> {
    validate_scope_path(path)?;
    replacement.validate()?;
    let compiled = CompiledTextReplacement::new(replacement)?;
    let mut accumulator = ReplacementAccumulator::default();
    match package.kind() {
        DocumentKind::Word => replace_word(package, path, &compiled, &mut accumulator)?,
        DocumentKind::Spreadsheet => {
            spreadsheet::replace(package, path, &compiled, &mut accumulator)?
        }
        DocumentKind::Presentation => {
            replace_presentation(package, path, &compiled, &mut accumulator)?
        }
    }
    Ok(NativeOfficeTextReplacementResult {
        path: path.to_string(),
        mode: replacement.mode,
        match_count: accumulator.match_count,
        changed: !accumulator.changed_parts.is_empty(),
        changed_parts: accumulator.changed_parts.into_iter().collect(),
    })
}

fn validate_scope_path(path: &str) -> UseResult<()> {
    if path == "/" {
        return Ok(());
    }
    super::validate_mutation_path(path)
}

#[derive(Debug)]
pub(super) struct CompiledTextReplacement {
    mode: NativeOfficeTextMatchMode,
    find: String,
    replace: String,
    regex: Option<Regex>,
}

#[derive(Debug)]
pub(super) struct SegmentTransform {
    pub(super) output: Vec<String>,
    pub(super) match_count: usize,
    pub(super) replacement_bytes: usize,
    pub(super) changed: bool,
}

#[derive(Debug)]
struct ReplacementSpan {
    start: usize,
    end: usize,
    value: String,
}

impl CompiledTextReplacement {
    fn new(replacement: &NativeOfficeTextReplacement) -> UseResult<Self> {
        let regex = (replacement.mode == NativeOfficeTextMatchMode::Regex)
            .then(|| Regex::new(&replacement.find))
            .transpose()
            .map_err(|error| {
                editor_error(
                    "use.office.text_regex_invalid",
                    format!("Native Office regular expression is invalid: {error}"),
                )
            })?;
        Ok(Self {
            mode: replacement.mode,
            find: replacement.find.clone(),
            replace: replacement.replace.clone(),
            regex,
        })
    }

    pub(super) fn transform(&self, segments: &[String]) -> UseResult<SegmentTransform> {
        let boundaries = segment_boundaries(segments)?;
        let mut source = String::with_capacity(*boundaries.last().unwrap_or(&0));
        for segment in segments {
            source.push_str(segment);
        }
        let mut spans = Vec::new();
        let mut replacement_bytes = 0_usize;
        match self.mode {
            NativeOfficeTextMatchMode::Literal => {
                for (start, _) in source.match_indices(&self.find) {
                    ensure_match_slot(spans.len())?;
                    push_replacement_span(
                        &mut spans,
                        &mut replacement_bytes,
                        ReplacementSpan {
                            start,
                            end: start + self.find.len(),
                            value: self.replace.clone(),
                        },
                    )?;
                }
            }
            NativeOfficeTextMatchMode::Regex => {
                let expression = self.regex.as_ref().ok_or_else(|| {
                    editor_error(
                        "use.office.text_regex_invalid",
                        "Native Office regular expression was not compiled.",
                    )
                })?;
                for captures in expression.captures_iter(&source) {
                    ensure_match_slot(spans.len())?;
                    let matched = captures.get(0).ok_or_else(|| {
                        editor_error(
                            "use.office.text_regex_invalid",
                            "Native Office regular expression produced no complete match.",
                        )
                    })?;
                    if matched.start() == matched.end() {
                        return Err(editor_error(
                            "use.office.text_regex_empty_match",
                            "Native Office regular expressions must consume at least one character.",
                        ));
                    }
                    let value =
                        expand_regex_replacement(&captures, &self.replace, replacement_bytes)?;
                    validate_expanded_xml_text(&value)?;
                    push_replacement_span(
                        &mut spans,
                        &mut replacement_bytes,
                        ReplacementSpan {
                            start: matched.start(),
                            end: matched.end(),
                            value,
                        },
                    )?;
                }
            }
        }
        if spans.is_empty() {
            return Ok(SegmentTransform {
                output: segments.to_vec(),
                match_count: 0,
                replacement_bytes: 0,
                changed: false,
            });
        }

        let mut output = vec![String::new(); segments.len()];
        let mut cursor = 0_usize;
        for span in &spans {
            append_original(segments, &boundaries, cursor, span.start, &mut output)?;
            let owner = owner_at(&boundaries, span.start).ok_or_else(|| {
                editor_error(
                    "use.office.text_replacement_invalid",
                    "Native Office text match has no source run owner.",
                )
            })?;
            output[owner].push_str(&span.value);
            cursor = span.end;
        }
        append_original(segments, &boundaries, cursor, source.len(), &mut output)?;
        let changed = output != segments;
        Ok(SegmentTransform {
            output,
            match_count: spans.len(),
            replacement_bytes,
            changed,
        })
    }
}

fn ensure_match_slot(matches: usize) -> UseResult<()> {
    if matches < MAX_NATIVE_OFFICE_TEXT_MATCHES {
        return Ok(());
    }
    Err(replacement_limit_error(format!(
        "Native Office text replacement exceeds {MAX_NATIVE_OFFICE_TEXT_MATCHES} matches."
    ))
    .with_detail("matches", matches.saturating_add(1)))
}

fn push_replacement_span(
    spans: &mut Vec<ReplacementSpan>,
    replacement_bytes: &mut usize,
    span: ReplacementSpan,
) -> UseResult<()> {
    let total = replacement_bytes
        .checked_add(span.value.len())
        .ok_or_else(|| {
            replacement_limit_error("Native Office replacement output size overflowed.")
        })?;
    if total > MAX_NATIVE_OFFICE_TEXT_REPLACEMENT_OUTPUT_BYTES {
        return Err(replacement_limit_error(format!(
            "Native Office expanded replacement output exceeds {MAX_NATIVE_OFFICE_TEXT_REPLACEMENT_OUTPUT_BYTES} bytes."
        ))
        .with_detail("replacementBytes", total));
    }
    *replacement_bytes = total;
    spans.push(span);
    Ok(())
}

#[derive(Debug, Clone, Copy)]
enum CaptureReference<'a> {
    Number(usize),
    Name(&'a str),
}

fn expand_regex_replacement(
    captures: &regex::Captures<'_>,
    replacement: &str,
    already_expanded: usize,
) -> UseResult<String> {
    let mut remaining = replacement;
    let mut output = String::new();
    while let Some(position) = remaining.find('$') {
        push_bounded_replacement(&mut output, &remaining[..position], already_expanded)?;
        remaining = &remaining[position..];
        if remaining.as_bytes().get(1) == Some(&b'$') {
            push_bounded_replacement(&mut output, "$", already_expanded)?;
            remaining = &remaining[2..];
            continue;
        }
        let Some((reference, end)) = capture_reference(remaining) else {
            push_bounded_replacement(&mut output, "$", already_expanded)?;
            remaining = &remaining[1..];
            continue;
        };
        remaining = &remaining[end..];
        let value = match reference {
            CaptureReference::Number(index) => captures.get(index),
            CaptureReference::Name(name) => captures.name(name),
        };
        if let Some(value) = value {
            push_bounded_replacement(&mut output, value.as_str(), already_expanded)?;
        }
    }
    push_bounded_replacement(&mut output, remaining, already_expanded)?;
    Ok(output)
}

fn capture_reference(replacement: &str) -> Option<(CaptureReference<'_>, usize)> {
    let bytes = replacement.as_bytes();
    if bytes.first() != Some(&b'$') || bytes.len() <= 1 {
        return None;
    }
    if bytes[1] == b'{' {
        let end = bytes[2..].iter().position(|byte| *byte == b'}')? + 2;
        let name = &replacement[2..end];
        let reference = name
            .parse::<usize>()
            .map(CaptureReference::Number)
            .unwrap_or(CaptureReference::Name(name));
        return Some((reference, end + 1));
    }
    let end = bytes[1..]
        .iter()
        .position(|byte| !byte.is_ascii_alphanumeric() && *byte != b'_')
        .map_or(bytes.len(), |position| position + 1);
    if end == 1 {
        return None;
    }
    let name = &replacement[1..end];
    let reference = name
        .parse::<usize>()
        .map(CaptureReference::Number)
        .unwrap_or(CaptureReference::Name(name));
    Some((reference, end))
}

fn push_bounded_replacement(
    output: &mut String,
    value: &str,
    already_expanded: usize,
) -> UseResult<()> {
    let expanded = output.len().checked_add(value.len()).ok_or_else(|| {
        replacement_limit_error("Native Office replacement output size overflowed.")
    })?;
    let total = already_expanded.checked_add(expanded).ok_or_else(|| {
        replacement_limit_error("Native Office replacement output size overflowed.")
    })?;
    if total > MAX_NATIVE_OFFICE_TEXT_REPLACEMENT_OUTPUT_BYTES {
        return Err(replacement_limit_error(format!(
            "Native Office expanded replacement output exceeds {MAX_NATIVE_OFFICE_TEXT_REPLACEMENT_OUTPUT_BYTES} bytes."
        ))
        .with_detail("replacementBytes", total));
    }
    output.push_str(value);
    Ok(())
}

fn segment_boundaries(segments: &[String]) -> UseResult<Vec<usize>> {
    let mut boundaries = Vec::<usize>::with_capacity(segments.len().saturating_add(1));
    boundaries.push(0);
    for segment in segments {
        let next = boundaries
            .last()
            .copied()
            .unwrap_or_default()
            .checked_add(segment.len())
            .ok_or_else(|| {
                replacement_limit_error("Native Office text segment size overflowed.")
            })?;
        boundaries.push(next);
    }
    Ok(boundaries)
}

fn append_original(
    segments: &[String],
    boundaries: &[usize],
    start: usize,
    end: usize,
    output: &mut [String],
) -> UseResult<()> {
    if start >= end {
        return Ok(());
    }
    let mut index = boundaries
        .partition_point(|boundary| *boundary <= start)
        .saturating_sub(1);
    while let Some(segment) = segments.get(index) {
        let segment_start = boundaries[index];
        let segment_end = boundaries[index + 1];
        if segment_start >= end {
            break;
        }
        let intersection_start = start.max(segment_start);
        let intersection_end = end.min(segment_end);
        if intersection_start < intersection_end {
            let local_start = intersection_start - segment_start;
            let local_end = intersection_end - segment_start;
            let original = segment.get(local_start..local_end).ok_or_else(|| {
                editor_error(
                    "use.office.text_replacement_invalid",
                    "Native Office text match did not align to UTF-8 boundaries.",
                )
            })?;
            output[index].push_str(original);
        }
        index = index.saturating_add(1);
    }
    Ok(())
}

fn owner_at(boundaries: &[usize], position: usize) -> Option<usize> {
    let index = boundaries
        .partition_point(|boundary| *boundary <= position)
        .checked_sub(1)?;
    (index + 1 < boundaries.len() && position < boundaries[index + 1]).then_some(index)
}

#[derive(Debug, Default)]
pub(super) struct ReplacementAccumulator {
    match_count: usize,
    replacement_bytes: usize,
    changed_parts: BTreeSet<String>,
}

impl ReplacementAccumulator {
    pub(super) fn record(
        &mut self,
        transform: &SegmentTransform,
        multiplier: usize,
    ) -> UseResult<()> {
        let matches = transform
            .match_count
            .checked_mul(multiplier)
            .ok_or_else(|| replacement_limit_error("Native Office text match count overflowed."))?;
        self.match_count = self
            .match_count
            .checked_add(matches)
            .ok_or_else(|| replacement_limit_error("Native Office text match count overflowed."))?;
        if self.match_count > MAX_NATIVE_OFFICE_TEXT_MATCHES {
            return Err(replacement_limit_error(format!(
                "Native Office text replacement exceeds {MAX_NATIVE_OFFICE_TEXT_MATCHES} matches."
            ))
            .with_detail("matches", self.match_count));
        }
        let bytes = transform
            .replacement_bytes
            .checked_mul(multiplier)
            .ok_or_else(|| {
                replacement_limit_error("Native Office replacement output size overflowed.")
            })?;
        self.replacement_bytes = self.replacement_bytes.checked_add(bytes).ok_or_else(|| {
            replacement_limit_error("Native Office replacement output size overflowed.")
        })?;
        if self.replacement_bytes > MAX_NATIVE_OFFICE_TEXT_REPLACEMENT_OUTPUT_BYTES {
            return Err(replacement_limit_error(format!(
                "Native Office expanded replacement output exceeds {MAX_NATIVE_OFFICE_TEXT_REPLACEMENT_OUTPUT_BYTES} bytes."
            ))
            .with_detail("replacementBytes", self.replacement_bytes));
        }
        Ok(())
    }

    pub(super) fn changed(&mut self, part: &str) {
        self.changed_parts
            .insert(format!("/{}", part.trim_start_matches('/')));
    }
}

pub(super) fn transform_text_elements(
    part: &LosslessXmlPart,
    elements: &[&IndexedXmlElement],
    compiled: &CompiledTextReplacement,
    accumulator: &mut ReplacementAccumulator,
    multiplier: usize,
    patches: &mut Vec<XmlPatch>,
) -> UseResult<SegmentTransform> {
    let source = elements
        .iter()
        .map(|element| decoded_element_text(part, element))
        .collect::<UseResult<Vec<_>>>()?;
    let transform = compiled.transform(&source)?;
    accumulator.record(&transform, multiplier)?;
    if transform.changed {
        for ((element, before), after) in elements.iter().zip(&source).zip(&transform.output) {
            if before != after {
                patches.push(replace_element_text_patch(element, after));
            }
        }
    }
    Ok(transform)
}

fn replace_word(
    package: &mut NativeOfficePackage,
    path: &str,
    compiled: &CompiledTextReplacement,
    accumulator: &mut ReplacementAccumulator,
) -> UseResult<()> {
    if path == "/" {
        for part_name in word_text_parts(package) {
            replace_whole_paragraph_part(
                package,
                &part_name,
                TextDialect::Word,
                compiled,
                accumulator,
            )?;
        }
        return Ok(());
    }

    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let requested = snapshot.get(path, 0)?;
    if !matches!(
        requested.node_type,
        OfficeNodeType::Body
            | OfficeNodeType::Header
            | OfficeNodeType::Footer
            | OfficeNodeType::Paragraph
            | OfficeNodeType::Run
            | OfficeNodeType::Hyperlink
            | OfficeNodeType::Comment
            | OfficeNodeType::Table
            | OfficeNodeType::TableRow
            | OfficeNodeType::TableCell
    ) {
        return Err(unsupported_scope(path, "Word"));
    }
    let part_name = word_scope_part(&snapshot, &requested, path)?;
    let part = package.xml_part(&part_name)?;
    let root = index_xml(&part)?;
    let target = locate_word_scope(&root, path)?;
    let mut patches = Vec::new();
    replace_element_paragraphs(
        &part,
        target,
        TextDialect::Word,
        compiled,
        accumulator,
        &mut patches,
    )?;
    if !patches.is_empty() {
        package.set_part(&part_name, apply_patches(&part, patches)?)?;
        accumulator.changed(&part_name);
    }
    Ok(())
}

fn word_scope_part(
    snapshot: &NativeOfficeDocument,
    requested: &crate::DocumentNode,
    path: &str,
) -> UseResult<String> {
    if path.starts_with("/body") {
        return Ok("word/document.xml".to_string());
    }
    if let Some(part) = requested.format.get("part") {
        return Ok(part.trim_start_matches('/').to_string());
    }
    let first = path
        .trim_start_matches('/')
        .split('/')
        .next()
        .ok_or_else(|| node_not_found(path))?;
    let owner_path = format!("/{first}");
    snapshot
        .get(&owner_path, 0)?
        .format
        .get("part")
        .map(|part| part.trim_start_matches('/').to_string())
        .ok_or_else(|| {
            editor_error(
                "use.office.text_scope_invalid",
                format!("Word scope '{path}' has no source OOXML part."),
            )
        })
}

fn locate_word_scope<'a>(
    root: &'a IndexedXmlElement,
    path: &str,
) -> UseResult<&'a IndexedXmlElement> {
    let segments = parse_segments(path)?;
    let skip_virtual = matches!(
        segments.first().map(|segment| segment.name.as_str()),
        Some("header" | "footer" | "comments")
    );
    let mut current = root;
    let mut paragraph = None;
    for segment in segments.iter().skip(usize::from(skip_virtual)) {
        let position = segment.position.unwrap_or(1);
        current = match segment.name.as_str() {
            "body" => current.child("body", position),
            "p" | "paragraph" => current.child("p", position),
            "tbl" | "table" => current.child("tbl", position),
            "tr" => current.child("tr", position),
            "tc" | "cell" => current.child("tc", position),
            "comment" => current.child("comment", position),
            "hyperlink" => current.child("hyperlink", position),
            "r" | "run" => {
                let owner = paragraph.unwrap_or(current);
                let mut runs = Vec::new();
                collect_word_runs(owner, &mut runs);
                let candidate = runs.get(position.saturating_sub(1)).copied();
                candidate.filter(|candidate| {
                    current.local_name != "hyperlink"
                        || (current.full_range.start <= candidate.full_range.start
                            && candidate.full_range.end <= current.full_range.end)
                })
            }
            _ => None,
        }
        .ok_or_else(|| node_not_found(path))?;
        if current.local_name == "p" {
            paragraph = Some(current);
        }
    }
    Ok(current)
}

fn collect_word_runs<'a>(element: &'a IndexedXmlElement, output: &mut Vec<&'a IndexedXmlElement>) {
    for child in &element.children {
        if child.local_name == "p" {
            continue;
        }
        if child.local_name == "r" {
            if !is_comment_reference_run(child) {
                output.push(child);
            }
        } else {
            collect_word_runs(child, output);
        }
    }
}

fn is_comment_reference_run(run: &IndexedXmlElement) -> bool {
    let meaningful = run
        .children
        .iter()
        .filter(|child| child.local_name != "rPr")
        .collect::<Vec<_>>();
    meaningful.len() == 1 && meaningful[0].local_name == "commentReference"
}

fn replace_presentation(
    package: &mut NativeOfficePackage,
    path: &str,
    compiled: &CompiledTextReplacement,
    accumulator: &mut ReplacementAccumulator,
) -> UseResult<()> {
    if path == "/" {
        for part_name in presentation_text_parts(package) {
            replace_whole_paragraph_part(
                package,
                &part_name,
                TextDialect::Drawing,
                compiled,
                accumulator,
            )?;
        }
        return Ok(());
    }
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let requested = snapshot.get(path, 0)?;
    if requested.node_type == OfficeNodeType::Notes {
        let part_name = requested.format.get("part").ok_or_else(|| {
            editor_error(
                "use.office.text_scope_invalid",
                format!("Presentation notes scope '{path}' has no source OOXML part."),
            )
        })?;
        return replace_whole_paragraph_part(
            package,
            part_name,
            TextDialect::Drawing,
            compiled,
            accumulator,
        );
    }
    if !matches!(
        requested.node_type,
        OfficeNodeType::Slide
            | OfficeNodeType::Shape
            | OfficeNodeType::Placeholder
            | OfficeNodeType::Group
            | OfficeNodeType::Table
            | OfficeNodeType::TableRow
            | OfficeNodeType::TableCell
            | OfficeNodeType::Paragraph
            | OfficeNodeType::Run
    ) {
        return Err(unsupported_scope(path, "Presentation"));
    }
    let slide_path = format!(
        "/{}",
        path.trim_start_matches('/')
            .split('/')
            .next()
            .ok_or_else(|| node_not_found(path))?
    );
    let slide = snapshot.get(&slide_path, 0)?;
    let part_name = slide.format.get("part").ok_or_else(|| {
        editor_error(
            "use.office.text_scope_invalid",
            format!("Presentation scope '{path}' has no source slide part."),
        )
    })?;
    let part = package.xml_part(part_name)?;
    let root = index_xml(&part)?;
    let target = super::presentation::locate_path(&root, path)?;
    let mut patches = Vec::new();
    replace_element_paragraphs(
        &part,
        target,
        TextDialect::Drawing,
        compiled,
        accumulator,
        &mut patches,
    )?;
    if !patches.is_empty() {
        package.set_part(part_name, apply_patches(&part, patches)?)?;
        accumulator.changed(part_name);
    }
    Ok(())
}

fn replace_whole_paragraph_part(
    package: &mut NativeOfficePackage,
    part_name: &str,
    dialect: TextDialect,
    compiled: &CompiledTextReplacement,
    accumulator: &mut ReplacementAccumulator,
) -> UseResult<()> {
    let part_name = part_name.trim_start_matches('/');
    let part = package.xml_part(part_name)?;
    let root = index_xml(&part)?;
    let mut patches = Vec::new();
    replace_element_paragraphs(&part, &root, dialect, compiled, accumulator, &mut patches)?;
    if !patches.is_empty() {
        package.set_part(part_name, apply_patches(&part, patches)?)?;
        accumulator.changed(part_name);
    }
    Ok(())
}

fn replace_element_paragraphs(
    part: &LosslessXmlPart,
    target: &IndexedXmlElement,
    dialect: TextDialect,
    compiled: &CompiledTextReplacement,
    accumulator: &mut ReplacementAccumulator,
    patches: &mut Vec<XmlPatch>,
) -> UseResult<()> {
    if target.local_name == "p" && dialect.accepts(target) {
        return replace_text_group(part, target, dialect, compiled, accumulator, patches);
    }
    let mut paragraphs = Vec::new();
    collect_paragraphs(target, dialect, &mut paragraphs);
    if paragraphs.is_empty() && matches!(target.local_name.as_str(), "r" | "hyperlink") {
        return replace_text_group(part, target, dialect, compiled, accumulator, patches);
    }
    for paragraph in paragraphs {
        replace_text_group(part, paragraph, dialect, compiled, accumulator, patches)?;
    }
    Ok(())
}

fn replace_text_group(
    part: &LosslessXmlPart,
    element: &IndexedXmlElement,
    dialect: TextDialect,
    compiled: &CompiledTextReplacement,
    accumulator: &mut ReplacementAccumulator,
    patches: &mut Vec<XmlPatch>,
) -> UseResult<()> {
    let mut text_elements = Vec::new();
    collect_text_elements(element, dialect, &mut text_elements);
    if !text_elements.is_empty() {
        transform_text_elements(part, &text_elements, compiled, accumulator, 1, patches)?;
    }
    Ok(())
}

fn collect_paragraphs<'a>(
    element: &'a IndexedXmlElement,
    dialect: TextDialect,
    output: &mut Vec<&'a IndexedXmlElement>,
) {
    for child in &element.children {
        if child.local_name == "p" && dialect.accepts(child) {
            output.push(child);
        } else {
            collect_paragraphs(child, dialect, output);
        }
    }
}

fn collect_text_elements<'a>(
    element: &'a IndexedXmlElement,
    dialect: TextDialect,
    output: &mut Vec<&'a IndexedXmlElement>,
) {
    for child in &element.children {
        if child.local_name == "p" && dialect.accepts(child) {
            continue;
        }
        if child.local_name == "t" && dialect.accepts(child) {
            output.push(child);
        } else {
            collect_text_elements(child, dialect, output);
        }
    }
}

fn word_text_parts(package: &NativeOfficePackage) -> Vec<String> {
    package
        .part_names()
        .filter(|name| {
            *name == "word/document.xml"
                || *name == "word/footnotes.xml"
                || *name == "word/endnotes.xml"
                || *name == "word/comments.xml"
                || (name.starts_with("word/header") && name.ends_with(".xml"))
                || (name.starts_with("word/footer") && name.ends_with(".xml"))
        })
        .map(str::to_string)
        .collect()
}

fn presentation_text_parts(package: &NativeOfficePackage) -> Vec<String> {
    package
        .part_names()
        .filter(|name| {
            (name.starts_with("ppt/slides/slide") || name.starts_with("ppt/notesSlides/notesSlide"))
                && name.ends_with(".xml")
        })
        .map(str::to_string)
        .collect()
}

fn validate_expanded_xml_text(value: &str) -> UseResult<()> {
    if let Some(character) = value.chars().find(|character| {
        !matches!(*character, '\u{9}' | '\u{a}' | '\u{d}')
            && (*character < '\u{20}' || matches!(*character, '\u{fffe}' | '\u{ffff}'))
    }) {
        return Err(editor_error(
            "use.office.text_replacement_invalid",
            format!(
                "Native Office expanded replacement contains XML-forbidden character U+{:04X}.",
                u32::from(character)
            ),
        ));
    }
    Ok(())
}

fn unsupported_scope(path: &str, format: &str) -> UseError {
    editor_error(
        "use.office.text_scope_unsupported",
        format!("Native {format} path '{path}' is not a supported text replacement scope."),
    )
    .with_detail("path", path)
}

fn replacement_limit_error(message: impl Into<String>) -> UseError {
    editor_error("use.office.text_replacement_limit", message)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn literal_transform_preserves_split_run_ownership() {
        let replacement = NativeOfficeTextReplacement::literal("alpha beta", "done").unwrap();
        let compiled = CompiledTextReplacement::new(&replacement).unwrap();
        let result = compiled
            .transform(&["before alpha ".into(), "beta after".into()])
            .unwrap();
        assert_eq!(result.output, ["before done", " after"]);
        assert_eq!(result.match_count, 1);
    }

    #[test]
    fn regex_transform_expands_captures_once() {
        let replacement = NativeOfficeTextReplacement::regex(
            r"(?P<name>[A-Za-z]+), (?P<year>\d{4})",
            "$name ($year)",
        )
        .unwrap();
        let compiled = CompiledTextReplacement::new(&replacement).unwrap();
        let result = compiled
            .transform(&["Road".into(), "map, 2026".into()])
            .unwrap();
        assert_eq!(result.output, ["Roadmap (2026)", ""]);
        assert_eq!(result.match_count, 1);
    }

    #[test]
    fn bounded_regex_expansion_matches_regex_crate_capture_syntax() {
        let template = "${word}:$2:$$:$missing:$1suffix:${1}suffix";
        let expression = Regex::new(r"(?P<word>[A-Za-z]+)-(\d+)").unwrap();
        let captures = expression.captures("Roadmap-42").unwrap();
        let mut expected = String::new();
        captures.expand(template, &mut expected);

        assert_eq!(
            expand_regex_replacement(&captures, template, 0).unwrap(),
            expected
        );

        let replacement =
            NativeOfficeTextReplacement::regex(expression.as_str(), template).unwrap();
        let result = CompiledTextReplacement::new(&replacement)
            .unwrap()
            .transform(&[String::new(), "Roadmap-42".into(), String::new()])
            .unwrap();
        assert_eq!(result.output, ["", expected.as_str(), ""]);
    }
}
