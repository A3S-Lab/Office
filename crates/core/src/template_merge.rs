use std::collections::{BTreeMap, BTreeSet};

use a3s_use_core::{UseError, UseResult};
use serde::{Deserialize, Serialize};
use serde_json::Value;

use crate::discovery::office_error;
use crate::xml_edit::{
    apply_patches, decoded_element_text, index_xml, replace_element_text_patch, IndexedXmlElement,
    XmlPatch,
};
use crate::{DocumentKind, LosslessXmlPart, NativeOfficePackage};

mod spreadsheet;

pub const MAX_TEMPLATE_DATA_ENTRIES: usize = 10_000;
pub const MAX_TEMPLATE_DATA_KEY_BYTES: usize = 1_024;
pub const MAX_TEMPLATE_DATA_FLATTENED_BYTES: usize = 16 * 1024 * 1024;
pub const MAX_TEMPLATE_DATA_DEPTH: usize = 64;

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeTemplateMergeResult {
    pub replaced_count: usize,
    pub used_keys: Vec<String>,
    pub unresolved_placeholders: Vec<String>,
    pub changed_parts: Vec<String>,
}

pub(crate) fn merge(
    package: &mut NativeOfficePackage,
    data: &Value,
) -> UseResult<NativeOfficeTemplateMergeResult> {
    let data = MergeData::from_json(data)?;
    let mut accumulator = MergeAccumulator::default();
    match package.kind() {
        DocumentKind::Word => {
            let parts = word_text_parts(package);
            merge_paragraph_parts(package, parts, &data, &mut accumulator)?
        }
        DocumentKind::Spreadsheet => spreadsheet::merge(package, &data, &mut accumulator)?,
        DocumentKind::Presentation => {
            let parts = presentation_text_parts(package);
            merge_paragraph_parts(package, parts, &data, &mut accumulator)?
        }
    }

    let mut unresolved = BTreeSet::new();
    match package.kind() {
        DocumentKind::Word => {
            scan_paragraph_parts(package, word_text_parts(package), &mut unresolved)?
        }
        DocumentKind::Spreadsheet => spreadsheet::scan(package, &mut unresolved)?,
        DocumentKind::Presentation => {
            scan_paragraph_parts(package, presentation_text_parts(package), &mut unresolved)?
        }
    }

    Ok(NativeOfficeTemplateMergeResult {
        replaced_count: accumulator.replaced_count,
        used_keys: accumulator.used_keys.into_iter().collect(),
        unresolved_placeholders: unresolved.into_iter().collect(),
        changed_parts: accumulator.changed_parts.into_iter().collect(),
    })
}

#[derive(Debug)]
pub(super) struct MergeData {
    values: BTreeMap<String, String>,
    flattened_bytes: usize,
}

impl MergeData {
    fn from_json(data: &Value) -> UseResult<Self> {
        let object = data.as_object().ok_or_else(|| {
            merge_error(
                "use.office.template_data_invalid",
                "Native Office template data must be a JSON object.",
            )
        })?;
        let mut output = Self {
            values: BTreeMap::new(),
            flattened_bytes: 0,
        };
        for (key, value) in object {
            output.insert(key.clone(), json_text(value)?)?;
        }
        for (key, value) in object {
            match value {
                Value::Object(nested) => output.flatten_object(nested, key, 1)?,
                Value::Array(array) => output.flatten_array(array, key, 1)?,
                _ => {}
            }
        }
        Ok(output)
    }

    pub(super) fn get(&self, key: &str) -> Option<&str> {
        self.values.get(key).map(String::as_str)
    }

    fn flatten_object(
        &mut self,
        object: &serde_json::Map<String, Value>,
        prefix: &str,
        depth: usize,
    ) -> UseResult<()> {
        self.check_depth(depth)?;
        for (key, value) in object {
            let path = format!("{prefix}.{key}");
            match value {
                Value::Object(child) => {
                    self.flatten_object(child, &path, depth + 1)?;
                    self.insert_if_absent(path, json_text(value)?)?;
                }
                Value::Array(array) => self.flatten_array(array, &path, depth + 1)?,
                _ => self.insert_if_absent(path, json_text(value)?)?,
            }
        }
        Ok(())
    }

    fn flatten_array(&mut self, array: &[Value], prefix: &str, depth: usize) -> UseResult<()> {
        self.check_depth(depth)?;
        for (index, value) in array.iter().enumerate() {
            let path = format!("{prefix}[{index}]");
            match value {
                Value::Object(child) => {
                    self.flatten_object(child, &path, depth + 1)?;
                    self.insert_if_absent(path, json_text(value)?)?;
                }
                Value::Array(child) => self.flatten_array(child, &path, depth + 1)?,
                _ => self.insert_if_absent(path, json_text(value)?)?,
            }
        }
        Ok(())
    }

    fn insert(&mut self, key: String, value: String) -> UseResult<()> {
        if let Some(previous) = self.values.remove(&key) {
            self.flattened_bytes = self
                .flattened_bytes
                .saturating_sub(key.len().saturating_add(previous.len()));
        }
        self.insert_new(key, value)
    }

    fn insert_if_absent(&mut self, key: String, value: String) -> UseResult<()> {
        if self.values.contains_key(&key) {
            return Ok(());
        }
        self.insert_new(key, value)
    }

    fn insert_new(&mut self, key: String, value: String) -> UseResult<()> {
        if key.len() > MAX_TEMPLATE_DATA_KEY_BYTES {
            return Err(merge_error(
                "use.office.template_data_limit",
                format!(
                    "Native Office template key exceeds the {MAX_TEMPLATE_DATA_KEY_BYTES}-byte limit."
                ),
            )
            .with_detail("keyBytes", key.len()));
        }
        if self.values.len() >= MAX_TEMPLATE_DATA_ENTRIES {
            return Err(merge_error(
                "use.office.template_data_limit",
                format!(
                    "Native Office template data exceeds the {MAX_TEMPLATE_DATA_ENTRIES}-entry flattened limit."
                ),
            ));
        }
        let entry_bytes = key.len().checked_add(value.len()).ok_or_else(|| {
            merge_error(
                "use.office.template_data_limit",
                "Native Office template data size overflowed.",
            )
        })?;
        self.flattened_bytes = self
            .flattened_bytes
            .checked_add(entry_bytes)
            .ok_or_else(|| {
                merge_error(
                    "use.office.template_data_limit",
                    "Native Office template data size overflowed.",
                )
            })?;
        if self.flattened_bytes > MAX_TEMPLATE_DATA_FLATTENED_BYTES {
            return Err(merge_error(
                "use.office.template_data_limit",
                format!(
                    "Native Office flattened template data exceeds the {MAX_TEMPLATE_DATA_FLATTENED_BYTES}-byte limit."
                ),
            )
            .with_detail("flattenedBytes", self.flattened_bytes));
        }
        self.values.insert(key, value);
        Ok(())
    }

    fn check_depth(&self, depth: usize) -> UseResult<()> {
        if depth > MAX_TEMPLATE_DATA_DEPTH {
            return Err(merge_error(
                "use.office.template_data_limit",
                format!(
                    "Native Office template data exceeds the {MAX_TEMPLATE_DATA_DEPTH}-level nesting limit."
                ),
            ));
        }
        Ok(())
    }
}

#[derive(Debug, Default)]
pub(super) struct MergeAccumulator {
    replaced_count: usize,
    used_keys: BTreeSet<String>,
    changed_parts: BTreeSet<String>,
}

impl MergeAccumulator {
    fn record(&mut self, key: &str, multiplier: usize) -> UseResult<()> {
        self.replaced_count = self.replaced_count.checked_add(multiplier).ok_or_else(|| {
            merge_error(
                "use.office.template_replacement_limit",
                "Native Office template replacement count overflowed.",
            )
        })?;
        self.used_keys.insert(key.to_string());
        Ok(())
    }

    pub(super) fn changed(&mut self, part: &str) {
        self.changed_parts.insert(format!("/{part}"));
    }
}

fn merge_paragraph_parts(
    package: &mut NativeOfficePackage,
    parts: Vec<String>,
    data: &MergeData,
    accumulator: &mut MergeAccumulator,
) -> UseResult<()> {
    for part_name in parts {
        let part = package.xml_part(&part_name)?;
        let index = index_xml(&part)?;
        let mut paragraphs = Vec::new();
        collect_elements(&index, "p", &mut paragraphs);
        let mut patches = Vec::new();
        for paragraph in paragraphs {
            let mut text_elements = Vec::new();
            collect_paragraph_text(paragraph, &mut text_elements);
            merge_text_elements(&part, &text_elements, data, 1, accumulator, &mut patches)?;
        }
        if !patches.is_empty() {
            package.set_part(&part_name, apply_patches(&part, patches)?)?;
            accumulator.changed(&part_name);
        }
    }
    Ok(())
}

fn scan_paragraph_parts(
    package: &NativeOfficePackage,
    parts: Vec<String>,
    unresolved: &mut BTreeSet<String>,
) -> UseResult<()> {
    for part_name in parts {
        let part = package.xml_part(&part_name)?;
        let index = index_xml(&part)?;
        let mut paragraphs = Vec::new();
        collect_elements(&index, "p", &mut paragraphs);
        for paragraph in paragraphs {
            let mut text_elements = Vec::new();
            collect_paragraph_text(paragraph, &mut text_elements);
            scan_text_elements(&part, &text_elements, unresolved)?;
        }
    }
    Ok(())
}

pub(super) fn merge_text_elements(
    part: &LosslessXmlPart,
    elements: &[&IndexedXmlElement],
    data: &MergeData,
    multiplier: usize,
    accumulator: &mut MergeAccumulator,
    patches: &mut Vec<XmlPatch>,
) -> UseResult<()> {
    if elements.is_empty() {
        return Ok(());
    }
    let source = elements
        .iter()
        .map(|element| decoded_element_text(part, element))
        .collect::<UseResult<Vec<_>>>()?;
    let output = replace_segments(&source, data, multiplier, accumulator)?;
    for ((element, before), after) in elements.iter().zip(&source).zip(output) {
        if before != &after {
            patches.push(replace_element_text_patch(element, &after));
        }
    }
    Ok(())
}

pub(super) fn scan_text_elements(
    part: &LosslessXmlPart,
    elements: &[&IndexedXmlElement],
    unresolved: &mut BTreeSet<String>,
) -> UseResult<()> {
    let text = elements
        .iter()
        .map(|element| decoded_element_text(part, element))
        .collect::<UseResult<Vec<_>>>()?
        .concat();
    for placeholder in placeholders(&text) {
        unresolved.insert(placeholder.key);
    }
    Ok(())
}

fn replace_segments(
    segments: &[String],
    data: &MergeData,
    multiplier: usize,
    accumulator: &mut MergeAccumulator,
) -> UseResult<Vec<String>> {
    let mut characters = Vec::new();
    let mut owners = Vec::new();
    for (owner, segment) in segments.iter().enumerate() {
        for character in segment.chars() {
            characters.push(character);
            owners.push(owner);
        }
    }
    let text = characters.iter().collect::<String>();
    let placeholders = placeholders(&text);
    if placeholders.is_empty() {
        return Ok(segments.to_vec());
    }

    let mut output = vec![String::new(); segments.len()];
    let mut cursor = 0_usize;
    for placeholder in placeholders {
        append_original(&characters, &owners, cursor, placeholder.start, &mut output);
        if let Some(replacement) = data.get(&placeholder.key) {
            validate_xml_text(&placeholder.key, replacement)?;
            let owner = owners.get(placeholder.start).copied().ok_or_else(|| {
                merge_error(
                    "use.office.template_placeholder_invalid",
                    "Native Office placeholder has no source text run.",
                )
            })?;
            output[owner].push_str(replacement);
            accumulator.record(&placeholder.key, multiplier)?;
        } else {
            append_original(
                &characters,
                &owners,
                placeholder.start,
                placeholder.end,
                &mut output,
            );
        }
        cursor = placeholder.end;
    }
    append_original(&characters, &owners, cursor, characters.len(), &mut output);
    Ok(output)
}

fn validate_xml_text(key: &str, value: &str) -> UseResult<()> {
    if let Some(character) = value.chars().find(|character| {
        !matches!(*character, '\u{9}' | '\u{a}' | '\u{d}')
            && (*character < '\u{20}' || matches!(*character, '\u{fffe}' | '\u{ffff}'))
    }) {
        return Err(merge_error(
            "use.office.template_value_invalid",
            format!(
                "Native Office template value for key '{key}' contains XML-forbidden character U+{:04X}.",
                u32::from(character)
            ),
        )
        .with_detail("key", key)
        .with_detail("codePoint", format!("U+{:04X}", u32::from(character))));
    }
    Ok(())
}

fn append_original(
    characters: &[char],
    owners: &[usize],
    start: usize,
    end: usize,
    output: &mut [String],
) {
    for index in start..end {
        if let (Some(character), Some(owner)) = (characters.get(index), owners.get(index)) {
            if let Some(segment) = output.get_mut(*owner) {
                segment.push(*character);
            }
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct Placeholder {
    start: usize,
    end: usize,
    key: String,
}

fn placeholders(text: &str) -> Vec<Placeholder> {
    let characters = text.chars().collect::<Vec<_>>();
    let mut output = Vec::new();
    let mut cursor = 0_usize;
    while cursor + 1 < characters.len() {
        if characters[cursor] != '{' || characters[cursor + 1] != '{' {
            cursor += 1;
            continue;
        }
        let mut close = cursor + 2;
        while close + 1 < characters.len()
            && (characters[close] != '}' || characters[close + 1] != '}')
        {
            close += 1;
        }
        if close + 1 >= characters.len() {
            break;
        }
        let inner = characters[cursor + 2..close].iter().collect::<String>();
        let key = inner.trim();
        if valid_placeholder_key(key) {
            output.push(Placeholder {
                start: cursor,
                end: close + 2,
                key: key.to_string(),
            });
            cursor = close + 2;
        } else {
            cursor += 2;
        }
    }
    output
}

fn valid_placeholder_key(key: &str) -> bool {
    if key.is_empty() || key.len() > MAX_TEMPLATE_DATA_KEY_BYTES {
        return false;
    }
    let mut characters = key.chars();
    let Some(first) = characters.next() else {
        return false;
    };
    (first.is_alphanumeric() || first == '_')
        && characters.all(|character| {
            character.is_alphanumeric() || matches!(character, '_' | '.' | '-' | '[' | ']' | ' ')
        })
}

fn collect_elements<'a>(
    element: &'a IndexedXmlElement,
    local_name: &str,
    output: &mut Vec<&'a IndexedXmlElement>,
) {
    if element.local_name == local_name {
        output.push(element);
    }
    for child in &element.children {
        collect_elements(child, local_name, output);
    }
}

fn collect_paragraph_text<'a>(
    paragraph: &'a IndexedXmlElement,
    output: &mut Vec<&'a IndexedXmlElement>,
) {
    for child in &paragraph.children {
        if child.local_name == "p" {
            continue;
        }
        if child.local_name == "t" {
            output.push(child);
        } else {
            collect_paragraph_text(child, output);
        }
    }
}

fn word_text_parts(package: &NativeOfficePackage) -> Vec<String> {
    package
        .part_names()
        .filter(|name| {
            *name == "word/document.xml"
                || (*name == "word/footnotes.xml")
                || (*name == "word/endnotes.xml")
                || (*name == "word/comments.xml")
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

fn json_text(value: &Value) -> UseResult<String> {
    match value {
        Value::Null => Ok(String::new()),
        Value::String(value) => Ok(value.clone()),
        _ => serde_json::to_string(value).map_err(|error| {
            merge_error(
                "use.office.template_data_invalid",
                format!("Failed to serialize native Office template data: {error}"),
            )
        }),
    }
}

pub(super) fn merge_error(code: &str, message: impl Into<String>) -> UseError {
    office_error(code, message)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn data_flattening_preserves_literal_precedence_and_nested_arrays() {
        let data = MergeData::from_json(&serde_json::json!({
            "user.name": "literal",
            "user": {"name": "nested"},
            "items": [{"name": "first"}, [1, 2]],
            "empty": null
        }))
        .unwrap();
        assert_eq!(data.get("user.name"), Some("literal"));
        assert_eq!(data.get("items[0].name"), Some("first"));
        assert_eq!(data.get("items[1][1]"), Some("2"));
        assert_eq!(data.get("empty"), Some(""));
    }

    #[test]
    fn replacement_is_single_pass_and_preserves_run_ownership() {
        let data = MergeData::from_json(&serde_json::json!({
            "user.name": "Alice",
            "literal": "{{user.name}}"
        }))
        .unwrap();
        let mut accumulator = MergeAccumulator::default();
        let output = replace_segments(
            &["Hi {{user.".into(), "name}} / {{literal}}!".into()],
            &data,
            1,
            &mut accumulator,
        )
        .unwrap();
        assert_eq!(output, ["Hi Alice", " / {{user.name}}!"]);
        assert_eq!(accumulator.replaced_count, 2);
        assert_eq!(
            accumulator.used_keys,
            BTreeSet::from(["literal".into(), "user.name".into()])
        );
    }

    #[test]
    fn placeholder_grammar_accepts_documented_keys_and_rejects_markup() {
        let found =
            placeholders("{{ name }} {{items[0].first name}} {{user-id}} {{bad/key}} {{<xml>}}");
        assert_eq!(
            found.into_iter().map(|item| item.key).collect::<Vec<_>>(),
            ["name", "items[0].first name", "user-id"]
        );
    }
}
