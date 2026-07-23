use std::collections::BTreeMap;

use a3s_use_core::{UseError, UseResult};
use serde::{Deserialize, Serialize};

use super::{DocumentNode, NativeOfficeDocument, OfficeNodeType};
use crate::DocumentKind;

/// Default maximum number of semantic entries returned by an annotated view.
pub const DEFAULT_NATIVE_OFFICE_ANNOTATED_LIMIT: usize = 200;
/// Hard maximum number of semantic entries returned by an annotated view.
pub const MAX_NATIVE_OFFICE_ANNOTATED_LIMIT: usize = 1_000;

const MAX_ANNOTATED_TEXT_BYTES: usize = 4 * 1_024;
const MAX_ANNOTATED_VALUE_BYTES: usize = 512;

/// Options for one bounded native Office annotated view.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficeAnnotatedOptions {
    pub limit: usize,
}

impl Default for NativeOfficeAnnotatedOptions {
    fn default() -> Self {
        Self {
            limit: DEFAULT_NATIVE_OFFICE_ANNOTATED_LIMIT,
        }
    }
}

/// One flattened semantic entry with its stable path and observed formatting.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeAnnotatedEntry {
    pub path: String,
    #[serde(rename = "type")]
    pub node_type: OfficeNodeType,
    pub level: usize,
    pub text: String,
    pub text_truncated: bool,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style: Option<String>,
    #[serde(default, skip_serializing_if = "BTreeMap::is_empty")]
    pub format: BTreeMap<String, String>,
}

/// Bounded annotated view shared by the native Rust, CLI, and MCP surfaces.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeAnnotatedView {
    pub kind: DocumentKind,
    pub limit: usize,
    pub total: usize,
    pub returned: usize,
    pub truncated: bool,
    pub text: String,
    pub entries: Vec<NativeOfficeAnnotatedEntry>,
}

impl NativeOfficeDocument {
    /// Produce a native annotated view with the default entry bound.
    pub fn annotated_view(&self) -> UseResult<NativeOfficeAnnotatedView> {
        self.annotated(NativeOfficeAnnotatedOptions::default())
    }

    /// Produce a bounded native annotated view from the existing semantic tree.
    pub fn annotated(
        &self,
        options: NativeOfficeAnnotatedOptions,
    ) -> UseResult<NativeOfficeAnnotatedView> {
        if !(1..=MAX_NATIVE_OFFICE_ANNOTATED_LIMIT).contains(&options.limit) {
            return Err(UseError::new(
                "use.office.annotated_limit_invalid",
                format!(
                    "Native Office annotated limit must be between 1 and {MAX_NATIVE_OFFICE_ANNOTATED_LIMIT}."
                ),
            )
            .with_detail("limit", options.limit)
            .with_detail("maxLimit", MAX_NATIVE_OFFICE_ANNOTATED_LIMIT));
        }

        let mut collector = AnnotatedCollector {
            kind: self.kind(),
            limit: options.limit,
            total: 0,
            entries: Vec::new(),
        };
        collector.visit(self.root(), 0);
        let truncated = collector.total > collector.entries.len();
        let mut text = collector
            .entries
            .iter()
            .map(format_entry)
            .collect::<Vec<_>>()
            .join("\n");
        if truncated {
            if !text.is_empty() {
                text.push('\n');
            }
            text.push_str(&format!(
                "... (showed {} of {} annotated entries; use --limit up to {MAX_NATIVE_OFFICE_ANNOTATED_LIMIT})",
                collector.entries.len(),
                collector.total
            ));
        }
        Ok(NativeOfficeAnnotatedView {
            kind: self.kind(),
            limit: options.limit,
            total: collector.total,
            returned: collector.entries.len(),
            truncated,
            text,
            entries: collector.entries,
        })
    }
}

struct AnnotatedCollector {
    kind: DocumentKind,
    limit: usize,
    total: usize,
    entries: Vec<NativeOfficeAnnotatedEntry>,
}

impl AnnotatedCollector {
    fn visit(&mut self, node: &DocumentNode, level: usize) {
        let include = include_node(self.kind, node.node_type);
        let child_level = if include {
            self.total += 1;
            if self.entries.len() < self.limit {
                self.entries.push(entry(node, level));
            }
            level + 1
        } else {
            level
        };
        for child in &node.children {
            self.visit(child, child_level);
        }
    }
}

fn include_node(kind: DocumentKind, node_type: OfficeNodeType) -> bool {
    match kind {
        DocumentKind::Word => matches!(
            node_type,
            OfficeNodeType::Paragraph
                | OfficeNodeType::Run
                | OfficeNodeType::Hyperlink
                | OfficeNodeType::Comment
                | OfficeNodeType::Table
                | OfficeNodeType::Picture
                | OfficeNodeType::Header
                | OfficeNodeType::Footer
        ),
        DocumentKind::Spreadsheet => matches!(
            node_type,
            OfficeNodeType::Worksheet
                | OfficeNodeType::Cell
                | OfficeNodeType::Picture
                | OfficeNodeType::Comment
        ),
        DocumentKind::Presentation => matches!(
            node_type,
            OfficeNodeType::Slide
                | OfficeNodeType::Shape
                | OfficeNodeType::Placeholder
                | OfficeNodeType::Picture
                | OfficeNodeType::Table
                | OfficeNodeType::Chart
                | OfficeNodeType::Connector
                | OfficeNodeType::Group
                | OfficeNodeType::Notes
                | OfficeNodeType::Paragraph
                | OfficeNodeType::Run
                | OfficeNodeType::Comment
        ),
    }
}

fn entry(node: &DocumentNode, level: usize) -> NativeOfficeAnnotatedEntry {
    let (text, text_truncated) = bounded_value(&node.text, MAX_ANNOTATED_TEXT_BYTES);
    let style = node
        .style
        .as_deref()
        .map(|value| bounded_value(value, MAX_ANNOTATED_VALUE_BYTES).0);
    let format = node
        .format
        .iter()
        .filter(|(key, _)| !internal_format_key(key))
        .map(|(key, value)| {
            (
                key.clone(),
                bounded_value(value, MAX_ANNOTATED_VALUE_BYTES).0,
            )
        })
        .collect();
    NativeOfficeAnnotatedEntry {
        path: node.path.clone(),
        node_type: node.node_type,
        level,
        text,
        text_truncated,
        style,
        format,
    }
}

fn internal_format_key(key: &str) -> bool {
    matches!(
        key,
        "id" | "ownerPart"
            | "part"
            | "relationshipId"
            | "linkRelationshipId"
            | "sheetId"
            | "slideId"
    )
}

fn bounded_value(value: &str, limit: usize) -> (String, bool) {
    if value.len() <= limit {
        return (value.to_string(), false);
    }
    let mut end = limit;
    while !value.is_char_boundary(end) {
        end -= 1;
    }
    let mut bounded = value[..end].to_string();
    bounded.push('…');
    (bounded, true)
}

fn format_entry(entry: &NativeOfficeAnnotatedEntry) -> String {
    let mut output = format!(
        "{}[{}] [{}]",
        "  ".repeat(entry.level),
        entry.path,
        entry.node_type.label()
    );
    if !entry.text.is_empty() {
        output.push_str(&format!(" \"{}\"", escape_line(&entry.text)));
    }
    if let Some(style) = &entry.style {
        output.push_str(&format!(" style={}", escape_line(style)));
    }
    for (key, value) in &entry.format {
        output.push_str(&format!(" {key}={}", escape_line(value)));
    }
    if entry.text_truncated {
        output.push_str(" textTruncated=true");
    }
    output
}

fn escape_line(value: &str) -> String {
    let mut output = String::with_capacity(value.len());
    for character in value.chars() {
        match character {
            '\\' => output.push_str("\\\\"),
            '\n' => output.push_str("\\n"),
            '\r' => output.push_str("\\r"),
            '\t' => output.push_str("\\t"),
            '"' => output.push_str("\\\""),
            character if character.is_control() => {
                output.push_str(&format!("\\u{{{:x}}}", u32::from(character)));
            }
            character => output.push(character),
        }
    }
    output
}
