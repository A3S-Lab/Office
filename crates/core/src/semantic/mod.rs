mod annotated;
mod presentation;
mod selector;
mod spreadsheet;
mod word;

pub use annotated::{
    NativeOfficeAnnotatedEntry, NativeOfficeAnnotatedOptions, NativeOfficeAnnotatedView,
    DEFAULT_NATIVE_OFFICE_ANNOTATED_LIMIT, MAX_NATIVE_OFFICE_ANNOTATED_LIMIT,
};

use std::collections::BTreeMap;
use std::path::Path;

use a3s_use_core::{UseError, UseResult};
use serde::{Deserialize, Serialize};

use crate::discovery::office_error;
use crate::{DocumentKind, NativeOfficePackage, OpcPackageModel};

/// Stable semantic node kinds shared by native Office read APIs.
#[derive(Debug, Clone, Copy, PartialEq, Eq, PartialOrd, Ord, Hash, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum OfficeNodeType {
    Document,
    Body,
    Paragraph,
    Run,
    Hyperlink,
    Comment,
    Table,
    TableRow,
    TableColumn,
    TableCell,
    Header,
    Footer,
    Workbook,
    Worksheet,
    Row,
    Column,
    Range,
    Cell,
    ConditionalFormatting,
    DataValidation,
    NamedRangeCollection,
    NamedRange,
    AutoFilter,
    FilterColumn,
    FilterValue,
    SortState,
    SortKey,
    FrozenPane,
    Presentation,
    Slide,
    Shape,
    Placeholder,
    Picture,
    Chart,
    Connector,
    Group,
    Notes,
}

impl OfficeNodeType {
    pub fn label(self) -> &'static str {
        match self {
            Self::Document => "Document",
            Self::Body => "Body",
            Self::Paragraph => "Paragraph",
            Self::Run => "Run",
            Self::Hyperlink => "Hyperlink",
            Self::Comment => "Comment",
            Self::Table => "Table",
            Self::TableRow => "TableRow",
            Self::TableColumn => "TableColumn",
            Self::TableCell => "TableCell",
            Self::Header => "Header",
            Self::Footer => "Footer",
            Self::Workbook => "Workbook",
            Self::Worksheet => "Worksheet",
            Self::Row => "Row",
            Self::Column => "Column",
            Self::Range => "Range",
            Self::Cell => "Cell",
            Self::ConditionalFormatting => "ConditionalFormatting",
            Self::DataValidation => "DataValidation",
            Self::NamedRangeCollection => "NamedRangeCollection",
            Self::NamedRange => "NamedRange",
            Self::AutoFilter => "AutoFilter",
            Self::FilterColumn => "FilterColumn",
            Self::FilterValue => "FilterValue",
            Self::SortState => "SortState",
            Self::SortKey => "SortKey",
            Self::FrozenPane => "FrozenPane",
            Self::Presentation => "Presentation",
            Self::Slide => "Slide",
            Self::Shape => "Shape",
            Self::Placeholder => "Placeholder",
            Self::Picture => "Picture",
            Self::Chart => "Chart",
            Self::Connector => "Connector",
            Self::Group => "Group",
            Self::Notes => "Notes",
        }
    }
}

/// Format-neutral document node returned by native `get` and `query`.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DocumentNode {
    pub path: String,
    pub tag: String,
    #[serde(rename = "type")]
    pub node_type: OfficeNodeType,
    pub text: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub preview: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub style: Option<String>,
    pub child_count: usize,
    #[serde(default, skip_serializing_if = "BTreeMap::is_empty")]
    pub format: BTreeMap<String, String>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub children: Vec<DocumentNode>,
}

impl DocumentNode {
    pub(crate) fn new(
        path: impl Into<String>,
        tag: impl Into<String>,
        node_type: OfficeNodeType,
    ) -> Self {
        Self {
            path: path.into(),
            tag: tag.into(),
            node_type,
            text: String::new(),
            preview: None,
            style: None,
            child_count: 0,
            format: BTreeMap::new(),
            children: Vec::new(),
        }
    }

    pub(crate) fn normalize(&mut self) {
        for child in &mut self.children {
            child.normalize();
        }
        self.child_count = self.children.len();
    }

    fn clone_to_depth(&self, depth: usize) -> Self {
        let mut node = self.clone();
        if depth == 0 {
            node.children.clear();
        } else {
            node.children = self
                .children
                .iter()
                .map(|child| child.clone_to_depth(depth - 1))
                .collect();
        }
        node
    }

    fn find(&self, path: &str) -> Option<&Self> {
        if self.path == path {
            return Some(self);
        }
        self.children.iter().find_map(|child| child.find(path))
    }
}

/// One stable path/text record in a native text view.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextBlock {
    pub path: String,
    pub text: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct TextView {
    pub kind: DocumentKind,
    pub text: String,
    pub blocks: Vec<TextBlock>,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct OutlineEntry {
    pub path: String,
    #[serde(rename = "type")]
    pub node_type: OfficeNodeType,
    pub level: usize,
    pub text: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct DocumentStatistics {
    pub kind: DocumentKind,
    pub node_count: usize,
    pub character_count: usize,
    pub word_count: usize,
    pub paragraph_count: usize,
    pub table_count: usize,
    pub row_count: usize,
    pub cell_count: usize,
    pub sheet_count: usize,
    pub slide_count: usize,
    pub shape_count: usize,
    pub picture_count: usize,
    pub formula_count: usize,
    pub by_type: BTreeMap<String, usize>,
}

/// Read-only native semantic document backed by a validated OOXML package.
#[derive(Debug, Clone)]
pub struct NativeOfficeDocument {
    package: NativeOfficePackage,
    opc: OpcPackageModel,
    root: DocumentNode,
}

impl NativeOfficeDocument {
    pub async fn open(path: impl AsRef<Path>) -> UseResult<Self> {
        let package = NativeOfficePackage::open(path).await?;
        Self::from_package(package)
    }

    pub fn from_package(package: NativeOfficePackage) -> UseResult<Self> {
        let opc = package.opc_model()?;
        let mut root = match package.kind() {
            DocumentKind::Word => word::read(&package, &opc)?,
            DocumentKind::Spreadsheet => spreadsheet::read(&package, &opc)?,
            DocumentKind::Presentation => presentation::read(&package, &opc)?,
        };
        root.normalize();
        Ok(Self { package, opc, root })
    }

    pub fn kind(&self) -> DocumentKind {
        self.package.kind()
    }

    pub fn package(&self) -> &NativeOfficePackage {
        &self.package
    }

    pub fn opc(&self) -> &OpcPackageModel {
        &self.opc
    }

    pub fn root(&self) -> &DocumentNode {
        &self.root
    }

    pub fn get(&self, path: &str, depth: usize) -> UseResult<DocumentNode> {
        validate_path(path)?;
        if let Some(node) = self
            .root
            .find(path)
            .or_else(|| find_case_insensitive(&self.root, path, self.kind()))
        {
            return Ok(node.clone_to_depth(depth));
        }
        let virtual_node = match self.kind() {
            DocumentKind::Spreadsheet => spreadsheet::virtual_get(&self.root, path, depth)?,
            DocumentKind::Presentation => presentation::virtual_get(&self.root, path, depth)?,
            DocumentKind::Word => None,
        };
        virtual_node.ok_or_else(|| {
            semantic_error(
                "use.office.node_not_found",
                format!("Office semantic path '{path}' does not exist."),
            )
            .with_detail("path", path)
        })
    }

    pub fn query(&self, expression: &str) -> UseResult<Vec<DocumentNode>> {
        selector::query(&self.root, expression).map(|nodes| {
            nodes
                .into_iter()
                .map(|node| node.clone_to_depth(0))
                .collect()
        })
    }

    pub fn text_view(&self) -> TextView {
        let mut blocks = Vec::new();
        collect_text_blocks(&self.root, self.kind(), &mut blocks);
        let text = match self.kind() {
            DocumentKind::Spreadsheet => blocks
                .iter()
                .map(|block| format!("{}={}", block.path, block.text))
                .collect::<Vec<_>>()
                .join("\n"),
            _ => blocks
                .iter()
                .map(|block| block.text.as_str())
                .collect::<Vec<_>>()
                .join("\n"),
        };
        TextView {
            kind: self.kind(),
            text,
            blocks,
        }
    }

    pub fn outline(&self) -> Vec<OutlineEntry> {
        let mut entries = Vec::new();
        collect_outline(&self.root, self.kind(), 0, &mut entries);
        entries
    }

    pub fn statistics(&self) -> DocumentStatistics {
        let view = self.text_view();
        let mut statistics = DocumentStatistics {
            kind: self.kind(),
            node_count: 0,
            character_count: view.text.chars().count(),
            word_count: view
                .text
                .split_whitespace()
                .filter(|word| !word.is_empty())
                .count(),
            paragraph_count: 0,
            table_count: 0,
            row_count: 0,
            cell_count: 0,
            sheet_count: 0,
            slide_count: 0,
            shape_count: 0,
            picture_count: 0,
            formula_count: 0,
            by_type: BTreeMap::new(),
        };
        collect_statistics(&self.root, &mut statistics);
        statistics
    }
}

fn collect_text_blocks(node: &DocumentNode, kind: DocumentKind, output: &mut Vec<TextBlock>) {
    let include = match kind {
        DocumentKind::Word => matches!(
            node.node_type,
            OfficeNodeType::Paragraph | OfficeNodeType::Table
        ),
        DocumentKind::Spreadsheet => node.node_type == OfficeNodeType::Cell,
        DocumentKind::Presentation => node.node_type == OfficeNodeType::Slide,
    };
    if include && (!node.text.is_empty() || kind == DocumentKind::Spreadsheet) {
        output.push(TextBlock {
            path: node.path.clone(),
            text: node.text.clone(),
        });
        if kind == DocumentKind::Presentation
            || (kind == DocumentKind::Word && node.node_type == OfficeNodeType::Table)
        {
            return;
        }
    }
    for child in &node.children {
        collect_text_blocks(child, kind, output);
    }
}

fn collect_outline(
    node: &DocumentNode,
    kind: DocumentKind,
    level: usize,
    output: &mut Vec<OutlineEntry>,
) {
    let include = match kind {
        DocumentKind::Word => {
            matches!(
                node.node_type,
                OfficeNodeType::Body
                    | OfficeNodeType::Paragraph
                    | OfficeNodeType::Table
                    | OfficeNodeType::TableRow
                    | OfficeNodeType::TableCell
            )
        }
        DocumentKind::Spreadsheet => matches!(
            node.node_type,
            OfficeNodeType::Worksheet | OfficeNodeType::Row
        ),
        DocumentKind::Presentation => matches!(
            node.node_type,
            OfficeNodeType::Slide
                | OfficeNodeType::Shape
                | OfficeNodeType::Placeholder
                | OfficeNodeType::Picture
                | OfficeNodeType::Table
                | OfficeNodeType::Chart
        ),
    };
    let next_level = if include {
        output.push(OutlineEntry {
            path: node.path.clone(),
            node_type: node.node_type,
            level,
            text: node.text.clone(),
        });
        level + 1
    } else {
        level
    };
    for child in &node.children {
        collect_outline(child, kind, next_level, output);
    }
}

fn collect_statistics(node: &DocumentNode, statistics: &mut DocumentStatistics) {
    statistics.node_count += 1;
    *statistics
        .by_type
        .entry(node.node_type.label().to_string())
        .or_default() += 1;
    match node.node_type {
        OfficeNodeType::Paragraph => statistics.paragraph_count += 1,
        OfficeNodeType::Table => statistics.table_count += 1,
        OfficeNodeType::TableRow | OfficeNodeType::Row => statistics.row_count += 1,
        OfficeNodeType::TableCell | OfficeNodeType::Cell => statistics.cell_count += 1,
        OfficeNodeType::Worksheet => statistics.sheet_count += 1,
        OfficeNodeType::Slide => statistics.slide_count += 1,
        OfficeNodeType::Shape | OfficeNodeType::Placeholder => statistics.shape_count += 1,
        OfficeNodeType::Picture => statistics.picture_count += 1,
        _ => {}
    }
    if node
        .format
        .get("formula")
        .is_some_and(|formula| !formula.is_empty())
    {
        statistics.formula_count += 1;
    }
    for child in &node.children {
        collect_statistics(child, statistics);
    }
}

fn validate_path(path: &str) -> UseResult<()> {
    if path.is_empty()
        || !path.starts_with('/')
        || path.len() > 4_096
        || path.contains('\\')
        || path.chars().any(char::is_control)
        || path.split('/').any(|segment| matches!(segment, "." | ".."))
    {
        return Err(semantic_error(
            "use.office.path_invalid",
            "Office semantic paths must be absolute, bounded, and traversal-free.",
        ));
    }
    Ok(())
}

fn find_case_insensitive<'a>(
    node: &'a DocumentNode,
    path: &str,
    kind: DocumentKind,
) -> Option<&'a DocumentNode> {
    if kind == DocumentKind::Spreadsheet && node.path.eq_ignore_ascii_case(path) {
        return Some(node);
    }
    node.children
        .iter()
        .find_map(|child| find_case_insensitive(child, path, kind))
}

pub(crate) fn semantic_error(code: &str, message: impl Into<String>) -> UseError {
    office_error(code, message)
}
