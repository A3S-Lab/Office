use std::collections::BTreeSet;

use a3s_use_core::{UseError, UseResult};

mod color;
mod contract;
mod formula;

use color::{contrast_ratio, parse_rgb};
pub use contract::{
    NativeOfficeIssue, NativeOfficeIssueCategory, NativeOfficeIssueFilter,
    NativeOfficeIssueOptions, NativeOfficeIssueReport, NativeOfficeIssueSeverity,
    NativeOfficeIssueSubtype, DEFAULT_NATIVE_OFFICE_ISSUE_LIMIT, MAX_NATIVE_OFFICE_ISSUE_LIMIT,
};
use formula::{formula_contains_error_literal, formula_references_missing_sheet, is_formula_error};

use crate::{
    DocumentNode, NativeOfficeDocument, OfficeNodeType, RelationshipSource, RelationshipTarget,
};

impl NativeOfficeDocument {
    /// Run all default native issue rules and return at most 200 matching records.
    pub fn issue_view(&self) -> UseResult<NativeOfficeIssueReport> {
        self.issues(NativeOfficeIssueOptions::default())
    }

    /// Run native issue rules with an optional category/subtype filter and bounded output.
    pub fn issues(&self, options: NativeOfficeIssueOptions) -> UseResult<NativeOfficeIssueReport> {
        if !(1..=MAX_NATIVE_OFFICE_ISSUE_LIMIT).contains(&options.limit) {
            return Err(UseError::new(
                "use.office.issue_limit_invalid",
                format!(
                    "Native Office issue limit must be between 1 and {MAX_NATIVE_OFFICE_ISSUE_LIMIT}."
                ),
            )
            .with_detail("limit", u64::try_from(options.limit).unwrap_or(u64::MAX))
            .with_detail(
                "maxLimit",
                u64::try_from(MAX_NATIVE_OFFICE_ISSUE_LIMIT).unwrap_or(u64::MAX),
            ));
        }
        let mut scanner = IssueScanner::new(self, options);
        scanner.visit(self.root(), None);
        Ok(scanner.finish())
    }
}

struct IssueScanner<'a> {
    document: &'a NativeOfficeDocument,
    options: NativeOfficeIssueOptions,
    sheet_names: BTreeSet<String>,
    count: usize,
    issues: Vec<NativeOfficeIssue>,
}

impl<'a> IssueScanner<'a> {
    fn new(document: &'a NativeOfficeDocument, options: NativeOfficeIssueOptions) -> Self {
        let sheet_names = document
            .root()
            .children
            .iter()
            .filter(|node| node.node_type == OfficeNodeType::Worksheet)
            .filter_map(|node| node.path.strip_prefix('/'))
            .map(str::to_lowercase)
            .collect();
        Self {
            document,
            options,
            sheet_names,
            count: 0,
            issues: Vec::new(),
        }
    }

    fn finish(self) -> NativeOfficeIssueReport {
        NativeOfficeIssueReport {
            kind: self.document.kind(),
            filter: self.options.filter,
            count: self.count,
            returned: self.issues.len(),
            truncated: self.count > self.issues.len(),
            issues: self.issues,
        }
    }

    fn visit(&mut self, node: &DocumentNode, owner_part: Option<&str>) {
        let owner_part = semantic_owner_part(node).or(owner_part);
        self.check_picture(node, owner_part);
        self.check_relationships(node, owner_part);
        self.check_formula(node);
        self.check_contrast(node);
        for child in &node.children {
            self.visit(child, owner_part);
        }
    }

    fn check_picture(&mut self, node: &DocumentNode, _owner_part: Option<&str>) {
        if node.node_type != OfficeNodeType::Picture
            || node
                .format
                .get("alt")
                .is_some_and(|alt| !alt.trim().is_empty())
        {
            return;
        }
        self.emit(NativeOfficeIssue {
            id: issue_id(NativeOfficeIssueSubtype::MissingAltText, &node.path),
            category: NativeOfficeIssueCategory::Content,
            subtype: NativeOfficeIssueSubtype::MissingAltText,
            severity: NativeOfficeIssueSeverity::Warning,
            path: node.path.clone(),
            message: "Picture has no alternative text.".to_string(),
            context: node.format.get("name").map(|name| bounded_context(name)),
            suggestion: Some(
                "Add concise alternative text that describes the picture's purpose.".to_string(),
            ),
        });
    }

    fn check_relationships(&mut self, node: &DocumentNode, owner_part: Option<&str>) {
        let Some(owner_part) = owner_part else {
            return;
        };
        for &(key, expected_suffix, internal_required) in relationship_checks(node) {
            let Some(id) = node.format.get(key) else {
                continue;
            };
            let source = RelationshipSource::Part {
                part_name: owner_part.to_string(),
            };
            let relationship = self
                .document
                .opc()
                .relationships()
                .relationship(&source, id);
            let valid = relationship.is_some_and(|relationship| {
                relationship.relationship_type.ends_with(expected_suffix)
                    && (!internal_required
                        || matches!(relationship.target, RelationshipTarget::Internal { .. }))
            });
            if valid {
                continue;
            }
            let bounded_id = bounded_context(id);
            self.emit(NativeOfficeIssue {
                id: format!(
                    "{}:{}:{key}",
                    NativeOfficeIssueSubtype::BrokenPartRef.as_str(),
                    node.path
                ),
                category: NativeOfficeIssueCategory::Structure,
                subtype: NativeOfficeIssueSubtype::BrokenPartRef,
                severity: NativeOfficeIssueSeverity::Error,
                path: node.path.clone(),
                message: format!(
                    "Element references missing or incompatible relationship '{bounded_id}' from '/{owner_part}'."
                ),
                context: Some(bounded_id),
                suggestion: Some(
                    "Repair the owner relationship or remove the broken element reference."
                        .to_string(),
                ),
            });
        }
    }

    fn check_formula(&mut self, node: &DocumentNode) {
        if node.node_type != OfficeNodeType::Cell {
            return;
        }
        let Some(formula) = node.format.get("formula") else {
            return;
        };
        let (subtype, severity, message, suggestion) = if formula_references_missing_sheet(
            formula,
            &self.sheet_names,
        ) {
            (
                NativeOfficeIssueSubtype::FormulaRefMissingSheet,
                NativeOfficeIssueSeverity::Error,
                "Formula references a worksheet that does not exist.",
                "Repair the formula reference or restore the missing worksheet.",
            )
        } else if node
            .format
            .get("valueType")
            .is_some_and(|kind| kind == "Error")
            || is_formula_error(&node.text)
            || formula_contains_error_literal(formula)
        {
            (
                NativeOfficeIssueSubtype::FormulaEvalError,
                NativeOfficeIssueSeverity::Error,
                "Formula has an error result.",
                "Inspect the formula inputs and replace the invalid expression or references.",
            )
        } else if node.format.get("formulaCached").map(String::as_str) == Some("false") {
            (
                NativeOfficeIssueSubtype::FormulaNotEvaluated,
                NativeOfficeIssueSeverity::Warning,
                "Formula has no cached result and requires recalculation.",
                "Run `a3s use office native recalculate` or open the workbook in a conforming spreadsheet engine.",
            )
        } else {
            return;
        };
        self.emit(NativeOfficeIssue {
            id: issue_id(subtype, &node.path),
            category: NativeOfficeIssueCategory::Content,
            subtype,
            severity,
            path: node.path.clone(),
            message: message.to_string(),
            context: Some(bounded_context(formula)),
            suggestion: Some(suggestion.to_string()),
        });
    }

    fn check_contrast(&mut self, node: &DocumentNode) {
        if !matches!(
            node.node_type,
            OfficeNodeType::Shape | OfficeNodeType::Placeholder
        ) {
            return;
        }
        let Some(background) = node.format.get("fill").and_then(|color| parse_rgb(color)) else {
            return;
        };
        self.check_contrast_descendants(node, background);
    }

    fn check_contrast_descendants(&mut self, node: &DocumentNode, background: [u8; 3]) {
        if node.node_type == OfficeNodeType::Run && !node.text.trim().is_empty() {
            if let Some(foreground) = node.format.get("color").and_then(|color| parse_rgb(color)) {
                if contrast_ratio(background, foreground) < 4.5 {
                    self.emit(NativeOfficeIssue {
                        id: issue_id(NativeOfficeIssueSubtype::LowContrast, &node.path),
                        category: NativeOfficeIssueCategory::Format,
                        subtype: NativeOfficeIssueSubtype::LowContrast,
                        severity: NativeOfficeIssueSeverity::Warning,
                        path: node.path.clone(),
                        message: "Explicit text and shape-fill colors have low contrast."
                            .to_string(),
                        context: Some(bounded_context(&node.text)),
                        suggestion: Some(
                            "Choose text and fill colors with a WCAG contrast ratio of at least 4.5:1."
                                .to_string(),
                        ),
                    });
                }
            }
        }
        for child in &node.children {
            self.check_contrast_descendants(child, background);
        }
    }

    fn emit(&mut self, issue: NativeOfficeIssue) {
        if self
            .options
            .filter
            .is_some_and(|filter| !filter.matches(&issue))
        {
            return;
        }
        self.count = self.count.saturating_add(1);
        if self.issues.len() < self.options.limit {
            self.issues.push(issue);
        }
    }
}

fn semantic_owner_part(node: &DocumentNode) -> Option<&str> {
    if node.node_type == OfficeNodeType::Picture {
        return node
            .format
            .get("ownerPart")
            .map(|part| part.trim_start_matches('/'));
    }
    if matches!(
        node.node_type,
        OfficeNodeType::Body
            | OfficeNodeType::Header
            | OfficeNodeType::Footer
            | OfficeNodeType::Worksheet
            | OfficeNodeType::Slide
    ) {
        return node
            .format
            .get("part")
            .map(|part| part.trim_start_matches('/'));
    }
    None
}

type RelationshipCheck = (&'static str, &'static str, bool);

const PICTURE_RELATIONSHIP_CHECKS: &[RelationshipCheck] = &[
    ("relationshipId", "/image", true),
    ("linkRelationshipId", "/image", false),
];
const CHART_RELATIONSHIP_CHECKS: &[RelationshipCheck] = &[("relationshipId", "/chart", true)];
const HYPERLINK_RELATIONSHIP_CHECKS: &[RelationshipCheck] =
    &[("relationshipId", "/hyperlink", false)];

fn relationship_checks(node: &DocumentNode) -> &'static [RelationshipCheck] {
    match node.node_type {
        OfficeNodeType::Picture => PICTURE_RELATIONSHIP_CHECKS,
        OfficeNodeType::Chart => CHART_RELATIONSHIP_CHECKS,
        OfficeNodeType::Hyperlink => HYPERLINK_RELATIONSHIP_CHECKS,
        _ => &[],
    }
}

fn issue_id(subtype: NativeOfficeIssueSubtype, path: &str) -> String {
    format!("{}:{path}", subtype.as_str())
}

fn bounded_context(value: &str) -> String {
    const MAX_CHARS: usize = 160;
    let mut chars = value.chars();
    let mut context = chars.by_ref().take(MAX_CHARS).collect::<String>();
    if chars.next().is_some() {
        context.push('…');
    }
    context
}
