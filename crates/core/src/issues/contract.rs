use a3s_use_core::{UseError, UseResult};
use serde::{Deserialize, Serialize};

use crate::DocumentKind;

/// Default maximum number of matching native Office issues returned.
pub const DEFAULT_NATIVE_OFFICE_ISSUE_LIMIT: usize = 200;
/// Hard maximum number of matching native Office issues returned.
pub const MAX_NATIVE_OFFICE_ISSUE_LIMIT: usize = 1_000;

/// Broad issue category compatible with the Office CLI vocabulary.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum NativeOfficeIssueCategory {
    Format,
    Content,
    Structure,
}

/// Stable severity assigned by a native issue rule.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum NativeOfficeIssueSeverity {
    Error,
    Warning,
    Info,
}

/// Stable subtype emitted by one conservative native issue rule.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum NativeOfficeIssueSubtype {
    MissingAltText,
    BrokenPartRef,
    FormulaNotEvaluated,
    FormulaRefMissingSheet,
    FormulaEvalError,
    LowContrast,
}

impl NativeOfficeIssueSubtype {
    /// Return the canonical snake-case subtype name.
    pub fn as_str(self) -> &'static str {
        match self {
            Self::MissingAltText => "missing_alt_text",
            Self::BrokenPartRef => "broken_part_ref",
            Self::FormulaNotEvaluated => "formula_not_evaluated",
            Self::FormulaRefMissingSheet => "formula_ref_missing_sheet",
            Self::FormulaEvalError => "formula_eval_error",
            Self::LowContrast => "low_contrast",
        }
    }
}

/// Broad category or exact subtype accepted by a native issue scan.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "snake_case")]
pub enum NativeOfficeIssueFilter {
    Format,
    Content,
    Structure,
    MissingAltText,
    BrokenPartRef,
    FormulaNotEvaluated,
    FormulaRefMissingSheet,
    FormulaEvalError,
    LowContrast,
}

impl NativeOfficeIssueFilter {
    /// Parse a case-insensitive CLI filter, including broad single-letter aliases.
    pub fn parse(value: &str) -> UseResult<Self> {
        let normalized = value.trim().to_ascii_lowercase().replace('-', "_");
        match normalized.as_str() {
            "format" | "f" => Ok(Self::Format),
            "content" | "c" => Ok(Self::Content),
            "structure" | "s" => Ok(Self::Structure),
            "missing_alt_text" => Ok(Self::MissingAltText),
            "broken_part_ref" => Ok(Self::BrokenPartRef),
            "formula_not_evaluated" => Ok(Self::FormulaNotEvaluated),
            "formula_ref_missing_sheet" => Ok(Self::FormulaRefMissingSheet),
            "formula_eval_error" => Ok(Self::FormulaEvalError),
            "low_contrast" => Ok(Self::LowContrast),
            _ => Err(UseError::new(
                "use.office.issue_filter_invalid",
                format!("Native Office issue filter '{value}' is not supported."),
            )
            .with_suggestion(format!("Use one of: {}.", Self::valid_values().join(", ")))
            .with_detail("filter", value.to_string())
            .with_detail("validValues", serde_json::json!(Self::valid_values()))),
        }
    }

    /// Return the canonical snake-case filter name.
    pub fn as_str(self) -> &'static str {
        match self {
            Self::Format => "format",
            Self::Content => "content",
            Self::Structure => "structure",
            Self::MissingAltText => "missing_alt_text",
            Self::BrokenPartRef => "broken_part_ref",
            Self::FormulaNotEvaluated => "formula_not_evaluated",
            Self::FormulaRefMissingSheet => "formula_ref_missing_sheet",
            Self::FormulaEvalError => "formula_eval_error",
            Self::LowContrast => "low_contrast",
        }
    }

    /// Return canonical user-facing filter values; single-letter aliases are omitted.
    pub fn valid_values() -> &'static [&'static str] {
        &[
            "format",
            "content",
            "structure",
            "missing_alt_text",
            "broken_part_ref",
            "formula_not_evaluated",
            "formula_ref_missing_sheet",
            "formula_eval_error",
            "low_contrast",
        ]
    }

    pub(super) fn matches(self, issue: &NativeOfficeIssue) -> bool {
        match self {
            Self::Format => issue.category == NativeOfficeIssueCategory::Format,
            Self::Content => issue.category == NativeOfficeIssueCategory::Content,
            Self::Structure => issue.category == NativeOfficeIssueCategory::Structure,
            Self::MissingAltText => issue.subtype == NativeOfficeIssueSubtype::MissingAltText,
            Self::BrokenPartRef => issue.subtype == NativeOfficeIssueSubtype::BrokenPartRef,
            Self::FormulaNotEvaluated => {
                issue.subtype == NativeOfficeIssueSubtype::FormulaNotEvaluated
            }
            Self::FormulaRefMissingSheet => {
                issue.subtype == NativeOfficeIssueSubtype::FormulaRefMissingSheet
            }
            Self::FormulaEvalError => issue.subtype == NativeOfficeIssueSubtype::FormulaEvalError,
            Self::LowContrast => issue.subtype == NativeOfficeIssueSubtype::LowContrast,
        }
    }
}

/// Options for one bounded native Office issue scan.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficeIssueOptions {
    pub filter: Option<NativeOfficeIssueFilter>,
    pub limit: usize,
}

impl Default for NativeOfficeIssueOptions {
    fn default() -> Self {
        Self {
            filter: None,
            limit: DEFAULT_NATIVE_OFFICE_ISSUE_LIMIT,
        }
    }
}

/// One deterministic issue produced by a conservative native rule.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeIssue {
    pub id: String,
    #[serde(rename = "type")]
    pub category: NativeOfficeIssueCategory,
    pub subtype: NativeOfficeIssueSubtype,
    pub severity: NativeOfficeIssueSeverity,
    pub path: String,
    pub message: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub context: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub suggestion: Option<String>,
}

/// Bounded issue scan result; `count` includes matching records omitted by `limit`.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeIssueReport {
    pub kind: DocumentKind,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub filter: Option<NativeOfficeIssueFilter>,
    pub count: usize,
    pub returned: usize,
    pub truncated: bool,
    pub issues: Vec<NativeOfficeIssue>,
}
