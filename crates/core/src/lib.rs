//! Native OOXML operations plus a temporary OfficeCLI compatibility boundary.
//!
//! Package, OPC, XML, selector, and semantic read APIs execute in-process and
//! do not require OfficeCLI, Microsoft Office, LibreOffice, or another runtime.
//! Compatibility-only commands still invoke the pinned OfficeCLI binary during
//! migration. A3S Use does not implement OfficeCLI's private resident protocol.

use std::collections::BTreeMap;
use std::path::PathBuf;

use a3s_use_core::{UseError, UseResult, UseSessionId};
use async_trait::async_trait;
use serde::{Deserialize, Serialize};

mod command;
mod discovery;
mod editor;
mod install;
mod issues;
mod opc;
mod opc_edit;
mod package;
mod render;
mod replay;
mod semantic;
mod spreadsheet_conditional_format;
mod spreadsheet_filter;
mod spreadsheet_formula;
mod spreadsheet_named_range;
mod spreadsheet_reference;
mod spreadsheet_table;
mod template;
mod template_merge;
mod xml;
mod xml_edit;
mod xml_tree;

pub use command::{delegate_native, OfficeCliProvider};
pub use discovery::{
    discover_office_cli, doctor, office_status, OfficeInstallSource, OfficeRuntimeStatus,
    SUPPORTED_OFFICECLI_VERSION,
};
pub use editor::{
    NativeBatchResult, NativeCreatedImage, NativeCreatedPart, NativeOfficeComment,
    NativeOfficeCommentPosition, NativeOfficeCommentUpdate, NativeOfficeEditor,
    NativeOfficeHighlightColor, NativeOfficeHorizontalAlignment, NativeOfficeHyperlink,
    NativeOfficeHyperlinkTarget, NativeOfficeImage, NativeOfficeImageFormat,
    NativeOfficeImageMetadata, NativeOfficeInsertPosition, NativeOfficeMutation,
    NativeOfficePartType, NativeOfficeRgbColor, NativeOfficeSwapResult, NativeOfficeTextCase,
    NativeOfficeTextFormat, NativeOfficeTextMatchMode, NativeOfficeTextReplacement,
    NativeOfficeTextReplacementResult, NativeOfficeTextScript, NativeOfficeUnderline,
    NativeRawXmlPart, NativeSpreadsheetAutoFilter, NativeSpreadsheetBorder,
    NativeSpreadsheetBorderLine, NativeSpreadsheetBorderStyle, NativeSpreadsheetCellFormat,
    NativeSpreadsheetConditionalFormat, NativeSpreadsheetConditionalFormatIconSet,
    NativeSpreadsheetConditionalFormatOperator, NativeSpreadsheetConditionalFormatRule,
    NativeSpreadsheetConditionalFormatThreshold, NativeSpreadsheetConditionalFormatThresholdKind,
    NativeSpreadsheetConditionalFormatTimePeriod, NativeSpreadsheetDataValidation,
    NativeSpreadsheetDataValidationErrorStyle, NativeSpreadsheetDataValidationOperator,
    NativeSpreadsheetDataValidationType, NativeSpreadsheetDelimitedFormat,
    NativeSpreadsheetDelimitedImport, NativeSpreadsheetDifferentialFormat,
    NativeSpreadsheetDynamicFilter, NativeSpreadsheetFill, NativeSpreadsheetFilterColumn,
    NativeSpreadsheetFilterCriteria, NativeSpreadsheetFrozenPane, NativeSpreadsheetImportResult,
    NativeSpreadsheetNamedRange, NativeSpreadsheetNamedRangeScope, NativeSpreadsheetReadingOrder,
    NativeSpreadsheetSort, NativeSpreadsheetSortDirection, NativeSpreadsheetSortKey,
    NativeSpreadsheetTable, NativeSpreadsheetTableColumn, NativeSpreadsheetTableStyle,
    NativeSpreadsheetVerticalAlignment, SpreadsheetCellValue, MAX_NATIVE_OFFICE_FIND_BYTES,
    MAX_NATIVE_OFFICE_REPLACEMENT_BYTES, MAX_NATIVE_OFFICE_TEXT_MATCHES,
    MAX_NATIVE_OFFICE_TEXT_REPLACEMENT_OUTPUT_BYTES, MAX_NATIVE_OFFICE_TEXT_SCOPE_CELLS,
    MAX_NATIVE_SPREADSHEET_IMPORT_BYTES, MAX_NATIVE_SPREADSHEET_IMPORT_CELLS,
};
pub use install::{install_office_cli, repair_office_cli, uninstall_managed_office_cli};
pub use issues::{
    NativeOfficeIssue, NativeOfficeIssueCategory, NativeOfficeIssueFilter,
    NativeOfficeIssueOptions, NativeOfficeIssueReport, NativeOfficeIssueSeverity,
    NativeOfficeIssueSubtype, DEFAULT_NATIVE_OFFICE_ISSUE_LIMIT, MAX_NATIVE_OFFICE_ISSUE_LIMIT,
};
pub use opc::{
    ContentTypes, OpcPackageModel, Relationship, RelationshipGraph, RelationshipSource,
    RelationshipTarget,
};
pub use package::{NativeOfficePackage, PackageLimits, PackageRevision};
pub use render::{
    NativeOfficeRenderFormat, NativeOfficeRenderedView, MAX_NATIVE_OFFICE_RENDER_BYTES,
};
pub use replay::{
    NativeOfficeReplayArtifact, NativeOfficeReplayBase, MAX_NATIVE_OFFICE_REPLAY_MUTATIONS,
    NATIVE_OFFICE_REPLAY_FORMAT, NATIVE_OFFICE_REPLAY_SCHEMA_VERSION,
};
pub use semantic::{
    DocumentNode, DocumentStatistics, NativeOfficeAnnotatedEntry, NativeOfficeAnnotatedOptions,
    NativeOfficeAnnotatedView, NativeOfficeDocument, OfficeNodeType, OutlineEntry, TextBlock,
    TextView, DEFAULT_NATIVE_OFFICE_ANNOTATED_LIMIT, MAX_NATIVE_OFFICE_ANNOTATED_LIMIT,
};
pub use spreadsheet_formula::{
    parse_spreadsheet_formula, SpreadsheetFormula, SpreadsheetFormulaBinaryOperator,
    SpreadsheetFormulaCalculatedCell, SpreadsheetFormulaCalculation, SpreadsheetFormulaCell,
    SpreadsheetFormulaDependencyGraph, SpreadsheetFormulaDependencyNode,
    SpreadsheetFormulaErrorLiteral, SpreadsheetFormulaExpression, SpreadsheetFormulaExpressionKind,
    SpreadsheetFormulaFunctionDefinition, SpreadsheetFormulaFunctionRegistry,
    SpreadsheetFormulaFunctionReturnKind, SpreadsheetFormulaFunctionVolatility,
    SpreadsheetFormulaLiteral, SpreadsheetFormulaPostfixOperator, SpreadsheetFormulaQualifier,
    SpreadsheetFormulaReference, SpreadsheetFormulaReferenceKind, SpreadsheetFormulaSpan,
    SpreadsheetFormulaUnaryOperator, SpreadsheetFormulaUnresolvedReference,
    SpreadsheetFormulaUnresolvedReferenceKind, SpreadsheetFormulaValue,
    MAX_SPREADSHEET_FORMULA_CALCULATION_TEXT_BYTES, MAX_SPREADSHEET_FORMULA_CELLS,
    MAX_SPREADSHEET_FORMULA_CHARACTERS, MAX_SPREADSHEET_FORMULA_DEPENDENCIES,
    MAX_SPREADSHEET_FORMULA_DEPTH, MAX_SPREADSHEET_FORMULA_NODES,
    MAX_SPREADSHEET_FORMULA_REFERENCE_AREAS, MAX_SPREADSHEET_FORMULA_REFERENCE_VISITS,
    MAX_SPREADSHEET_FORMULA_SPILL_CELLS, MAX_SPREADSHEET_FORMULA_TEXT_BYTES,
};
pub use template_merge::{
    NativeOfficeTemplateMergeResult, MAX_TEMPLATE_DATA_DEPTH, MAX_TEMPLATE_DATA_ENTRIES,
    MAX_TEMPLATE_DATA_FLATTENED_BYTES, MAX_TEMPLATE_DATA_KEY_BYTES,
};
pub use xml::{LosslessXmlPart, XmlEncoding, XmlLimits, XmlRootName};

#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum DocumentKind {
    Word,
    Spreadsheet,
    Presentation,
}

impl DocumentKind {
    fn accepts(self, path: &std::path::Path) -> bool {
        let extension = path
            .extension()
            .and_then(|value| value.to_str())
            .map(str::to_ascii_lowercase);
        matches!(
            (self, extension.as_deref()),
            (Self::Word, Some("docx"))
                | (Self::Spreadsheet, Some("xlsx"))
                | (Self::Presentation, Some("pptx"))
        )
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct OpenDocument {
    pub path: PathBuf,
    pub kind: DocumentKind,
    pub read_only: bool,
}

impl OpenDocument {
    fn validate(&self) -> UseResult<()> {
        if self.kind.accepts(&self.path) {
            Ok(())
        } else {
            Err(UseError::new(
                "use.office.document_kind_mismatch",
                format!(
                    "Document '{}' does not match the requested {:?} kind.",
                    self.path.display(),
                    self.kind
                ),
            ))
        }
    }
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct ReadRequest {
    pub session: UseSessionId,
    pub selector: String,
}

/// One command in OfficeCLI's documented batch-item shape.
///
/// This is an Office-specific data model passed to `officecli batch`; it is not
/// an A3S RPC envelope or a cross-domain action protocol.
#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct BatchCommand {
    pub command: String,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub path: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub parent: Option<String>,
    #[serde(rename = "type", skip_serializing_if = "Option::is_none")]
    pub element_type: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub from: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub index: Option<i32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub after: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub before: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub to: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub path2: Option<String>,
    #[serde(default, skip_serializing_if = "BTreeMap::is_empty")]
    pub props: BTreeMap<String, String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub selector: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub mode: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub depth: Option<i32>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub part: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub xpath: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub action: Option<String>,
    #[serde(skip_serializing_if = "Option::is_none")]
    pub xml: Option<String>,
}

impl BatchCommand {
    pub fn new(command: impl Into<String>) -> Self {
        Self {
            command: command.into(),
            ..Self::default()
        }
    }
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct BatchRequest {
    pub session: UseSessionId,
    pub request_id: String,
    pub commands: Vec<BatchCommand>,
    pub stop_on_error: bool,
}

#[derive(Debug, Clone, PartialEq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct OperationResult {
    pub request_id: String,
    pub output: serde_json::Value,
}

#[async_trait]
pub trait OfficeProvider: Send + Sync {
    async fn open(&self, request: OpenDocument) -> UseResult<UseSessionId>;
    async fn read(&self, request: ReadRequest) -> UseResult<serde_json::Value>;
    async fn batch(&self, request: BatchRequest) -> UseResult<OperationResult>;
    async fn save(&self, session: UseSessionId) -> UseResult<()>;
    async fn close(&self, session: UseSessionId) -> UseResult<()>;
}

/// Return this error when a mutation may have reached OfficeCLI but its reply
/// was lost. Callers must inspect the document before deciding to retry.
pub fn outcome_unknown(request_id: &str) -> UseError {
    UseError::new(
        "use.office.outcome_unknown",
        "The Office mutation may have completed, but its outcome is unknown.",
    )
    .with_suggestion("Inspect or reopen the document before issuing another mutation.")
    .with_detail("requestId", request_id)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn ambiguous_mutation_has_stable_non_retryable_error() {
        let error = outcome_unknown("request-7");
        assert_eq!(error.code, "use.office.outcome_unknown");
        assert_eq!(error.details["requestId"], "request-7");
        assert!(error.suggestion.unwrap().contains("Inspect"));
    }

    #[test]
    fn provider_contract_is_send_and_sync() {
        fn assert_send_sync<T: Send + Sync + ?Sized>() {}
        assert_send_sync::<dyn OfficeProvider>();
    }

    #[test]
    fn document_kind_is_checked_before_launching_officecli() {
        let request = OpenDocument {
            path: "report.xlsx".into(),
            kind: DocumentKind::Word,
            read_only: false,
        };
        assert_eq!(
            request.validate().unwrap_err().code,
            "use.office.document_kind_mismatch"
        );
    }
}

#[cfg(test)]
mod annotated_tests;
#[cfg(test)]
mod arrange_tests;
#[cfg(test)]
mod cell_border_tests;
#[cfg(test)]
mod cell_format_tests;
#[cfg(test)]
mod cell_merge_tests;
#[cfg(test)]
mod comment_tests;
#[cfg(test)]
mod conditional_formatting_tests;
#[cfg(test)]
mod data_validation_tests;
#[cfg(test)]
mod format_tests;
#[cfg(test)]
mod hyperlink_tests;
#[cfg(test)]
mod image_tests;
#[cfg(test)]
mod issues_tests;
#[cfg(test)]
mod named_range_tests;
#[cfg(test)]
mod spreadsheet_edit_tests;

#[cfg(test)]
mod spreadsheet_filter_tests;

#[cfg(test)]
mod spreadsheet_formula_calculation_tests;

#[cfg(test)]
mod spreadsheet_formula_graph_tests;

#[cfg(test)]
mod spreadsheet_import_tests;
#[cfg(test)]
mod spreadsheet_sort_tests;

#[cfg(test)]
mod spreadsheet_table_tests;

#[cfg(test)]
mod opc_tests;

#[cfg(test)]
mod package_tests;

#[cfg(test)]
mod presentation_table_tests;

#[cfg(test)]
mod replay_tests;

#[cfg(test)]
mod render_tests;

#[cfg(test)]
mod semantic_tests;

#[cfg(test)]
mod template_merge_tests;

#[cfg(test)]
mod text_replace_tests;
