use a3s_use_core::UseResult;
use base64::Engine as _;
use serde::{Deserialize, Serialize};
use url::Url;

use super::part::{NativeCreatedPart, NativeOfficePartType};
use crate::spreadsheet_formula::SpreadsheetFormulaCalculation;

mod conditional_formatting;
mod data_validation;
mod formatting;
mod named_range;
mod spreadsheet_filter;
mod spreadsheet_import;
mod spreadsheet_sort;
mod spreadsheet_table;
mod spreadsheet_view;

pub use conditional_formatting::{
    NativeSpreadsheetConditionalFormat, NativeSpreadsheetConditionalFormatIconSet,
    NativeSpreadsheetConditionalFormatOperator, NativeSpreadsheetConditionalFormatRule,
    NativeSpreadsheetConditionalFormatThreshold, NativeSpreadsheetConditionalFormatThresholdKind,
    NativeSpreadsheetConditionalFormatTimePeriod, NativeSpreadsheetDifferentialFormat,
};
pub use data_validation::{
    NativeSpreadsheetDataValidation, NativeSpreadsheetDataValidationErrorStyle,
    NativeSpreadsheetDataValidationOperator, NativeSpreadsheetDataValidationType,
};
pub use formatting::{
    NativeOfficeHighlightColor, NativeOfficeHorizontalAlignment, NativeOfficeRgbColor,
    NativeOfficeTextCase, NativeOfficeTextFormat, NativeOfficeTextScript, NativeOfficeUnderline,
    NativeSpreadsheetBorder, NativeSpreadsheetBorderLine, NativeSpreadsheetBorderStyle,
    NativeSpreadsheetCellFormat, NativeSpreadsheetFill, NativeSpreadsheetReadingOrder,
    NativeSpreadsheetVerticalAlignment,
};
pub use named_range::{NativeSpreadsheetNamedRange, NativeSpreadsheetNamedRangeScope};
pub use spreadsheet_filter::{
    NativeSpreadsheetAutoFilter, NativeSpreadsheetDynamicFilter, NativeSpreadsheetFilterColumn,
    NativeSpreadsheetFilterCriteria,
};
pub use spreadsheet_import::{
    NativeSpreadsheetDelimitedFormat, NativeSpreadsheetDelimitedImport,
    NativeSpreadsheetImportResult, MAX_NATIVE_SPREADSHEET_IMPORT_BYTES,
    MAX_NATIVE_SPREADSHEET_IMPORT_CELLS,
};
pub use spreadsheet_sort::{
    NativeSpreadsheetSort, NativeSpreadsheetSortDirection, NativeSpreadsheetSortKey,
};
pub use spreadsheet_table::{
    NativeSpreadsheetTable, NativeSpreadsheetTableColumn, NativeSpreadsheetTableStyle,
};
pub use spreadsheet_view::NativeSpreadsheetFrozenPane;

/// A hyperlink destination represented without executing or resolving it.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "kebab-case", deny_unknown_fields)]
pub enum NativeOfficeHyperlinkTarget {
    /// An inert absolute HTTP, HTTPS, or mail address stored as an external
    /// OOXML relationship.
    External { uri: String },
    /// A format-specific in-document location such as a Word bookmark,
    /// Spreadsheet cell location, or Presentation slide path.
    Internal { location: String },
}

/// A complete typed hyperlink value shared by native Office formats.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficeHyperlink {
    pub target: NativeOfficeHyperlinkTarget,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub display: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub tooltip: Option<String>,
}

impl NativeOfficeHyperlink {
    /// Creates a validated inert external hyperlink.
    pub fn external(uri: impl Into<String>) -> UseResult<Self> {
        let hyperlink = Self {
            target: NativeOfficeHyperlinkTarget::External { uri: uri.into() },
            display: None,
            tooltip: None,
        };
        hyperlink.validate()?;
        Ok(hyperlink)
    }

    /// Creates a validated format-specific internal hyperlink target.
    pub fn internal(location: impl Into<String>) -> UseResult<Self> {
        let hyperlink = Self {
            target: NativeOfficeHyperlinkTarget::Internal {
                location: location.into(),
            },
            display: None,
            tooltip: None,
        };
        hyperlink.validate()?;
        Ok(hyperlink)
    }

    pub fn with_display(mut self, display: impl Into<String>) -> Self {
        self.display = Some(display.into());
        self
    }

    pub fn with_tooltip(mut self, tooltip: impl Into<String>) -> Self {
        self.tooltip = Some(tooltip.into());
        self
    }

    pub(crate) fn validate(&self) -> UseResult<()> {
        match &self.target {
            NativeOfficeHyperlinkTarget::External { uri } => validate_external_uri(uri)?,
            NativeOfficeHyperlinkTarget::Internal { location } => {
                validate_hyperlink_text(
                    location,
                    2_048,
                    "use.office.hyperlink_location_invalid",
                    "Native Office internal hyperlink locations",
                )?;
            }
        }
        if let Some(display) = &self.display {
            validate_hyperlink_text(
                display,
                32_768,
                "use.office.hyperlink_display_invalid",
                "Native Office hyperlink display text",
            )?;
        }
        if let Some(tooltip) = &self.tooltip {
            validate_hyperlink_text(
                tooltip,
                4_096,
                "use.office.hyperlink_tooltip_invalid",
                "Native Office hyperlink tooltips",
            )?;
        }
        Ok(())
    }

    pub(crate) fn default_display(&self) -> &str {
        match &self.target {
            NativeOfficeHyperlinkTarget::External { uri } => uri,
            NativeOfficeHyperlinkTarget::Internal { location } => location,
        }
    }
}

fn validate_external_uri(uri: &str) -> UseResult<()> {
    validate_hyperlink_text(
        uri,
        2_048,
        "use.office.hyperlink_uri_invalid",
        "Native Office external hyperlink URIs",
    )?;
    let parsed = Url::parse(uri).map_err(|error| {
        super::editor_error(
            "use.office.hyperlink_uri_invalid",
            format!("Native Office external hyperlink URI is invalid: {error}"),
        )
    })?;
    let valid = match parsed.scheme() {
        "http" | "https" => {
            parsed.host_str().is_some()
                && parsed.username().is_empty()
                && parsed.password().is_none()
        }
        "mailto" => !parsed.path().is_empty(),
        _ => false,
    };
    if !valid {
        return Err(super::editor_error(
            "use.office.hyperlink_uri_invalid",
            "Native Office external hyperlinks require an absolute HTTP, HTTPS, or mailto URI without embedded credentials.",
        ));
    }
    Ok(())
}

fn validate_hyperlink_text(
    value: &str,
    max_bytes: usize,
    code: &str,
    label: &str,
) -> UseResult<()> {
    if value.is_empty()
        || value.len() > max_bytes
        || value.trim() != value
        || value.chars().any(char::is_control)
    {
        return Err(super::editor_error(
            code,
            format!(
                "{label} must contain 1-{max_bytes} non-control UTF-8 bytes without surrounding whitespace."
            ),
        ));
    }
    Ok(())
}

/// An exact Presentation comment position in English Metric Units.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficeCommentPosition {
    pub x_emu: i32,
    pub y_emu: i32,
}

impl NativeOfficeCommentPosition {
    pub const fn new(x_emu: i32, y_emu: i32) -> Self {
        Self { x_emu, y_emu }
    }
}

/// A complete typed legacy Office comment.
///
/// Word uses the mutation parent as its paragraph or run anchor, Spreadsheet
/// uses a cell parent, and Presentation uses a slide parent plus an optional
/// position. Modern threaded comments are a separate format and are not
/// represented by this type.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficeComment {
    pub author: String,
    pub text: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub initials: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub position: Option<NativeOfficeCommentPosition>,
}

impl NativeOfficeComment {
    /// Creates a validated legacy Office comment.
    pub fn new(author: impl Into<String>, text: impl Into<String>) -> UseResult<Self> {
        let comment = Self {
            author: author.into(),
            text: text.into(),
            initials: None,
            position: None,
        };
        comment.validate()?;
        Ok(comment)
    }

    pub fn with_initials(mut self, initials: impl Into<String>) -> Self {
        self.initials = Some(initials.into());
        self
    }

    pub fn with_position(mut self, position: NativeOfficeCommentPosition) -> Self {
        self.position = Some(position);
        self
    }

    pub(crate) fn validate(&self) -> UseResult<()> {
        validate_comment_scalar(
            &self.author,
            255,
            false,
            "use.office.comment_author_invalid",
            "Native Office comment authors",
        )?;
        validate_comment_scalar(
            &self.text,
            32_768,
            true,
            "use.office.comment_text_invalid",
            "Native Office comment text",
        )?;
        if let Some(initials) = &self.initials {
            validate_comment_scalar(
                initials,
                32,
                false,
                "use.office.comment_initials_invalid",
                "Native Office comment initials",
            )?;
        }
        Ok(())
    }
}

/// A partial typed update for an existing legacy Office comment.
#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficeCommentUpdate {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub author: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub initials: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub position: Option<NativeOfficeCommentPosition>,
}

impl NativeOfficeCommentUpdate {
    pub fn is_empty(&self) -> bool {
        self.author.is_none()
            && self.text.is_none()
            && self.initials.is_none()
            && self.position.is_none()
    }

    pub(crate) fn validate(&self) -> UseResult<()> {
        if self.is_empty() {
            return Err(super::editor_error(
                "use.office.comment_update_empty",
                "Native Office comment updates require at least one typed property.",
            ));
        }
        if let Some(author) = &self.author {
            validate_comment_scalar(
                author,
                255,
                false,
                "use.office.comment_author_invalid",
                "Native Office comment authors",
            )?;
        }
        if let Some(text) = &self.text {
            validate_comment_scalar(
                text,
                32_768,
                true,
                "use.office.comment_text_invalid",
                "Native Office comment text",
            )?;
        }
        if let Some(initials) = &self.initials {
            validate_comment_scalar(
                initials,
                32,
                false,
                "use.office.comment_initials_invalid",
                "Native Office comment initials",
            )?;
        }
        Ok(())
    }
}

fn validate_comment_scalar(
    value: &str,
    max_bytes: usize,
    multiline: bool,
    code: &str,
    label: &str,
) -> UseResult<()> {
    let invalid_control = value.chars().any(|character| {
        character.is_control() && !(multiline && matches!(character, '\n' | '\t'))
    });
    if value.is_empty() || value.len() > max_bytes || value.trim() != value || invalid_control {
        let allowed = if multiline {
            "; line feeds and tabs are allowed"
        } else {
            ""
        };
        return Err(super::editor_error(
            code,
            format!(
                "{label} must contain 1-{max_bytes} UTF-8 bytes without surrounding whitespace or unsupported control characters{allowed}."
            ),
        ));
    }
    Ok(())
}

/// Typed Spreadsheet cell content written without a shared-string dependency.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(tag = "type", rename_all = "kebab-case", deny_unknown_fields)]
pub enum SpreadsheetCellValue {
    Text { value: String },
    Number { value: String },
    Boolean { value: bool },
    Formula { expression: String },
}

/// Zero-based insertion selector shared by native move and copy operations.
///
/// `Index` is evaluated after removing the source for a move. `Before` and
/// `After` use stable semantic paths and are resolved before the mutation.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "kebab-case", deny_unknown_fields)]
pub enum NativeOfficeInsertPosition {
    Index { index: usize },
    Before { path: String },
    After { path: String },
}

impl NativeOfficeInsertPosition {
    pub fn at_index(index: usize) -> Self {
        Self::Index { index }
    }

    pub fn before(path: impl Into<String>) -> Self {
        Self::Before { path: path.into() }
    }

    pub fn after(path: impl Into<String>) -> Self {
        Self::After { path: path.into() }
    }
}

/// Raster formats that can be embedded without an external Office runtime.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum NativeOfficeImageFormat {
    Png,
    #[serde(alias = "jpg")]
    Jpeg,
    Gif,
}

/// Validated dimensions and format of a native raster image byte slice.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeImageMetadata {
    pub format: NativeOfficeImageFormat,
    pub width_px: u32,
    pub height_px: u32,
}

impl NativeOfficeImageFormat {
    pub(crate) fn extension(self) -> &'static str {
        match self {
            Self::Png => "png",
            Self::Jpeg => "jpeg",
            Self::Gif => "gif",
        }
    }

    pub(crate) fn content_type(self) -> &'static str {
        match self {
            Self::Png => "image/png",
            Self::Jpeg => "image/jpeg",
            Self::Gif => "image/gif",
        }
    }
}

/// A bounded, base64-serializable raster image for a native Office mutation.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficeImage {
    pub(super) format: NativeOfficeImageFormat,
    pub(super) data: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub(super) name: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub(super) alt_text: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub(super) width_px: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub(super) height_px: Option<u32>,
}

impl NativeOfficeImage {
    /// Validates PNG, JPEG, or GIF bytes without base64-encoding or mutating them.
    pub fn inspect_bytes(bytes: impl AsRef<[u8]>) -> UseResult<NativeOfficeImageMetadata> {
        let metadata = super::image::inspect_image(bytes.as_ref(), None)?;
        Ok(NativeOfficeImageMetadata {
            format: metadata.format,
            width_px: metadata.width_px,
            height_px: metadata.height_px,
        })
    }

    /// Detects and validates PNG, JPEG, or GIF bytes before serializing them.
    pub fn from_bytes(bytes: impl AsRef<[u8]>) -> UseResult<Self> {
        let bytes = bytes.as_ref();
        let metadata = Self::inspect_bytes(bytes)?;
        Ok(Self {
            format: metadata.format,
            data: base64::engine::general_purpose::STANDARD.encode(bytes),
            name: None,
            alt_text: None,
            width_px: None,
            height_px: None,
        })
    }

    pub fn format(&self) -> NativeOfficeImageFormat {
        self.format
    }

    pub fn with_name(mut self, name: impl Into<String>) -> Self {
        self.name = Some(name.into());
        self
    }

    pub fn with_alt_text(mut self, alt_text: impl Into<String>) -> Self {
        self.alt_text = Some(alt_text.into());
        self
    }

    pub fn with_width_px(mut self, width_px: u32) -> Self {
        self.width_px = Some(width_px);
        self
    }

    pub fn with_height_px(mut self, height_px: u32) -> Self {
        self.height_px = Some(height_px);
        self
    }
}

/// Receipt for an image embedded into a semantic Office document location.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeCreatedImage {
    pub path: String,
    pub part: String,
    pub parent: String,
    pub owner_part: String,
    pub relationship_id: String,
    pub format: NativeOfficeImageFormat,
    pub width_px: u32,
    pub height_px: u32,
}

/// Matching semantics for a bounded native Office text replacement.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum NativeOfficeTextMatchMode {
    /// Case-sensitive, non-overlapping substring matching.
    Literal,
    /// Linear-time Rust regular expressions with capture expansion in the
    /// replacement string.
    Regex,
}

/// Maximum UTF-8 byte length accepted for a native Office find expression.
pub const MAX_NATIVE_OFFICE_FIND_BYTES: usize = 64 * 1024;
/// Maximum UTF-8 byte length accepted for a native Office replacement value.
pub const MAX_NATIVE_OFFICE_REPLACEMENT_BYTES: usize = 1024 * 1024;
/// Maximum semantic matches accepted in one native Office replacement.
pub const MAX_NATIVE_OFFICE_TEXT_MATCHES: usize = 100_000;
/// Maximum expanded replacement bytes accepted in one native Office operation.
pub const MAX_NATIVE_OFFICE_TEXT_REPLACEMENT_OUTPUT_BYTES: usize = 64 * 1024 * 1024;
/// Maximum Spreadsheet cells addressable by one replacement scope.
pub const MAX_NATIVE_OFFICE_TEXT_SCOPE_CELLS: usize = 100_000;

const MAX_NATIVE_OFFICE_REGEX_COMPILED_BYTES: usize = 8 * 1024 * 1024;

/// Typed, bounded find/replace value shared by Word, Spreadsheet, and
/// Presentation.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficeTextReplacement {
    pub find: String,
    pub replace: String,
    pub mode: NativeOfficeTextMatchMode,
}

impl NativeOfficeTextReplacement {
    /// Creates a validated literal replacement.
    pub fn literal(find: impl Into<String>, replace: impl Into<String>) -> UseResult<Self> {
        Self::new(find, replace, NativeOfficeTextMatchMode::Literal)
    }

    /// Creates a validated regular-expression replacement.
    pub fn regex(find: impl Into<String>, replace: impl Into<String>) -> UseResult<Self> {
        Self::new(find, replace, NativeOfficeTextMatchMode::Regex)
    }

    pub fn new(
        find: impl Into<String>,
        replace: impl Into<String>,
        mode: NativeOfficeTextMatchMode,
    ) -> UseResult<Self> {
        let replacement = Self {
            find: find.into(),
            replace: replace.into(),
            mode,
        };
        replacement.validate()?;
        Ok(replacement)
    }

    pub(crate) fn validate(&self) -> UseResult<()> {
        if self.find.is_empty() || self.find.len() > MAX_NATIVE_OFFICE_FIND_BYTES {
            return Err(super::editor_error(
                "use.office.text_find_invalid",
                format!(
                    "Native Office find expressions must contain 1-{MAX_NATIVE_OFFICE_FIND_BYTES} UTF-8 bytes."
                ),
            )
            .with_detail("findBytes", self.find.len()));
        }
        if self.replace.len() > MAX_NATIVE_OFFICE_REPLACEMENT_BYTES {
            return Err(super::editor_error(
                "use.office.text_replacement_invalid",
                format!(
                    "Native Office replacement values cannot exceed {MAX_NATIVE_OFFICE_REPLACEMENT_BYTES} UTF-8 bytes."
                ),
            )
            .with_detail("replacementBytes", self.replace.len()));
        }
        validate_replacement_xml_text(&self.replace)?;
        if self.mode == NativeOfficeTextMatchMode::Regex {
            let expression = regex::RegexBuilder::new(&self.find)
                .size_limit(MAX_NATIVE_OFFICE_REGEX_COMPILED_BYTES)
                .build()
                .map_err(|error| {
                    super::editor_error(
                        "use.office.text_regex_invalid",
                        format!("Native Office regular expression is invalid: {error}"),
                    )
                })?;
            if expression.is_match("") {
                return Err(super::editor_error(
                    "use.office.text_regex_empty_match",
                    "Native Office regular expressions must consume at least one character.",
                ));
            }
        }
        Ok(())
    }
}

fn validate_replacement_xml_text(value: &str) -> UseResult<()> {
    if let Some(character) = value.chars().find(|character| {
        !matches!(*character, '\u{9}' | '\u{a}' | '\u{d}')
            && (*character < '\u{20}' || matches!(*character, '\u{fffe}' | '\u{ffff}'))
    }) {
        return Err(super::editor_error(
            "use.office.text_replacement_invalid",
            format!(
                "Native Office replacement contains XML-forbidden character U+{:04X}.",
                u32::from(character)
            ),
        )
        .with_detail("codePoint", format!("U+{:04X}", u32::from(character))));
    }
    Ok(())
}

/// Per-mutation receipt for a native Office text replacement.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeTextReplacementResult {
    pub path: String,
    pub mode: NativeOfficeTextMatchMode,
    pub match_count: usize,
    pub changed: bool,
    pub changed_parts: Vec<String>,
}

/// Typed in-process mutation supported by an atomic native batch.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(tag = "operation", rename_all = "kebab-case", deny_unknown_fields)]
pub enum NativeOfficeMutation {
    ReplaceText {
        path: String,
        replacement: NativeOfficeTextReplacement,
    },
    SetText {
        path: String,
        text: String,
    },
    SetTextFormat {
        path: String,
        format: NativeOfficeTextFormat,
    },
    SetCellFormat {
        path: String,
        format: NativeSpreadsheetCellFormat,
    },
    SetHyperlink {
        path: String,
        hyperlink: NativeOfficeHyperlink,
    },
    SetComment {
        path: String,
        update: NativeOfficeCommentUpdate,
    },
    SetTableColumnWidth {
        path: String,
        #[serde(rename = "widthEmu")]
        width_emu: u64,
    },
    SetCellValue {
        path: String,
        value: SpreadsheetCellValue,
    },
    RecalculateSpreadsheetFormulas,
    AddSpreadsheetTable {
        sheet: String,
        table: NativeSpreadsheetTable,
    },
    SetSpreadsheetTable {
        path: String,
        table: NativeSpreadsheetTable,
    },
    AddSpreadsheetAutoFilter {
        sheet: String,
        filter: NativeSpreadsheetAutoFilter,
    },
    SetSpreadsheetAutoFilter {
        path: String,
        filter: NativeSpreadsheetAutoFilter,
    },
    SortSpreadsheetRange {
        path: String,
        sort: NativeSpreadsheetSort,
    },
    ImportSpreadsheetDelimited {
        sheet: String,
        import: NativeSpreadsheetDelimitedImport,
    },
    SetSpreadsheetFrozenPane {
        sheet: String,
        pane: NativeSpreadsheetFrozenPane,
    },
    AddNamedRange {
        #[serde(rename = "namedRange")]
        named_range: NativeSpreadsheetNamedRange,
    },
    SetNamedRange {
        path: String,
        #[serde(rename = "namedRange")]
        named_range: NativeSpreadsheetNamedRange,
    },
    AddConditionalFormat {
        sheet: String,
        #[serde(rename = "conditionalFormat")]
        conditional_format: NativeSpreadsheetConditionalFormat,
    },
    SetConditionalFormat {
        path: String,
        #[serde(rename = "conditionalFormat")]
        conditional_format: NativeSpreadsheetConditionalFormat,
    },
    AddDataValidation {
        sheet: String,
        validation: NativeSpreadsheetDataValidation,
    },
    SetDataValidation {
        path: String,
        validation: NativeSpreadsheetDataValidation,
    },
    MergeCells {
        path: String,
    },
    UnmergeCells {
        path: String,
    },
    AddParagraph {
        parent: String,
        text: String,
    },
    AddTable {
        parent: String,
        rows: usize,
        columns: usize,
    },
    AddTableRow {
        parent: String,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        columns: Option<usize>,
    },
    AddTableColumn {
        parent: String,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        index: Option<usize>,
        #[serde(default, skip_serializing_if = "String::is_empty")]
        text: String,
    },
    AddTableCell {
        parent: String,
        text: String,
    },
    AddSlide {
        parent: String,
        title: String,
    },
    AddShape {
        parent: String,
        text: String,
    },
    AddImage {
        parent: String,
        image: NativeOfficeImage,
    },
    AddComment {
        parent: String,
        comment: NativeOfficeComment,
    },
    AddPart {
        parent: String,
        #[serde(rename = "type")]
        part_type: NativeOfficePartType,
    },
    AddWorksheet {
        name: String,
    },
    InsertRows {
        sheet: String,
        start: u32,
        count: u32,
    },
    DeleteRows {
        sheet: String,
        start: u32,
        count: u32,
    },
    InsertColumns {
        sheet: String,
        start: String,
        count: u32,
    },
    DeleteColumns {
        sheet: String,
        start: String,
        count: u32,
    },
    RenameWorksheet {
        path: String,
        name: String,
    },
    MoveWorksheet {
        path: String,
        position: usize,
    },
    CopyWorksheet {
        path: String,
        name: String,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        position: Option<usize>,
    },
    Move {
        path: String,
        #[serde(rename = "to", default, skip_serializing_if = "Option::is_none")]
        target_parent: Option<String>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        position: Option<NativeOfficeInsertPosition>,
    },
    Copy {
        path: String,
        #[serde(rename = "to", default, skip_serializing_if = "Option::is_none")]
        target_parent: Option<String>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        position: Option<NativeOfficeInsertPosition>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        name: Option<String>,
    },
    Swap {
        path: String,
        with: String,
    },
    ReplaceXmlPart {
        part: String,
        xml: String,
    },
    Remove {
        path: String,
    },
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeOfficeSwapResult {
    pub first: String,
    pub second: String,
}

#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub struct NativeBatchResult {
    pub applied: usize,
    pub paths: Vec<String>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub swaps: Vec<NativeOfficeSwapResult>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub created_parts: Vec<NativeCreatedPart>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub created_images: Vec<NativeCreatedImage>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub text_replacements: Vec<NativeOfficeTextReplacementResult>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub spreadsheet_imports: Vec<NativeSpreadsheetImportResult>,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub spreadsheet_calculations: Vec<SpreadsheetFormulaCalculation>,
}
