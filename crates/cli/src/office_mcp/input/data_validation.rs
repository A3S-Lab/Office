use a3s_office::{
    NativeSpreadsheetDataValidation, NativeSpreadsheetDataValidationErrorStyle,
    NativeSpreadsheetDataValidationOperator, NativeSpreadsheetDataValidationType,
};
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Copy, Serialize, Deserialize, schemars::JsonSchema)]
#[serde(rename_all = "camelCase")]
enum OfficeDataValidationType {
    List,
    Whole,
    Decimal,
    Date,
    Time,
    TextLength,
    Custom,
}

impl From<OfficeDataValidationType> for NativeSpreadsheetDataValidationType {
    fn from(value: OfficeDataValidationType) -> Self {
        match value {
            OfficeDataValidationType::List => Self::List,
            OfficeDataValidationType::Whole => Self::Whole,
            OfficeDataValidationType::Decimal => Self::Decimal,
            OfficeDataValidationType::Date => Self::Date,
            OfficeDataValidationType::Time => Self::Time,
            OfficeDataValidationType::TextLength => Self::TextLength,
            OfficeDataValidationType::Custom => Self::Custom,
        }
    }
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, schemars::JsonSchema)]
#[serde(rename_all = "camelCase")]
enum OfficeDataValidationOperator {
    Between,
    NotBetween,
    Equal,
    NotEqual,
    GreaterThan,
    GreaterThanOrEqual,
    LessThan,
    LessThanOrEqual,
}

impl From<OfficeDataValidationOperator> for NativeSpreadsheetDataValidationOperator {
    fn from(value: OfficeDataValidationOperator) -> Self {
        match value {
            OfficeDataValidationOperator::Between => Self::Between,
            OfficeDataValidationOperator::NotBetween => Self::NotBetween,
            OfficeDataValidationOperator::Equal => Self::Equal,
            OfficeDataValidationOperator::NotEqual => Self::NotEqual,
            OfficeDataValidationOperator::GreaterThan => Self::GreaterThan,
            OfficeDataValidationOperator::GreaterThanOrEqual => Self::GreaterThanOrEqual,
            OfficeDataValidationOperator::LessThan => Self::LessThan,
            OfficeDataValidationOperator::LessThanOrEqual => Self::LessThanOrEqual,
        }
    }
}

#[derive(Debug, Default, Clone, Copy, Serialize, Deserialize, schemars::JsonSchema)]
#[serde(rename_all = "camelCase")]
enum OfficeDataValidationErrorStyle {
    #[default]
    Stop,
    Warning,
    Information,
}

impl From<OfficeDataValidationErrorStyle> for NativeSpreadsheetDataValidationErrorStyle {
    fn from(value: OfficeDataValidationErrorStyle) -> Self {
        match value {
            OfficeDataValidationErrorStyle::Stop => Self::Stop,
            OfficeDataValidationErrorStyle::Warning => Self::Warning,
            OfficeDataValidationErrorStyle::Information => Self::Information,
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, schemars::JsonSchema)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub(in crate::office_mcp) struct OfficeDataValidation {
    /// List, whole, decimal, date, time, text-length, or custom formula rule.
    #[serde(rename = "type")]
    validation_type: OfficeDataValidationType,
    /// One or more disjoint rectangular A1 ranges on the target worksheet.
    ranges: Vec<String>,
    /// Required for numeric, date, time, and text-length validation.
    operator: Option<OfficeDataValidationOperator>,
    /// Required rule formula or inline list source, limited to 255 characters.
    formula1: String,
    /// Second bound for between and not-between rules.
    formula2: Option<String>,
    /// Whether empty cells are valid. Defaults to true for new A3S rules.
    #[serde(default = "default_true")]
    allow_blank: bool,
    /// Whether selecting a target cell shows its input prompt. Defaults to true.
    #[serde(default = "default_true")]
    show_input: bool,
    /// Whether invalid input shows an error alert. Defaults to true.
    #[serde(default = "default_true")]
    show_error: bool,
    prompt_title: Option<String>,
    prompt: Option<String>,
    error_title: Option<String>,
    error: Option<String>,
    #[serde(default)]
    error_style: OfficeDataValidationErrorStyle,
    /// User-facing dropdown visibility for list rules. Defaults to true.
    #[serde(default = "default_true")]
    in_cell_dropdown: bool,
}

impl From<OfficeDataValidation> for NativeSpreadsheetDataValidation {
    fn from(value: OfficeDataValidation) -> Self {
        Self {
            validation_type: value.validation_type.into(),
            ranges: value.ranges,
            operator: value.operator.map(Into::into),
            formula1: value.formula1,
            formula2: value.formula2,
            allow_blank: value.allow_blank,
            show_input: value.show_input,
            show_error: value.show_error,
            prompt_title: value.prompt_title,
            prompt: value.prompt,
            error_title: value.error_title,
            error: value.error,
            error_style: value.error_style.into(),
            in_cell_dropdown: value.in_cell_dropdown,
        }
    }
}

const fn default_true() -> bool {
    true
}
