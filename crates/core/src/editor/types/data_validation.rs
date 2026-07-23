use serde::{Deserialize, Serialize};

/// Closed SpreadsheetML data-validation rule families supported natively.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum NativeSpreadsheetDataValidationType {
    List,
    Whole,
    Decimal,
    Date,
    Time,
    TextLength,
    Custom,
}

/// Comparison operators used by numeric, date, time, and text-length rules.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum NativeSpreadsheetDataValidationOperator {
    Between,
    NotBetween,
    Equal,
    NotEqual,
    GreaterThan,
    GreaterThanOrEqual,
    LessThan,
    LessThanOrEqual,
}

/// Spreadsheet input-error alert severity.
#[derive(Debug, Default, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum NativeSpreadsheetDataValidationErrorStyle {
    #[default]
    Stop,
    Warning,
    Information,
}

/// One complete typed Spreadsheet data-validation rule.
///
/// `ranges` contains one or more rectangular A1 areas on one worksheet.
/// Formula text is normalized for the selected rule type when the mutation is
/// applied; for example, inline list values are quoted and ISO dates/times are
/// converted to Spreadsheet serial values.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeSpreadsheetDataValidation {
    #[serde(rename = "type")]
    pub validation_type: NativeSpreadsheetDataValidationType,
    pub ranges: Vec<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub operator: Option<NativeSpreadsheetDataValidationOperator>,
    pub formula1: String,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub formula2: Option<String>,
    #[serde(default = "default_true", skip_serializing_if = "is_true")]
    pub allow_blank: bool,
    #[serde(default = "default_true", skip_serializing_if = "is_true")]
    pub show_input: bool,
    #[serde(default = "default_true", skip_serializing_if = "is_true")]
    pub show_error: bool,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub prompt_title: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub prompt: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub error_title: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub error: Option<String>,
    #[serde(default, skip_serializing_if = "is_stop")]
    pub error_style: NativeSpreadsheetDataValidationErrorStyle,
    #[serde(default = "default_true", skip_serializing_if = "is_true")]
    pub in_cell_dropdown: bool,
}

impl NativeSpreadsheetDataValidation {
    /// Creates a one-range rule with Office-friendly prompt, error, blank, and
    /// dropdown defaults enabled.
    pub fn new(
        validation_type: NativeSpreadsheetDataValidationType,
        range: impl Into<String>,
        formula1: impl Into<String>,
    ) -> Self {
        Self {
            validation_type,
            ranges: vec![range.into()],
            operator: None,
            formula1: formula1.into(),
            formula2: None,
            allow_blank: true,
            show_input: true,
            show_error: true,
            prompt_title: None,
            prompt: None,
            error_title: None,
            error: None,
            error_style: NativeSpreadsheetDataValidationErrorStyle::Stop,
            in_cell_dropdown: true,
        }
    }

    pub fn with_range(mut self, range: impl Into<String>) -> Self {
        self.ranges.push(range.into());
        self
    }

    pub fn with_operator(mut self, operator: NativeSpreadsheetDataValidationOperator) -> Self {
        self.operator = Some(operator);
        self
    }

    pub fn with_formula2(mut self, formula: impl Into<String>) -> Self {
        self.formula2 = Some(formula.into());
        self
    }

    pub fn with_allow_blank(mut self, allow_blank: bool) -> Self {
        self.allow_blank = allow_blank;
        self
    }

    pub fn with_input_message(
        mut self,
        title: impl Into<String>,
        message: impl Into<String>,
    ) -> Self {
        self.prompt_title = Some(title.into());
        self.prompt = Some(message.into());
        self.show_input = true;
        self
    }

    pub fn with_error_message(
        mut self,
        style: NativeSpreadsheetDataValidationErrorStyle,
        title: impl Into<String>,
        message: impl Into<String>,
    ) -> Self {
        self.error_style = style;
        self.error_title = Some(title.into());
        self.error = Some(message.into());
        self.show_error = true;
        self
    }

    pub fn with_in_cell_dropdown(mut self, visible: bool) -> Self {
        self.in_cell_dropdown = visible;
        self
    }
}

const fn default_true() -> bool {
    true
}

const fn is_true(value: &bool) -> bool {
    *value
}

const fn is_stop(value: &NativeSpreadsheetDataValidationErrorStyle) -> bool {
    matches!(value, NativeSpreadsheetDataValidationErrorStyle::Stop)
}
