use std::collections::BTreeSet;

use a3s_use_core::UseResult;
use serde::{Deserialize, Serialize};

use crate::spreadsheet_reference::CellRange;

const MAX_FILTER_COLUMNS: usize = 16_384;
const MAX_FILTER_VALUES_PER_COLUMN: usize = 10_000;
const MAX_FILTER_VALUE_CHARACTERS: usize = 32_767;
const MAX_FILTER_TEXT_BYTES: usize = 1_048_576;

/// One supported dynamic Spreadsheet filter family.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum NativeSpreadsheetDynamicFilter {
    AboveAverage,
    BelowAverage,
    Tomorrow,
    Today,
    Yesterday,
    NextWeek,
    ThisWeek,
    LastWeek,
    NextMonth,
    ThisMonth,
    LastMonth,
    NextQuarter,
    ThisQuarter,
    LastQuarter,
    NextYear,
    ThisYear,
    LastYear,
    YearToDate,
    Quarter1,
    Quarter2,
    Quarter3,
    Quarter4,
    Month1,
    Month2,
    Month3,
    Month4,
    Month5,
    Month6,
    Month7,
    Month8,
    Month9,
    Month10,
    Month11,
    Month12,
}

impl NativeSpreadsheetDynamicFilter {
    pub(crate) fn ooxml_name(self) -> &'static str {
        match self {
            Self::AboveAverage => "aboveAverage",
            Self::BelowAverage => "belowAverage",
            Self::Tomorrow => "tomorrow",
            Self::Today => "today",
            Self::Yesterday => "yesterday",
            Self::NextWeek => "nextWeek",
            Self::ThisWeek => "thisWeek",
            Self::LastWeek => "lastWeek",
            Self::NextMonth => "nextMonth",
            Self::ThisMonth => "thisMonth",
            Self::LastMonth => "lastMonth",
            Self::NextQuarter => "nextQuarter",
            Self::ThisQuarter => "thisQuarter",
            Self::LastQuarter => "lastQuarter",
            Self::NextYear => "nextYear",
            Self::ThisYear => "thisYear",
            Self::LastYear => "lastYear",
            Self::YearToDate => "yearToDate",
            Self::Quarter1 => "Q1",
            Self::Quarter2 => "Q2",
            Self::Quarter3 => "Q3",
            Self::Quarter4 => "Q4",
            Self::Month1 => "M1",
            Self::Month2 => "M2",
            Self::Month3 => "M3",
            Self::Month4 => "M4",
            Self::Month5 => "M5",
            Self::Month6 => "M6",
            Self::Month7 => "M7",
            Self::Month8 => "M8",
            Self::Month9 => "M9",
            Self::Month10 => "M10",
            Self::Month11 => "M11",
            Self::Month12 => "M12",
        }
    }

    pub(crate) fn from_ooxml_name(value: &str) -> Option<Self> {
        Some(match value {
            "aboveAverage" => Self::AboveAverage,
            "belowAverage" => Self::BelowAverage,
            "tomorrow" => Self::Tomorrow,
            "today" => Self::Today,
            "yesterday" => Self::Yesterday,
            "nextWeek" => Self::NextWeek,
            "thisWeek" => Self::ThisWeek,
            "lastWeek" => Self::LastWeek,
            "nextMonth" => Self::NextMonth,
            "thisMonth" => Self::ThisMonth,
            "lastMonth" => Self::LastMonth,
            "nextQuarter" => Self::NextQuarter,
            "thisQuarter" => Self::ThisQuarter,
            "lastQuarter" => Self::LastQuarter,
            "nextYear" => Self::NextYear,
            "thisYear" => Self::ThisYear,
            "lastYear" => Self::LastYear,
            "yearToDate" => Self::YearToDate,
            "Q1" => Self::Quarter1,
            "Q2" => Self::Quarter2,
            "Q3" => Self::Quarter3,
            "Q4" => Self::Quarter4,
            "M1" => Self::Month1,
            "M2" => Self::Month2,
            "M3" => Self::Month3,
            "M4" => Self::Month4,
            "M5" => Self::Month5,
            "M6" => Self::Month6,
            "M7" => Self::Month7,
            "M8" => Self::Month8,
            "M9" => Self::Month9,
            "M10" => Self::Month10,
            "M11" => Self::Month11,
            "M12" => Self::Month12,
            _ => return None,
        })
    }
}

/// One closed filter criterion applied to a zero-based AutoFilter column.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(
    tag = "type",
    rename_all = "kebab-case",
    rename_all_fields = "camelCase",
    deny_unknown_fields
)]
pub enum NativeSpreadsheetFilterCriteria {
    Values {
        values: Vec<String>,
        #[serde(default)]
        include_blanks: bool,
    },
    Equals {
        value: String,
    },
    NotEquals {
        value: String,
    },
    Contains {
        value: String,
    },
    DoesNotContain {
        value: String,
    },
    BeginsWith {
        value: String,
    },
    EndsWith {
        value: String,
    },
    GreaterThan {
        value: String,
    },
    GreaterThanOrEqual {
        value: String,
    },
    LessThan {
        value: String,
    },
    LessThanOrEqual {
        value: String,
    },
    Between {
        lower: String,
        upper: String,
    },
    NotBetween {
        lower: String,
        upper: String,
    },
    Blanks,
    NonBlanks,
    Top {
        count: u16,
    },
    TopPercent {
        percent: u8,
    },
    Bottom {
        count: u16,
    },
    BottomPercent {
        percent: u8,
    },
    Dynamic {
        kind: NativeSpreadsheetDynamicFilter,
    },
}

impl NativeSpreadsheetFilterCriteria {
    pub(crate) fn validate(&self) -> UseResult<()> {
        match self {
            Self::Values {
                values,
                include_blanks,
            } => {
                if values.len() > MAX_FILTER_VALUES_PER_COLUMN {
                    return Err(filter_error(
                        "use.office.spreadsheet_filter_value_limit",
                        format!(
                            "Spreadsheet value filters accept at most {MAX_FILTER_VALUES_PER_COLUMN} values per column."
                        ),
                    ));
                }
                if values.is_empty() && !include_blanks {
                    return Err(filter_error(
                        "use.office.spreadsheet_filter_values_empty",
                        "Spreadsheet value filters require at least one value or includeBlanks=true.",
                    ));
                }
                let mut observed = BTreeSet::new();
                for value in values {
                    validate_filter_text(value, "value-filter entries", false)?;
                    if !observed.insert(value) {
                        return Err(filter_error(
                            "use.office.spreadsheet_filter_value_duplicate",
                            format!("Spreadsheet value filter entry '{value}' is duplicated."),
                        ));
                    }
                }
            }
            Self::Equals { value }
            | Self::NotEquals { value }
            | Self::Contains { value }
            | Self::DoesNotContain { value }
            | Self::BeginsWith { value }
            | Self::EndsWith { value }
            | Self::GreaterThan { value }
            | Self::GreaterThanOrEqual { value }
            | Self::LessThan { value }
            | Self::LessThanOrEqual { value } => {
                validate_filter_text(value, "comparison values", false)?;
            }
            Self::Between { lower, upper } | Self::NotBetween { lower, upper } => {
                validate_filter_text(lower, "lower comparison bounds", false)?;
                validate_filter_text(upper, "upper comparison bounds", false)?;
            }
            Self::Top { count } | Self::Bottom { count } => {
                if !(1..=500).contains(count) {
                    return Err(filter_error(
                        "use.office.spreadsheet_filter_top_count_invalid",
                        "Spreadsheet top/bottom count filters require a value from 1 through 500.",
                    ));
                }
            }
            Self::TopPercent { percent } | Self::BottomPercent { percent } => {
                if !(1..=100).contains(percent) {
                    return Err(filter_error(
                        "use.office.spreadsheet_filter_percent_invalid",
                        "Spreadsheet top/bottom percentage filters require a value from 1 through 100.",
                    ));
                }
            }
            Self::Blanks | Self::NonBlanks | Self::Dynamic { .. } => {}
        }
        Ok(())
    }

    pub(crate) fn text_bytes(&self) -> usize {
        match self {
            Self::Values { values, .. } => values.iter().map(String::len).sum(),
            Self::Equals { value }
            | Self::NotEquals { value }
            | Self::Contains { value }
            | Self::DoesNotContain { value }
            | Self::BeginsWith { value }
            | Self::EndsWith { value }
            | Self::GreaterThan { value }
            | Self::GreaterThanOrEqual { value }
            | Self::LessThan { value }
            | Self::LessThanOrEqual { value } => value.len(),
            Self::Between { lower, upper } | Self::NotBetween { lower, upper } => {
                lower.len().saturating_add(upper.len())
            }
            Self::Blanks
            | Self::NonBlanks
            | Self::Top { .. }
            | Self::TopPercent { .. }
            | Self::Bottom { .. }
            | Self::BottomPercent { .. }
            | Self::Dynamic { .. } => 0,
        }
    }
}

/// One criterion assigned to a zero-based column offset inside a filter range.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeSpreadsheetFilterColumn {
    pub column: u32,
    pub criteria: NativeSpreadsheetFilterCriteria,
}

impl NativeSpreadsheetFilterColumn {
    pub fn new(column: u32, criteria: NativeSpreadsheetFilterCriteria) -> Self {
        Self { column, criteria }
    }
}

/// A complete worksheet AutoFilter value.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeSpreadsheetAutoFilter {
    pub range: String,
    #[serde(default, skip_serializing_if = "Vec::is_empty")]
    pub columns: Vec<NativeSpreadsheetFilterColumn>,
}

impl NativeSpreadsheetAutoFilter {
    pub fn new(range: impl Into<String>) -> Self {
        Self {
            range: range.into(),
            columns: Vec::new(),
        }
    }

    pub fn with_filter(mut self, column: u32, criteria: NativeSpreadsheetFilterCriteria) -> Self {
        self.columns
            .push(NativeSpreadsheetFilterColumn::new(column, criteria));
        self
    }

    pub(crate) fn validate(&self) -> UseResult<CellRange> {
        let range = CellRange::parse(&self.range).map_err(|error| {
            filter_error(
                "use.office.spreadsheet_filter_range_invalid",
                format!(
                    "Spreadsheet AutoFilter range '{}' is invalid: {error}",
                    self.range
                ),
            )
        })?;
        let width = usize::try_from(range.end.column - range.start.column + 1).map_err(|_| {
            filter_error(
                "use.office.spreadsheet_filter_range_invalid",
                "Spreadsheet AutoFilter width does not fit this platform.",
            )
        })?;
        self.validate_for_width(width)?;
        Ok(range)
    }

    pub(crate) fn validate_for_width(&self, width: usize) -> UseResult<()> {
        if self.columns.len() > MAX_FILTER_COLUMNS {
            return Err(filter_error(
                "use.office.spreadsheet_filter_column_limit",
                format!(
                    "Spreadsheet AutoFilters accept at most {MAX_FILTER_COLUMNS} filter columns."
                ),
            ));
        }
        let mut observed = BTreeSet::new();
        let mut text_bytes = 0_usize;
        for filter in &self.columns {
            let column = usize::try_from(filter.column).map_err(|_| {
                filter_error(
                    "use.office.spreadsheet_filter_column_invalid",
                    "Spreadsheet filter column index does not fit this platform.",
                )
            })?;
            if column >= width {
                return Err(filter_error(
                    "use.office.spreadsheet_filter_column_invalid",
                    format!(
                        "Spreadsheet filter column {} is outside the zero-based range width {width}.",
                        filter.column
                    ),
                ));
            }
            if !observed.insert(filter.column) {
                return Err(filter_error(
                    "use.office.spreadsheet_filter_column_duplicate",
                    format!(
                        "Spreadsheet filter column {} is defined more than once.",
                        filter.column
                    ),
                ));
            }
            filter.criteria.validate()?;
            text_bytes = text_bytes.saturating_add(filter.criteria.text_bytes());
        }
        if text_bytes > MAX_FILTER_TEXT_BYTES {
            return Err(filter_error(
                "use.office.spreadsheet_filter_text_limit",
                format!(
                    "Spreadsheet filter text uses {text_bytes} bytes; the limit is {MAX_FILTER_TEXT_BYTES}."
                ),
            ));
        }
        Ok(())
    }
}

fn validate_filter_text(value: &str, label: &str, allow_empty: bool) -> UseResult<()> {
    if (!allow_empty && value.is_empty()) || value.chars().count() > MAX_FILTER_VALUE_CHARACTERS {
        return Err(filter_error(
            "use.office.spreadsheet_filter_value_invalid",
            format!("Spreadsheet {label} must contain 1-{MAX_FILTER_VALUE_CHARACTERS} characters."),
        ));
    }
    if let Some(character) = value.chars().find(|character| {
        !matches!(
            u32::from(*character),
            0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF
        )
    }) {
        return Err(filter_error(
            "use.office.spreadsheet_filter_value_invalid",
            format!(
                "Spreadsheet {label} contains XML-forbidden character U+{:04X}.",
                u32::from(character)
            ),
        ));
    }
    Ok(())
}

fn filter_error(code: &str, message: impl Into<String>) -> a3s_use_core::UseError {
    super::super::editor_error(code, message)
}
