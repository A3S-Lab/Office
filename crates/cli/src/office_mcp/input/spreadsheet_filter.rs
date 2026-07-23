use a3s_office::{
    NativeSpreadsheetAutoFilter, NativeSpreadsheetDynamicFilter, NativeSpreadsheetFilterColumn,
    NativeSpreadsheetFilterCriteria,
};
use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Copy, Serialize, Deserialize, schemars::JsonSchema)]
#[serde(rename_all = "kebab-case")]
enum OfficeSpreadsheetDynamicFilter {
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

impl From<OfficeSpreadsheetDynamicFilter> for NativeSpreadsheetDynamicFilter {
    fn from(value: OfficeSpreadsheetDynamicFilter) -> Self {
        match value {
            OfficeSpreadsheetDynamicFilter::AboveAverage => Self::AboveAverage,
            OfficeSpreadsheetDynamicFilter::BelowAverage => Self::BelowAverage,
            OfficeSpreadsheetDynamicFilter::Tomorrow => Self::Tomorrow,
            OfficeSpreadsheetDynamicFilter::Today => Self::Today,
            OfficeSpreadsheetDynamicFilter::Yesterday => Self::Yesterday,
            OfficeSpreadsheetDynamicFilter::NextWeek => Self::NextWeek,
            OfficeSpreadsheetDynamicFilter::ThisWeek => Self::ThisWeek,
            OfficeSpreadsheetDynamicFilter::LastWeek => Self::LastWeek,
            OfficeSpreadsheetDynamicFilter::NextMonth => Self::NextMonth,
            OfficeSpreadsheetDynamicFilter::ThisMonth => Self::ThisMonth,
            OfficeSpreadsheetDynamicFilter::LastMonth => Self::LastMonth,
            OfficeSpreadsheetDynamicFilter::NextQuarter => Self::NextQuarter,
            OfficeSpreadsheetDynamicFilter::ThisQuarter => Self::ThisQuarter,
            OfficeSpreadsheetDynamicFilter::LastQuarter => Self::LastQuarter,
            OfficeSpreadsheetDynamicFilter::NextYear => Self::NextYear,
            OfficeSpreadsheetDynamicFilter::ThisYear => Self::ThisYear,
            OfficeSpreadsheetDynamicFilter::LastYear => Self::LastYear,
            OfficeSpreadsheetDynamicFilter::YearToDate => Self::YearToDate,
            OfficeSpreadsheetDynamicFilter::Quarter1 => Self::Quarter1,
            OfficeSpreadsheetDynamicFilter::Quarter2 => Self::Quarter2,
            OfficeSpreadsheetDynamicFilter::Quarter3 => Self::Quarter3,
            OfficeSpreadsheetDynamicFilter::Quarter4 => Self::Quarter4,
            OfficeSpreadsheetDynamicFilter::Month1 => Self::Month1,
            OfficeSpreadsheetDynamicFilter::Month2 => Self::Month2,
            OfficeSpreadsheetDynamicFilter::Month3 => Self::Month3,
            OfficeSpreadsheetDynamicFilter::Month4 => Self::Month4,
            OfficeSpreadsheetDynamicFilter::Month5 => Self::Month5,
            OfficeSpreadsheetDynamicFilter::Month6 => Self::Month6,
            OfficeSpreadsheetDynamicFilter::Month7 => Self::Month7,
            OfficeSpreadsheetDynamicFilter::Month8 => Self::Month8,
            OfficeSpreadsheetDynamicFilter::Month9 => Self::Month9,
            OfficeSpreadsheetDynamicFilter::Month10 => Self::Month10,
            OfficeSpreadsheetDynamicFilter::Month11 => Self::Month11,
            OfficeSpreadsheetDynamicFilter::Month12 => Self::Month12,
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, schemars::JsonSchema)]
#[serde(
    tag = "type",
    rename_all = "kebab-case",
    rename_all_fields = "camelCase",
    deny_unknown_fields
)]
enum OfficeSpreadsheetFilterCriteria {
    Values {
        /// Exact display values accepted by the filter.
        values: Vec<String>,
        /// Also include blank cells.
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
        /// Number of largest values, from 1 through 500.
        count: u16,
    },
    TopPercent {
        /// Percentage of largest values, from 1 through 100.
        percent: u8,
    },
    Bottom {
        /// Number of smallest values, from 1 through 500.
        count: u16,
    },
    BottomPercent {
        /// Percentage of smallest values, from 1 through 100.
        percent: u8,
    },
    Dynamic {
        kind: OfficeSpreadsheetDynamicFilter,
    },
}

impl From<OfficeSpreadsheetFilterCriteria> for NativeSpreadsheetFilterCriteria {
    fn from(value: OfficeSpreadsheetFilterCriteria) -> Self {
        match value {
            OfficeSpreadsheetFilterCriteria::Values {
                values,
                include_blanks,
            } => Self::Values {
                values,
                include_blanks,
            },
            OfficeSpreadsheetFilterCriteria::Equals { value } => Self::Equals { value },
            OfficeSpreadsheetFilterCriteria::NotEquals { value } => Self::NotEquals { value },
            OfficeSpreadsheetFilterCriteria::Contains { value } => Self::Contains { value },
            OfficeSpreadsheetFilterCriteria::DoesNotContain { value } => {
                Self::DoesNotContain { value }
            }
            OfficeSpreadsheetFilterCriteria::BeginsWith { value } => Self::BeginsWith { value },
            OfficeSpreadsheetFilterCriteria::EndsWith { value } => Self::EndsWith { value },
            OfficeSpreadsheetFilterCriteria::GreaterThan { value } => Self::GreaterThan { value },
            OfficeSpreadsheetFilterCriteria::GreaterThanOrEqual { value } => {
                Self::GreaterThanOrEqual { value }
            }
            OfficeSpreadsheetFilterCriteria::LessThan { value } => Self::LessThan { value },
            OfficeSpreadsheetFilterCriteria::LessThanOrEqual { value } => {
                Self::LessThanOrEqual { value }
            }
            OfficeSpreadsheetFilterCriteria::Between { lower, upper } => {
                Self::Between { lower, upper }
            }
            OfficeSpreadsheetFilterCriteria::NotBetween { lower, upper } => {
                Self::NotBetween { lower, upper }
            }
            OfficeSpreadsheetFilterCriteria::Blanks => Self::Blanks,
            OfficeSpreadsheetFilterCriteria::NonBlanks => Self::NonBlanks,
            OfficeSpreadsheetFilterCriteria::Top { count } => Self::Top { count },
            OfficeSpreadsheetFilterCriteria::TopPercent { percent } => Self::TopPercent { percent },
            OfficeSpreadsheetFilterCriteria::Bottom { count } => Self::Bottom { count },
            OfficeSpreadsheetFilterCriteria::BottomPercent { percent } => {
                Self::BottomPercent { percent }
            }
            OfficeSpreadsheetFilterCriteria::Dynamic { kind } => {
                Self::Dynamic { kind: kind.into() }
            }
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, schemars::JsonSchema)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub(in crate::office_mcp) struct OfficeSpreadsheetFilterColumn {
    /// Zero-based column offset inside the AutoFilter range.
    column: u32,
    criteria: OfficeSpreadsheetFilterCriteria,
}

impl From<OfficeSpreadsheetFilterColumn> for NativeSpreadsheetFilterColumn {
    fn from(value: OfficeSpreadsheetFilterColumn) -> Self {
        Self {
            column: value.column,
            criteria: value.criteria.into(),
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, schemars::JsonSchema)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub(in crate::office_mcp) struct OfficeSpreadsheetAutoFilter {
    /// Rectangular A1 range containing the header and filtered data rows.
    range: String,
    /// At most one criterion per zero-based range column.
    #[serde(default)]
    columns: Vec<OfficeSpreadsheetFilterColumn>,
}

impl From<OfficeSpreadsheetAutoFilter> for NativeSpreadsheetAutoFilter {
    fn from(value: OfficeSpreadsheetAutoFilter) -> Self {
        Self {
            range: value.range,
            columns: value.columns.into_iter().map(Into::into).collect(),
        }
    }
}
