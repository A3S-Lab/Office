use a3s_office::{
    NativeOfficeRgbColor, NativeSpreadsheetConditionalFormat,
    NativeSpreadsheetConditionalFormatIconSet, NativeSpreadsheetConditionalFormatOperator,
    NativeSpreadsheetConditionalFormatRule, NativeSpreadsheetConditionalFormatThreshold,
    NativeSpreadsheetConditionalFormatThresholdKind, NativeSpreadsheetConditionalFormatTimePeriod,
    NativeSpreadsheetDifferentialFormat,
};
use serde::{Deserialize, Serialize};

use super::OfficeRgbColor;

#[derive(Debug, Clone, Copy, Serialize, Deserialize, schemars::JsonSchema)]
#[serde(rename_all = "camelCase")]
enum OfficeConditionalFormatOperator {
    Between,
    NotBetween,
    Equal,
    NotEqual,
    GreaterThan,
    GreaterThanOrEqual,
    LessThan,
    LessThanOrEqual,
}

impl From<OfficeConditionalFormatOperator> for NativeSpreadsheetConditionalFormatOperator {
    fn from(value: OfficeConditionalFormatOperator) -> Self {
        match value {
            OfficeConditionalFormatOperator::Between => Self::Between,
            OfficeConditionalFormatOperator::NotBetween => Self::NotBetween,
            OfficeConditionalFormatOperator::Equal => Self::Equal,
            OfficeConditionalFormatOperator::NotEqual => Self::NotEqual,
            OfficeConditionalFormatOperator::GreaterThan => Self::GreaterThan,
            OfficeConditionalFormatOperator::GreaterThanOrEqual => Self::GreaterThanOrEqual,
            OfficeConditionalFormatOperator::LessThan => Self::LessThan,
            OfficeConditionalFormatOperator::LessThanOrEqual => Self::LessThanOrEqual,
        }
    }
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, schemars::JsonSchema)]
#[serde(rename_all = "camelCase")]
enum OfficeConditionalFormatTimePeriod {
    Today,
    Yesterday,
    Tomorrow,
    Last7Days,
    ThisWeek,
    LastWeek,
    NextWeek,
    ThisMonth,
    LastMonth,
    NextMonth,
}

impl From<OfficeConditionalFormatTimePeriod> for NativeSpreadsheetConditionalFormatTimePeriod {
    fn from(value: OfficeConditionalFormatTimePeriod) -> Self {
        match value {
            OfficeConditionalFormatTimePeriod::Today => Self::Today,
            OfficeConditionalFormatTimePeriod::Yesterday => Self::Yesterday,
            OfficeConditionalFormatTimePeriod::Tomorrow => Self::Tomorrow,
            OfficeConditionalFormatTimePeriod::Last7Days => Self::Last7Days,
            OfficeConditionalFormatTimePeriod::ThisWeek => Self::ThisWeek,
            OfficeConditionalFormatTimePeriod::LastWeek => Self::LastWeek,
            OfficeConditionalFormatTimePeriod::NextWeek => Self::NextWeek,
            OfficeConditionalFormatTimePeriod::ThisMonth => Self::ThisMonth,
            OfficeConditionalFormatTimePeriod::LastMonth => Self::LastMonth,
            OfficeConditionalFormatTimePeriod::NextMonth => Self::NextMonth,
        }
    }
}

#[derive(Debug, Default, Clone, Copy, Serialize, Deserialize, schemars::JsonSchema)]
enum OfficeConditionalFormatIconSet {
    #[serde(rename = "3Arrows")]
    ThreeArrows,
    #[serde(rename = "3ArrowsGray")]
    ThreeArrowsGray,
    #[serde(rename = "3Flags")]
    ThreeFlags,
    #[default]
    #[serde(rename = "3TrafficLights1")]
    ThreeTrafficLights1,
    #[serde(rename = "3TrafficLights2")]
    ThreeTrafficLights2,
    #[serde(rename = "3Signs")]
    ThreeSigns,
    #[serde(rename = "3Symbols")]
    ThreeSymbols,
    #[serde(rename = "3Symbols2")]
    ThreeSymbols2,
    #[serde(rename = "4Arrows")]
    FourArrows,
    #[serde(rename = "4ArrowsGray")]
    FourArrowsGray,
    #[serde(rename = "4RedToBlack")]
    FourRedToBlack,
    #[serde(rename = "4Rating")]
    FourRating,
    #[serde(rename = "4TrafficLights")]
    FourTrafficLights,
    #[serde(rename = "5Arrows")]
    FiveArrows,
    #[serde(rename = "5ArrowsGray")]
    FiveArrowsGray,
    #[serde(rename = "5Rating")]
    FiveRating,
    #[serde(rename = "5Quarters")]
    FiveQuarters,
}

impl From<OfficeConditionalFormatIconSet> for NativeSpreadsheetConditionalFormatIconSet {
    fn from(value: OfficeConditionalFormatIconSet) -> Self {
        match value {
            OfficeConditionalFormatIconSet::ThreeArrows => Self::ThreeArrows,
            OfficeConditionalFormatIconSet::ThreeArrowsGray => Self::ThreeArrowsGray,
            OfficeConditionalFormatIconSet::ThreeFlags => Self::ThreeFlags,
            OfficeConditionalFormatIconSet::ThreeTrafficLights1 => Self::ThreeTrafficLights1,
            OfficeConditionalFormatIconSet::ThreeTrafficLights2 => Self::ThreeTrafficLights2,
            OfficeConditionalFormatIconSet::ThreeSigns => Self::ThreeSigns,
            OfficeConditionalFormatIconSet::ThreeSymbols => Self::ThreeSymbols,
            OfficeConditionalFormatIconSet::ThreeSymbols2 => Self::ThreeSymbols2,
            OfficeConditionalFormatIconSet::FourArrows => Self::FourArrows,
            OfficeConditionalFormatIconSet::FourArrowsGray => Self::FourArrowsGray,
            OfficeConditionalFormatIconSet::FourRedToBlack => Self::FourRedToBlack,
            OfficeConditionalFormatIconSet::FourRating => Self::FourRating,
            OfficeConditionalFormatIconSet::FourTrafficLights => Self::FourTrafficLights,
            OfficeConditionalFormatIconSet::FiveArrows => Self::FiveArrows,
            OfficeConditionalFormatIconSet::FiveArrowsGray => Self::FiveArrowsGray,
            OfficeConditionalFormatIconSet::FiveRating => Self::FiveRating,
            OfficeConditionalFormatIconSet::FiveQuarters => Self::FiveQuarters,
        }
    }
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, schemars::JsonSchema)]
#[serde(rename_all = "camelCase")]
enum OfficeConditionalFormatThresholdKind {
    Min,
    Max,
    Number,
    Percent,
    Percentile,
    Formula,
}

impl From<OfficeConditionalFormatThresholdKind>
    for NativeSpreadsheetConditionalFormatThresholdKind
{
    fn from(value: OfficeConditionalFormatThresholdKind) -> Self {
        match value {
            OfficeConditionalFormatThresholdKind::Min => Self::Min,
            OfficeConditionalFormatThresholdKind::Max => Self::Max,
            OfficeConditionalFormatThresholdKind::Number => Self::Number,
            OfficeConditionalFormatThresholdKind::Percent => Self::Percent,
            OfficeConditionalFormatThresholdKind::Percentile => Self::Percentile,
            OfficeConditionalFormatThresholdKind::Formula => Self::Formula,
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, schemars::JsonSchema)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct OfficeConditionalFormatThreshold {
    /// Min, max, number, percent, percentile, or formula threshold.
    kind: OfficeConditionalFormatThresholdKind,
    /// Scalar or formula value. Omit only for min and max thresholds.
    value: Option<String>,
}

impl From<OfficeConditionalFormatThreshold> for NativeSpreadsheetConditionalFormatThreshold {
    fn from(value: OfficeConditionalFormatThreshold) -> Self {
        Self {
            kind: value.kind.into(),
            value: value.value,
        }
    }
}

fn minimum_threshold() -> OfficeConditionalFormatThreshold {
    OfficeConditionalFormatThreshold {
        kind: OfficeConditionalFormatThresholdKind::Min,
        value: None,
    }
}

fn maximum_threshold() -> OfficeConditionalFormatThreshold {
    OfficeConditionalFormatThreshold {
        kind: OfficeConditionalFormatThresholdKind::Max,
        value: None,
    }
}

#[derive(Debug, Default, Clone, Copy, Serialize, Deserialize, schemars::JsonSchema)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct OfficeDifferentialFormat {
    /// Solid 24-bit RGB fill applied when the rule matches.
    fill: Option<OfficeRgbColor>,
    /// Optional 24-bit RGB font color.
    font_color: Option<OfficeRgbColor>,
    /// Optional explicit bold state.
    bold: Option<bool>,
}

impl From<OfficeDifferentialFormat> for NativeSpreadsheetDifferentialFormat {
    fn from(value: OfficeDifferentialFormat) -> Self {
        Self {
            fill: value.fill.map(native_color),
            font_color: value.font_color.map(native_color),
            bold: value.bold,
        }
    }
}

/// Closed native Spreadsheet conditional-format rule families.
#[derive(Debug, Clone, Serialize, Deserialize, schemars::JsonSchema)]
#[serde(
    tag = "type",
    rename_all = "camelCase",
    rename_all_fields = "camelCase",
    deny_unknown_fields
)]
enum OfficeConditionalFormatRule {
    CellIs {
        operator: OfficeConditionalFormatOperator,
        formula1: String,
        formula2: Option<String>,
        #[serde(default)]
        format: OfficeDifferentialFormat,
    },
    Formula {
        formula: String,
        #[serde(default)]
        format: OfficeDifferentialFormat,
    },
    ContainsText {
        text: String,
        #[serde(default)]
        format: OfficeDifferentialFormat,
    },
    NotContainsText {
        text: String,
        #[serde(default)]
        format: OfficeDifferentialFormat,
    },
    BeginsWith {
        text: String,
        #[serde(default)]
        format: OfficeDifferentialFormat,
    },
    EndsWith {
        text: String,
        #[serde(default)]
        format: OfficeDifferentialFormat,
    },
    Top {
        rank: u32,
        #[serde(default)]
        percent: bool,
        #[serde(default)]
        bottom: bool,
        #[serde(default)]
        format: OfficeDifferentialFormat,
    },
    AboveAverage {
        #[serde(default = "default_true")]
        above: bool,
        #[serde(default)]
        equal: bool,
        standard_deviations: Option<u32>,
        #[serde(default)]
        format: OfficeDifferentialFormat,
    },
    DuplicateValues {
        #[serde(default)]
        format: OfficeDifferentialFormat,
    },
    UniqueValues {
        #[serde(default)]
        format: OfficeDifferentialFormat,
    },
    ContainsBlanks {
        #[serde(default)]
        format: OfficeDifferentialFormat,
    },
    NotContainsBlanks {
        #[serde(default)]
        format: OfficeDifferentialFormat,
    },
    ContainsErrors {
        #[serde(default)]
        format: OfficeDifferentialFormat,
    },
    NotContainsErrors {
        #[serde(default)]
        format: OfficeDifferentialFormat,
    },
    TimePeriod {
        period: OfficeConditionalFormatTimePeriod,
        #[serde(default)]
        format: OfficeDifferentialFormat,
    },
    DataBar {
        color: OfficeRgbColor,
        #[serde(default = "minimum_threshold")]
        min: OfficeConditionalFormatThreshold,
        #[serde(default = "maximum_threshold")]
        max: OfficeConditionalFormatThreshold,
        #[serde(default = "default_true")]
        show_value: bool,
        min_length: Option<u8>,
        max_length: Option<u8>,
    },
    ColorScale {
        min: OfficeConditionalFormatThreshold,
        min_color: OfficeRgbColor,
        mid: Option<OfficeConditionalFormatThreshold>,
        mid_color: Option<OfficeRgbColor>,
        max: OfficeConditionalFormatThreshold,
        max_color: OfficeRgbColor,
    },
    IconSet {
        #[serde(default)]
        icon_set: OfficeConditionalFormatIconSet,
        #[serde(default)]
        thresholds: Vec<OfficeConditionalFormatThreshold>,
        #[serde(default)]
        reverse: bool,
        #[serde(default = "default_true")]
        show_value: bool,
    },
}

impl From<OfficeConditionalFormatRule> for NativeSpreadsheetConditionalFormatRule {
    fn from(value: OfficeConditionalFormatRule) -> Self {
        match value {
            OfficeConditionalFormatRule::CellIs {
                operator,
                formula1,
                formula2,
                format,
            } => Self::CellIs {
                operator: operator.into(),
                formula1,
                formula2,
                format: format.into(),
            },
            OfficeConditionalFormatRule::Formula { formula, format } => Self::Formula {
                formula,
                format: format.into(),
            },
            OfficeConditionalFormatRule::ContainsText { text, format } => Self::ContainsText {
                text,
                format: format.into(),
            },
            OfficeConditionalFormatRule::NotContainsText { text, format } => {
                Self::NotContainsText {
                    text,
                    format: format.into(),
                }
            }
            OfficeConditionalFormatRule::BeginsWith { text, format } => Self::BeginsWith {
                text,
                format: format.into(),
            },
            OfficeConditionalFormatRule::EndsWith { text, format } => Self::EndsWith {
                text,
                format: format.into(),
            },
            OfficeConditionalFormatRule::Top {
                rank,
                percent,
                bottom,
                format,
            } => Self::Top {
                rank,
                percent,
                bottom,
                format: format.into(),
            },
            OfficeConditionalFormatRule::AboveAverage {
                above,
                equal,
                standard_deviations,
                format,
            } => Self::AboveAverage {
                above,
                equal,
                standard_deviations,
                format: format.into(),
            },
            OfficeConditionalFormatRule::DuplicateValues { format } => Self::DuplicateValues {
                format: format.into(),
            },
            OfficeConditionalFormatRule::UniqueValues { format } => Self::UniqueValues {
                format: format.into(),
            },
            OfficeConditionalFormatRule::ContainsBlanks { format } => Self::ContainsBlanks {
                format: format.into(),
            },
            OfficeConditionalFormatRule::NotContainsBlanks { format } => Self::NotContainsBlanks {
                format: format.into(),
            },
            OfficeConditionalFormatRule::ContainsErrors { format } => Self::ContainsErrors {
                format: format.into(),
            },
            OfficeConditionalFormatRule::NotContainsErrors { format } => Self::NotContainsErrors {
                format: format.into(),
            },
            OfficeConditionalFormatRule::TimePeriod { period, format } => Self::TimePeriod {
                period: period.into(),
                format: format.into(),
            },
            OfficeConditionalFormatRule::DataBar {
                color,
                min,
                max,
                show_value,
                min_length,
                max_length,
            } => Self::DataBar {
                color: native_color(color),
                min: min.into(),
                max: max.into(),
                show_value,
                min_length,
                max_length,
            },
            OfficeConditionalFormatRule::ColorScale {
                min,
                min_color,
                mid,
                mid_color,
                max,
                max_color,
            } => Self::ColorScale {
                min: min.into(),
                min_color: native_color(min_color),
                mid: mid.map(Into::into),
                mid_color: mid_color.map(native_color),
                max: max.into(),
                max_color: native_color(max_color),
            },
            OfficeConditionalFormatRule::IconSet {
                icon_set,
                thresholds,
                reverse,
                show_value,
            } => Self::IconSet {
                icon_set: icon_set.into(),
                thresholds: thresholds.into_iter().map(Into::into).collect(),
                reverse,
                show_value,
            },
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, schemars::JsonSchema)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub(in crate::office_mcp) struct OfficeConditionalFormat {
    /// One or more disjoint rectangular A1 ranges on the target worksheet.
    ranges: Vec<String>,
    /// Stop evaluating later rules when this rule matches.
    #[serde(default)]
    stop_if_true: bool,
    rule: OfficeConditionalFormatRule,
}

impl From<OfficeConditionalFormat> for NativeSpreadsheetConditionalFormat {
    fn from(value: OfficeConditionalFormat) -> Self {
        Self {
            ranges: value.ranges,
            stop_if_true: value.stop_if_true,
            rule: value.rule.into(),
        }
    }
}

const fn default_true() -> bool {
    true
}

fn native_color(value: OfficeRgbColor) -> NativeOfficeRgbColor {
    NativeOfficeRgbColor::new(value.red, value.green, value.blue)
}
