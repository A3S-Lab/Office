use serde::{Deserialize, Serialize};

use super::NativeOfficeRgbColor;

/// Comparison operators supported by Spreadsheet conditional-format rules.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum NativeSpreadsheetConditionalFormatOperator {
    Between,
    NotBetween,
    Equal,
    NotEqual,
    GreaterThan,
    GreaterThanOrEqual,
    LessThan,
    LessThanOrEqual,
}

/// Date windows understood by SpreadsheetML `timePeriod` rules.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum NativeSpreadsheetConditionalFormatTimePeriod {
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

/// Standard legacy SpreadsheetML icon sets.
#[derive(Debug, Default, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
pub enum NativeSpreadsheetConditionalFormatIconSet {
    #[serde(rename = "3Arrows")]
    ThreeArrows,
    #[serde(rename = "3ArrowsGray")]
    ThreeArrowsGray,
    #[serde(rename = "3Flags")]
    ThreeFlags,
    #[serde(rename = "3TrafficLights1")]
    #[default]
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

impl NativeSpreadsheetConditionalFormatIconSet {
    pub(crate) const fn token(self) -> &'static str {
        match self {
            Self::ThreeArrows => "3Arrows",
            Self::ThreeArrowsGray => "3ArrowsGray",
            Self::ThreeFlags => "3Flags",
            Self::ThreeTrafficLights1 => "3TrafficLights1",
            Self::ThreeTrafficLights2 => "3TrafficLights2",
            Self::ThreeSigns => "3Signs",
            Self::ThreeSymbols => "3Symbols",
            Self::ThreeSymbols2 => "3Symbols2",
            Self::FourArrows => "4Arrows",
            Self::FourArrowsGray => "4ArrowsGray",
            Self::FourRedToBlack => "4RedToBlack",
            Self::FourRating => "4Rating",
            Self::FourTrafficLights => "4TrafficLights",
            Self::FiveArrows => "5Arrows",
            Self::FiveArrowsGray => "5ArrowsGray",
            Self::FiveRating => "5Rating",
            Self::FiveQuarters => "5Quarters",
        }
    }

    pub(crate) const fn icon_count(self) -> usize {
        match self {
            Self::ThreeArrows
            | Self::ThreeArrowsGray
            | Self::ThreeFlags
            | Self::ThreeTrafficLights1
            | Self::ThreeTrafficLights2
            | Self::ThreeSigns
            | Self::ThreeSymbols
            | Self::ThreeSymbols2 => 3,
            Self::FourArrows
            | Self::FourArrowsGray
            | Self::FourRedToBlack
            | Self::FourRating
            | Self::FourTrafficLights => 4,
            Self::FiveArrows | Self::FiveArrowsGray | Self::FiveRating | Self::FiveQuarters => 5,
        }
    }
}

/// SpreadsheetML conditional-format threshold kind.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum NativeSpreadsheetConditionalFormatThresholdKind {
    Min,
    Max,
    Number,
    Percent,
    Percentile,
    Formula,
}

/// One typed threshold used by data bars, color scales, and icon sets.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeSpreadsheetConditionalFormatThreshold {
    pub kind: NativeSpreadsheetConditionalFormatThresholdKind,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub value: Option<String>,
}

impl NativeSpreadsheetConditionalFormatThreshold {
    pub fn min() -> Self {
        Self {
            kind: NativeSpreadsheetConditionalFormatThresholdKind::Min,
            value: None,
        }
    }

    pub fn max() -> Self {
        Self {
            kind: NativeSpreadsheetConditionalFormatThresholdKind::Max,
            value: None,
        }
    }

    pub fn number(value: impl Into<String>) -> Self {
        Self {
            kind: NativeSpreadsheetConditionalFormatThresholdKind::Number,
            value: Some(value.into()),
        }
    }

    pub fn percentile(value: impl Into<String>) -> Self {
        Self {
            kind: NativeSpreadsheetConditionalFormatThresholdKind::Percentile,
            value: Some(value.into()),
        }
    }
}

/// Differential formatting applied when a classic rule matches.
///
/// This intentionally mirrors the currently documented OfficeCLI core: a
/// solid fill plus optional font color and bold state. Other differential
/// formatting properties remain owned by future typed extensions.
#[derive(Debug, Default, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeSpreadsheetDifferentialFormat {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill: Option<NativeOfficeRgbColor>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_color: Option<NativeOfficeRgbColor>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
}

impl NativeSpreadsheetDifferentialFormat {
    pub fn is_empty(&self) -> bool {
        self.fill.is_none() && self.font_color.is_none() && self.bold.is_none()
    }

    pub fn with_fill(mut self, color: NativeOfficeRgbColor) -> Self {
        self.fill = Some(color);
        self
    }

    pub fn with_font_color(mut self, color: NativeOfficeRgbColor) -> Self {
        self.font_color = Some(color);
        self
    }

    pub fn with_bold(mut self, value: bool) -> Self {
        self.bold = Some(value);
        self
    }
}

/// Closed Spreadsheet conditional-format rule families supported natively.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(
    tag = "type",
    rename_all = "camelCase",
    rename_all_fields = "camelCase",
    deny_unknown_fields
)]
pub enum NativeSpreadsheetConditionalFormatRule {
    CellIs {
        operator: NativeSpreadsheetConditionalFormatOperator,
        formula1: String,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        formula2: Option<String>,
        #[serde(
            default,
            skip_serializing_if = "NativeSpreadsheetDifferentialFormat::is_empty"
        )]
        format: NativeSpreadsheetDifferentialFormat,
    },
    Formula {
        formula: String,
        #[serde(
            default,
            skip_serializing_if = "NativeSpreadsheetDifferentialFormat::is_empty"
        )]
        format: NativeSpreadsheetDifferentialFormat,
    },
    ContainsText {
        text: String,
        #[serde(
            default,
            skip_serializing_if = "NativeSpreadsheetDifferentialFormat::is_empty"
        )]
        format: NativeSpreadsheetDifferentialFormat,
    },
    NotContainsText {
        text: String,
        #[serde(
            default,
            skip_serializing_if = "NativeSpreadsheetDifferentialFormat::is_empty"
        )]
        format: NativeSpreadsheetDifferentialFormat,
    },
    BeginsWith {
        text: String,
        #[serde(
            default,
            skip_serializing_if = "NativeSpreadsheetDifferentialFormat::is_empty"
        )]
        format: NativeSpreadsheetDifferentialFormat,
    },
    EndsWith {
        text: String,
        #[serde(
            default,
            skip_serializing_if = "NativeSpreadsheetDifferentialFormat::is_empty"
        )]
        format: NativeSpreadsheetDifferentialFormat,
    },
    Top {
        rank: u32,
        #[serde(default, skip_serializing_if = "is_false")]
        percent: bool,
        #[serde(default, skip_serializing_if = "is_false")]
        bottom: bool,
        #[serde(
            default,
            skip_serializing_if = "NativeSpreadsheetDifferentialFormat::is_empty"
        )]
        format: NativeSpreadsheetDifferentialFormat,
    },
    AboveAverage {
        #[serde(default = "default_true", skip_serializing_if = "is_true")]
        above: bool,
        #[serde(default, skip_serializing_if = "is_false")]
        equal: bool,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        standard_deviations: Option<u32>,
        #[serde(
            default,
            skip_serializing_if = "NativeSpreadsheetDifferentialFormat::is_empty"
        )]
        format: NativeSpreadsheetDifferentialFormat,
    },
    DuplicateValues {
        #[serde(
            default,
            skip_serializing_if = "NativeSpreadsheetDifferentialFormat::is_empty"
        )]
        format: NativeSpreadsheetDifferentialFormat,
    },
    UniqueValues {
        #[serde(
            default,
            skip_serializing_if = "NativeSpreadsheetDifferentialFormat::is_empty"
        )]
        format: NativeSpreadsheetDifferentialFormat,
    },
    ContainsBlanks {
        #[serde(
            default,
            skip_serializing_if = "NativeSpreadsheetDifferentialFormat::is_empty"
        )]
        format: NativeSpreadsheetDifferentialFormat,
    },
    NotContainsBlanks {
        #[serde(
            default,
            skip_serializing_if = "NativeSpreadsheetDifferentialFormat::is_empty"
        )]
        format: NativeSpreadsheetDifferentialFormat,
    },
    ContainsErrors {
        #[serde(
            default,
            skip_serializing_if = "NativeSpreadsheetDifferentialFormat::is_empty"
        )]
        format: NativeSpreadsheetDifferentialFormat,
    },
    NotContainsErrors {
        #[serde(
            default,
            skip_serializing_if = "NativeSpreadsheetDifferentialFormat::is_empty"
        )]
        format: NativeSpreadsheetDifferentialFormat,
    },
    TimePeriod {
        period: NativeSpreadsheetConditionalFormatTimePeriod,
        #[serde(
            default,
            skip_serializing_if = "NativeSpreadsheetDifferentialFormat::is_empty"
        )]
        format: NativeSpreadsheetDifferentialFormat,
    },
    DataBar {
        color: NativeOfficeRgbColor,
        #[serde(default = "NativeSpreadsheetConditionalFormatThreshold::min")]
        min: NativeSpreadsheetConditionalFormatThreshold,
        #[serde(default = "NativeSpreadsheetConditionalFormatThreshold::max")]
        max: NativeSpreadsheetConditionalFormatThreshold,
        #[serde(default = "default_true", skip_serializing_if = "is_true")]
        show_value: bool,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        min_length: Option<u8>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        max_length: Option<u8>,
    },
    ColorScale {
        min: NativeSpreadsheetConditionalFormatThreshold,
        min_color: NativeOfficeRgbColor,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        mid: Option<NativeSpreadsheetConditionalFormatThreshold>,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        mid_color: Option<NativeOfficeRgbColor>,
        max: NativeSpreadsheetConditionalFormatThreshold,
        max_color: NativeOfficeRgbColor,
    },
    IconSet {
        #[serde(default)]
        icon_set: NativeSpreadsheetConditionalFormatIconSet,
        #[serde(default, skip_serializing_if = "Vec::is_empty")]
        thresholds: Vec<NativeSpreadsheetConditionalFormatThreshold>,
        #[serde(default, skip_serializing_if = "is_false")]
        reverse: bool,
        #[serde(default = "default_true", skip_serializing_if = "is_true")]
        show_value: bool,
    },
}

impl NativeSpreadsheetConditionalFormatRule {
    pub(crate) fn differential_format(&self) -> Option<&NativeSpreadsheetDifferentialFormat> {
        match self {
            Self::CellIs { format, .. }
            | Self::Formula { format, .. }
            | Self::ContainsText { format, .. }
            | Self::NotContainsText { format, .. }
            | Self::BeginsWith { format, .. }
            | Self::EndsWith { format, .. }
            | Self::Top { format, .. }
            | Self::AboveAverage { format, .. }
            | Self::DuplicateValues { format }
            | Self::UniqueValues { format }
            | Self::ContainsBlanks { format }
            | Self::NotContainsBlanks { format }
            | Self::ContainsErrors { format }
            | Self::NotContainsErrors { format }
            | Self::TimePeriod { format, .. } => Some(format),
            Self::DataBar { .. } | Self::ColorScale { .. } | Self::IconSet { .. } => None,
        }
    }
}

/// One complete typed Spreadsheet conditional-formatting rule.
#[derive(Debug, Clone, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeSpreadsheetConditionalFormat {
    pub ranges: Vec<String>,
    #[serde(default, skip_serializing_if = "is_false")]
    pub stop_if_true: bool,
    pub rule: NativeSpreadsheetConditionalFormatRule,
}

impl NativeSpreadsheetConditionalFormat {
    pub fn new(range: impl Into<String>, rule: NativeSpreadsheetConditionalFormatRule) -> Self {
        Self {
            ranges: vec![range.into()],
            stop_if_true: false,
            rule,
        }
    }

    pub fn with_range(mut self, range: impl Into<String>) -> Self {
        self.ranges.push(range.into());
        self
    }

    pub fn with_stop_if_true(mut self, stop_if_true: bool) -> Self {
        self.stop_if_true = stop_if_true;
        self
    }
}

const fn default_true() -> bool {
    true
}

const fn is_true(value: &bool) -> bool {
    *value
}

const fn is_false(value: &bool) -> bool {
    !*value
}
