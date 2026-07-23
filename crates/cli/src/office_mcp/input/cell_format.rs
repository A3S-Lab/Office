use a3s_office::{
    NativeOfficeRgbColor, NativeSpreadsheetBorder, NativeSpreadsheetBorderLine,
    NativeSpreadsheetBorderStyle, NativeSpreadsheetCellFormat, NativeSpreadsheetFill,
    NativeSpreadsheetReadingOrder, NativeSpreadsheetVerticalAlignment,
};
use serde::{Deserialize, Serialize};

use super::OfficeRgbColor;

#[derive(Debug, Clone, Copy, Serialize, Deserialize, schemars::JsonSchema)]
#[serde(tag = "kind", rename_all = "kebab-case", deny_unknown_fields)]
enum OfficeCellFill {
    /// Explicitly remove the cell fill.
    None,
    /// Apply a solid 24-bit RGB fill.
    Solid { color: OfficeRgbColor },
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, schemars::JsonSchema)]
#[serde(rename_all = "camelCase")]
enum OfficeCellBorderStyle {
    Thin,
    Medium,
    Thick,
    Double,
    Dashed,
    Dotted,
    DashDot,
    DashDotDot,
    Hair,
    MediumDashed,
    MediumDashDot,
    MediumDashDotDot,
    SlantDashDot,
}

impl From<OfficeCellBorderStyle> for NativeSpreadsheetBorderStyle {
    fn from(value: OfficeCellBorderStyle) -> Self {
        match value {
            OfficeCellBorderStyle::Thin => Self::Thin,
            OfficeCellBorderStyle::Medium => Self::Medium,
            OfficeCellBorderStyle::Thick => Self::Thick,
            OfficeCellBorderStyle::Double => Self::Double,
            OfficeCellBorderStyle::Dashed => Self::Dashed,
            OfficeCellBorderStyle::Dotted => Self::Dotted,
            OfficeCellBorderStyle::DashDot => Self::DashDot,
            OfficeCellBorderStyle::DashDotDot => Self::DashDotDot,
            OfficeCellBorderStyle::Hair => Self::Hair,
            OfficeCellBorderStyle::MediumDashed => Self::MediumDashed,
            OfficeCellBorderStyle::MediumDashDot => Self::MediumDashDot,
            OfficeCellBorderStyle::MediumDashDotDot => Self::MediumDashDotDot,
            OfficeCellBorderStyle::SlantDashDot => Self::SlantDashDot,
        }
    }
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, schemars::JsonSchema)]
#[serde(tag = "kind", rename_all = "kebab-case", deny_unknown_fields)]
enum OfficeCellBorderLine {
    /// Explicitly remove this border line.
    None,
    /// Apply one native Excel line style and optional 24-bit RGB color.
    Line {
        style: OfficeCellBorderStyle,
        color: Option<OfficeRgbColor>,
    },
}

impl From<OfficeCellBorderLine> for NativeSpreadsheetBorderLine {
    fn from(value: OfficeCellBorderLine) -> Self {
        match value {
            OfficeCellBorderLine::None => Self::None,
            OfficeCellBorderLine::Line { style, color } => Self::Line {
                style: style.into(),
                color: color
                    .map(|color| NativeOfficeRgbColor::new(color.red, color.green, color.blue)),
            },
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, schemars::JsonSchema)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
struct OfficeCellBorder {
    /// Left cell-border line update.
    left: Option<OfficeCellBorderLine>,
    /// Right cell-border line update.
    right: Option<OfficeCellBorderLine>,
    /// Top cell-border line update.
    top: Option<OfficeCellBorderLine>,
    /// Bottom cell-border line update.
    bottom: Option<OfficeCellBorderLine>,
    /// Shared diagonal line update.
    diagonal: Option<OfficeCellBorderLine>,
    /// Draw the bottom-left to top-right diagonal.
    diagonal_up: Option<bool>,
    /// Draw the top-left to bottom-right diagonal.
    diagonal_down: Option<bool>,
}

impl From<OfficeCellBorder> for NativeSpreadsheetBorder {
    fn from(value: OfficeCellBorder) -> Self {
        Self {
            left: value.left.map(Into::into),
            right: value.right.map(Into::into),
            top: value.top.map(Into::into),
            bottom: value.bottom.map(Into::into),
            diagonal: value.diagonal.map(Into::into),
            diagonal_up: value.diagonal_up,
            diagonal_down: value.diagonal_down,
        }
    }
}

impl From<OfficeCellFill> for NativeSpreadsheetFill {
    fn from(value: OfficeCellFill) -> Self {
        match value {
            OfficeCellFill::None => Self::None,
            OfficeCellFill::Solid { color } => Self::Solid {
                color: NativeOfficeRgbColor::new(color.red, color.green, color.blue),
            },
        }
    }
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, schemars::JsonSchema)]
#[serde(rename_all = "kebab-case")]
enum OfficeCellVerticalAlignment {
    Top,
    Center,
    Bottom,
    Justify,
    Distributed,
}

impl From<OfficeCellVerticalAlignment> for NativeSpreadsheetVerticalAlignment {
    fn from(value: OfficeCellVerticalAlignment) -> Self {
        match value {
            OfficeCellVerticalAlignment::Top => Self::Top,
            OfficeCellVerticalAlignment::Center => Self::Center,
            OfficeCellVerticalAlignment::Bottom => Self::Bottom,
            OfficeCellVerticalAlignment::Justify => Self::Justify,
            OfficeCellVerticalAlignment::Distributed => Self::Distributed,
        }
    }
}

#[derive(Debug, Clone, Copy, Serialize, Deserialize, schemars::JsonSchema)]
#[serde(rename_all = "kebab-case")]
enum OfficeCellReadingOrder {
    Context,
    LeftToRight,
    RightToLeft,
}

impl From<OfficeCellReadingOrder> for NativeSpreadsheetReadingOrder {
    fn from(value: OfficeCellReadingOrder) -> Self {
        match value {
            OfficeCellReadingOrder::Context => Self::Context,
            OfficeCellReadingOrder::LeftToRight => Self::LeftToRight,
            OfficeCellReadingOrder::RightToLeft => Self::RightToLeft,
        }
    }
}

#[derive(Debug, Clone, Serialize, Deserialize, schemars::JsonSchema)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub(in crate::office_mcp) struct OfficeCellFormat {
    /// Built-in alias or explicit Excel number format code, up to 255 characters.
    number_format: Option<String>,
    /// Solid RGB fill or an explicit fill clear.
    fill: Option<OfficeCellFill>,
    /// Partial typed update for cardinal and diagonal cell borders.
    border: Option<OfficeCellBorder>,
    /// Vertical cell alignment.
    vertical_alignment: Option<OfficeCellVerticalAlignment>,
    /// Explicitly enable or disable wrapped text.
    wrap_text: Option<bool>,
    /// Text rotation from 0 through 180, or 255 for stacked vertical text.
    text_rotation: Option<u16>,
    /// Alignment indentation from 0 through 255.
    indent: Option<u8>,
    /// Explicitly enable or disable shrink-to-fit.
    shrink_to_fit: Option<bool>,
    /// Contextual, left-to-right, or right-to-left cell reading order.
    reading_order: Option<OfficeCellReadingOrder>,
}

impl From<OfficeCellFormat> for NativeSpreadsheetCellFormat {
    fn from(value: OfficeCellFormat) -> Self {
        Self {
            number_format: value.number_format,
            fill: value.fill.map(Into::into),
            border: value.border.map(Into::into),
            vertical_alignment: value.vertical_alignment.map(Into::into),
            wrap_text: value.wrap_text,
            text_rotation: value.text_rotation,
            indent: value.indent,
            shrink_to_fit: value.shrink_to_fit,
            reading_order: value.reading_order.map(Into::into),
        }
    }
}
