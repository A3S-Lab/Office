use pdfium_render::prelude::{PdfColor, PdfFontWeight, PdfPageTextRenderMode};
use serde::{Deserialize, Serialize};

/// Source-native RGBA text color returned by PDFium.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficePdfTextColor {
    pub red: u8,
    pub green: u8,
    pub blue: u8,
    pub alpha: u8,
}

impl From<PdfColor> for NativeOfficePdfTextColor {
    fn from(value: PdfColor) -> Self {
        Self {
            red: value.red(),
            green: value.green(),
            blue: value.blue(),
            alpha: value.alpha(),
        }
    }
}

/// PDF text painting mode needed to distinguish fill, stroke, and invisible
/// source text without guessing which color is visually authoritative.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum NativeOfficePdfTextRenderMode {
    Unknown,
    FilledUnstroked,
    StrokedUnfilled,
    FilledThenStroked,
    Invisible,
    FilledUnstrokedClipping,
    StrokedUnfilledClipping,
    FilledThenStrokedClipping,
    InvisibleClipping,
}

impl From<PdfPageTextRenderMode> for NativeOfficePdfTextRenderMode {
    fn from(value: PdfPageTextRenderMode) -> Self {
        match value {
            PdfPageTextRenderMode::Unknown => Self::Unknown,
            PdfPageTextRenderMode::FilledUnstroked => Self::FilledUnstroked,
            PdfPageTextRenderMode::StrokedUnfilled => Self::StrokedUnfilled,
            PdfPageTextRenderMode::FilledThenStroked => Self::FilledThenStroked,
            PdfPageTextRenderMode::Invisible => Self::Invisible,
            PdfPageTextRenderMode::FilledUnstrokedClipping => Self::FilledUnstrokedClipping,
            PdfPageTextRenderMode::StrokedUnfilledClipping => Self::StrokedUnfilledClipping,
            PdfPageTextRenderMode::FilledThenStrokedClipping => Self::FilledThenStrokedClipping,
            PdfPageTextRenderMode::InvisibleClipping => Self::InvisibleClipping,
        }
    }
}

pub(super) fn font_weight_value(value: PdfFontWeight) -> Option<u16> {
    let value = match value {
        PdfFontWeight::Weight100 => 100,
        PdfFontWeight::Weight200 => 200,
        PdfFontWeight::Weight300 => 300,
        PdfFontWeight::Weight400Normal => 400,
        PdfFontWeight::Weight500 => 500,
        PdfFontWeight::Weight600 => 600,
        PdfFontWeight::Weight700Bold => 700,
        PdfFontWeight::Weight800 => 800,
        PdfFontWeight::Weight900 => 900,
        PdfFontWeight::Custom(value) => u16::try_from(value).ok()?,
    };
    (1..=1_000).contains(&value).then_some(value)
}
