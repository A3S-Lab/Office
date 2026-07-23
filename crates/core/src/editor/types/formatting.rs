use a3s_use_core::UseResult;
use serde::{Deserialize, Serialize};

use crate::editor::editor_error;

/// Horizontal text alignment supported consistently by native Office formats.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum NativeOfficeHorizontalAlignment {
    Left,
    Center,
    Right,
    Justify,
}

/// Explicit underline style supported consistently by native Office formats.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum NativeOfficeUnderline {
    None,
    Single,
    Double,
}

/// Vertical text script supported consistently by native Office formats.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum NativeOfficeTextScript {
    Baseline,
    Superscript,
    Subscript,
}

/// Display-only capitalization shared by native Word and Presentation runs.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum NativeOfficeTextCase {
    None,
    SmallCaps,
    AllCaps,
}

/// Portable highlight colors shared by native Word and Presentation runs.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum NativeOfficeHighlightColor {
    None,
    Black,
    Blue,
    Cyan,
    DarkBlue,
    DarkCyan,
    DarkGray,
    DarkGreen,
    DarkMagenta,
    DarkRed,
    DarkYellow,
    Green,
    LightGray,
    Magenta,
    Red,
    White,
    Yellow,
}

impl NativeOfficeHighlightColor {
    pub(crate) const fn word_value(self) -> &'static str {
        match self {
            Self::None => "none",
            Self::Black => "black",
            Self::Blue => "blue",
            Self::Cyan => "cyan",
            Self::DarkBlue => "darkBlue",
            Self::DarkCyan => "darkCyan",
            Self::DarkGray => "darkGray",
            Self::DarkGreen => "darkGreen",
            Self::DarkMagenta => "darkMagenta",
            Self::DarkRed => "darkRed",
            Self::DarkYellow => "darkYellow",
            Self::Green => "green",
            Self::LightGray => "lightGray",
            Self::Magenta => "magenta",
            Self::Red => "red",
            Self::White => "white",
            Self::Yellow => "yellow",
        }
    }

    pub(crate) const fn rgb_hex(self) -> Option<&'static str> {
        match self {
            Self::None => None,
            Self::Black => Some("000000"),
            Self::Blue => Some("0000FF"),
            Self::Cyan => Some("00FFFF"),
            Self::DarkBlue => Some("000080"),
            Self::DarkCyan => Some("008080"),
            Self::DarkGray => Some("808080"),
            Self::DarkGreen => Some("008000"),
            Self::DarkMagenta => Some("800080"),
            Self::DarkRed => Some("800000"),
            Self::DarkYellow => Some("808000"),
            Self::Green => Some("00FF00"),
            Self::LightGray => Some("C0C0C0"),
            Self::Magenta => Some("FF00FF"),
            Self::Red => Some("FF0000"),
            Self::White => Some("FFFFFF"),
            Self::Yellow => Some("FFFF00"),
        }
    }
}

/// An exact 24-bit RGB text color.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficeRgbColor {
    pub red: u8,
    pub green: u8,
    pub blue: u8,
}

impl NativeOfficeRgbColor {
    pub const fn new(red: u8, green: u8, blue: u8) -> Self {
        Self { red, green, blue }
    }

    pub(crate) fn hex(self) -> String {
        format!("{:02X}{:02X}{:02X}", self.red, self.green, self.blue)
    }
}

/// Typed text formatting shared by Word runs, Spreadsheet cells, and
/// Presentation runs.
///
/// Font sizes use integer centipoints (1/100 point), which is DrawingML's
/// native unit and avoids floating-point serialization. Word supports only
/// half-point increments and rejects values it cannot represent exactly.
#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeOfficeTextFormat {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bold: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub italic: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub underline: Option<NativeOfficeUnderline>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub script: Option<NativeOfficeTextScript>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub strikethrough: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub double_strikethrough: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_case: Option<NativeOfficeTextCase>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub highlight: Option<NativeOfficeHighlightColor>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub language: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_family: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub font_size_centipoints: Option<u32>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_color: Option<NativeOfficeRgbColor>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub alignment: Option<NativeOfficeHorizontalAlignment>,
}

impl NativeOfficeTextFormat {
    pub fn is_empty(&self) -> bool {
        self.bold.is_none()
            && self.italic.is_none()
            && self.underline.is_none()
            && self.script.is_none()
            && self.strikethrough.is_none()
            && self.double_strikethrough.is_none()
            && self.text_case.is_none()
            && self.highlight.is_none()
            && self.language.is_none()
            && self.font_family.is_none()
            && self.font_size_centipoints.is_none()
            && self.text_color.is_none()
            && self.alignment.is_none()
    }

    pub(crate) fn validate(&self) -> UseResult<()> {
        if self.is_empty() {
            return Err(editor_error(
                "use.office.text_format_empty",
                "Native Office text formatting requires at least one typed property.",
            ));
        }
        if let Some(family) = &self.font_family {
            if family.is_empty()
                || family.len() > 255
                || family.trim() != family
                || family.chars().any(char::is_control)
            {
                return Err(editor_error(
                    "use.office.font_family_invalid",
                    "Native Office font families must contain 1-255 non-control UTF-8 bytes without surrounding whitespace.",
                ));
            }
        }
        if let Some(language) = &self.language {
            if !is_language_tag(language) {
                return Err(editor_error(
                    "use.office.language_invalid",
                    "Native Office language tags must use a conservative 2-35 character BCP-47 shape.",
                ));
            }
        }
        if let Some(size) = self.font_size_centipoints {
            if !(100..=40_000).contains(&size) {
                return Err(editor_error(
                    "use.office.font_size_invalid",
                    "Native Office font sizes must be from 100 through 40000 centipoints (1-400 points).",
                ));
            }
        }
        Ok(())
    }

    pub(crate) fn has_character_properties(&self) -> bool {
        self.bold.is_some()
            || self.italic.is_some()
            || self.underline.is_some()
            || self.script.is_some()
            || self.strikethrough.is_some()
            || self.double_strikethrough.is_some()
            || self.text_case.is_some()
            || self.highlight.is_some()
            || self.language.is_some()
            || self.font_family.is_some()
            || self.font_size_centipoints.is_some()
            || self.text_color.is_some()
    }
}

fn is_language_tag(value: &str) -> bool {
    if !(2..=35).contains(&value.len()) || !value.is_ascii() {
        return false;
    }
    let mut subtags = value.split('-');
    let Some(primary) = subtags.next() else {
        return false;
    };
    if !((2..=8).contains(&primary.len()) && primary.bytes().all(|byte| byte.is_ascii_alphabetic())
        || primary.eq_ignore_ascii_case("x"))
    {
        return false;
    }
    subtags.all(|subtag| {
        (1..=8).contains(&subtag.len()) && subtag.bytes().all(|byte| byte.is_ascii_alphanumeric())
    })
}

/// Solid or explicitly cleared Spreadsheet cell fill.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "kebab-case", deny_unknown_fields)]
pub enum NativeSpreadsheetFill {
    None,
    Solid { color: NativeOfficeRgbColor },
}

/// Native Spreadsheet border line styles supported by SpreadsheetML.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase")]
pub enum NativeSpreadsheetBorderStyle {
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

impl NativeSpreadsheetBorderStyle {
    pub(crate) const fn spreadsheet_value(self) -> &'static str {
        match self {
            Self::Thin => "thin",
            Self::Medium => "medium",
            Self::Thick => "thick",
            Self::Double => "double",
            Self::Dashed => "dashed",
            Self::Dotted => "dotted",
            Self::DashDot => "dashDot",
            Self::DashDotDot => "dashDotDot",
            Self::Hair => "hair",
            Self::MediumDashed => "mediumDashed",
            Self::MediumDashDot => "mediumDashDot",
            Self::MediumDashDotDot => "mediumDashDotDot",
            Self::SlantDashDot => "slantDashDot",
        }
    }
}

/// One explicitly cleared or styled Spreadsheet cell-border line.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(tag = "kind", rename_all = "kebab-case", deny_unknown_fields)]
pub enum NativeSpreadsheetBorderLine {
    None,
    Line {
        style: NativeSpreadsheetBorderStyle,
        #[serde(default, skip_serializing_if = "Option::is_none")]
        color: Option<NativeOfficeRgbColor>,
    },
}

/// Partial update for the five Spreadsheet cell-border lines and diagonal
/// direction flags.
#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeSpreadsheetBorder {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub left: Option<NativeSpreadsheetBorderLine>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub right: Option<NativeSpreadsheetBorderLine>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub top: Option<NativeSpreadsheetBorderLine>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub bottom: Option<NativeSpreadsheetBorderLine>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub diagonal: Option<NativeSpreadsheetBorderLine>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub diagonal_up: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub diagonal_down: Option<bool>,
}

impl NativeSpreadsheetBorder {
    pub fn is_empty(&self) -> bool {
        self.left.is_none()
            && self.right.is_none()
            && self.top.is_none()
            && self.bottom.is_none()
            && self.diagonal.is_none()
            && self.diagonal_up.is_none()
            && self.diagonal_down.is_none()
    }
}

/// Vertical alignment supported by Spreadsheet cell styles.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum NativeSpreadsheetVerticalAlignment {
    Top,
    Center,
    Bottom,
    Justify,
    Distributed,
}

/// Bidirectional reading order supported by Spreadsheet cell styles.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "kebab-case")]
pub enum NativeSpreadsheetReadingOrder {
    Context,
    LeftToRight,
    RightToLeft,
}

/// Typed Spreadsheet cell presentation properties that are not text-run
/// formatting.
#[derive(Debug, Clone, Default, PartialEq, Eq, Serialize, Deserialize)]
#[serde(rename_all = "camelCase", deny_unknown_fields)]
pub struct NativeSpreadsheetCellFormat {
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub number_format: Option<String>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub fill: Option<NativeSpreadsheetFill>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub border: Option<NativeSpreadsheetBorder>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub vertical_alignment: Option<NativeSpreadsheetVerticalAlignment>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub wrap_text: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub text_rotation: Option<u16>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub indent: Option<u8>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub shrink_to_fit: Option<bool>,
    #[serde(default, skip_serializing_if = "Option::is_none")]
    pub reading_order: Option<NativeSpreadsheetReadingOrder>,
}

impl NativeSpreadsheetCellFormat {
    pub fn is_empty(&self) -> bool {
        self.number_format.is_none()
            && self.fill.is_none()
            && self.border.is_none()
            && self.vertical_alignment.is_none()
            && self.wrap_text.is_none()
            && self.text_rotation.is_none()
            && self.indent.is_none()
            && self.shrink_to_fit.is_none()
            && self.reading_order.is_none()
    }

    pub(crate) fn validate(&self) -> UseResult<()> {
        if self.is_empty() {
            return Err(editor_error(
                "use.office.cell_format_empty",
                "Native Spreadsheet cell formatting requires at least one typed property.",
            ));
        }
        if let Some(number_format) = &self.number_format {
            validate_number_format(&normalize_number_format(number_format)?)?;
        }
        if self
            .border
            .as_ref()
            .is_some_and(NativeSpreadsheetBorder::is_empty)
        {
            return Err(editor_error(
                "use.office.cell_border_empty",
                "Native Spreadsheet cell borders require at least one typed line or diagonal direction property.",
            ));
        }
        if self
            .text_rotation
            .is_some_and(|rotation| rotation > 180 && rotation != 255)
        {
            return Err(editor_error(
                "use.office.text_rotation_invalid",
                "Spreadsheet text rotation must be from 0 through 180 degrees or 255 for vertical stacked text.",
            ));
        }
        Ok(())
    }

    pub(crate) fn normalized_number_format(&self) -> UseResult<Option<String>> {
        self.number_format
            .as_deref()
            .map(normalize_number_format)
            .transpose()
    }

    pub(crate) fn has_alignment_properties(&self) -> bool {
        self.vertical_alignment.is_some()
            || self.wrap_text.is_some()
            || self.text_rotation.is_some()
            || self.indent.is_some()
            || self.shrink_to_fit.is_some()
            || self.reading_order.is_some()
    }
}

fn normalize_number_format(value: &str) -> UseResult<String> {
    let alias = match value.trim().to_ascii_lowercase().as_str() {
        "general" => Some("General"),
        "number" | "comma" => Some("#,##0.00"),
        "currency" => Some("\"$\"#,##0.00"),
        "accounting" => {
            Some("_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)")
        }
        "percent" | "percentage" => Some("0.00%"),
        "scientific" => Some("0.00E+00"),
        "text" => Some("@"),
        "date" => Some("yyyy-mm-dd"),
        "time" => Some("h:mm:ss"),
        "datetime" => Some("yyyy-mm-dd h:mm:ss"),
        _ => None,
    };
    Ok(alias.map_or_else(|| value.to_string(), str::to_string))
}

fn validate_number_format(value: &str) -> UseResult<()> {
    if value.trim().is_empty() || value.chars().count() > 255 {
        return Err(editor_error(
            "use.office.number_format_invalid",
            "Spreadsheet number formats must contain 1-255 characters.",
        ));
    }
    if value.chars().any(|character| {
        !matches!(character, '\u{9}' | '\u{a}' | '\u{d}')
            && (character < '\u{20}' || matches!(character, '\u{fffe}' | '\u{ffff}'))
    }) {
        return Err(editor_error(
            "use.office.number_format_invalid",
            "Spreadsheet number formats cannot contain XML-forbidden characters.",
        ));
    }

    let mut quoted = false;
    let mut bracket_depth = 0usize;
    let mut sections = 1usize;
    let mut characters = value.chars();
    while let Some(character) = characters.next() {
        match character {
            '"' => quoted = !quoted,
            '\\' | '_' | '*' if !quoted => {
                characters.next();
            }
            '[' if !quoted => bracket_depth += 1,
            ']' if !quoted => {
                bracket_depth = bracket_depth.checked_sub(1).ok_or_else(|| {
                    editor_error(
                        "use.office.number_format_invalid",
                        "Spreadsheet number formats must use balanced square brackets.",
                    )
                })?;
            }
            ';' if !quoted && bracket_depth == 0 => sections += 1,
            _ => {}
        }
    }
    if quoted || bracket_depth != 0 {
        return Err(editor_error(
            "use.office.number_format_invalid",
            "Spreadsheet number formats must use balanced quotes and square brackets.",
        ));
    }
    if sections > 4 {
        return Err(editor_error(
            "use.office.number_format_invalid",
            "Spreadsheet number formats can contain at most four semicolon-separated sections.",
        ));
    }
    Ok(())
}
