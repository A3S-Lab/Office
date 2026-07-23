use a3s_office::{
    NativeOfficeHighlightColor, NativeOfficeHorizontalAlignment, NativeOfficeHyperlink,
    NativeOfficeRgbColor, NativeOfficeTextCase, NativeOfficeTextFormat, NativeOfficeTextScript,
    NativeOfficeUnderline, NativeSpreadsheetCellFormat, NativeSpreadsheetFill,
    NativeSpreadsheetReadingOrder, NativeSpreadsheetVerticalAlignment,
};
use a3s_use_core::UseResult;

use super::{usage_error, ParsedArguments};

mod border;

pub(super) fn parse_text_format(
    parsed: &ParsedArguments,
) -> UseResult<Option<NativeOfficeTextFormat>> {
    let format = NativeOfficeTextFormat {
        bold: parsed
            .bold
            .as_deref()
            .map(|value| parse_format_boolean("--bold", value))
            .transpose()?,
        italic: parsed
            .italic
            .as_deref()
            .map(|value| parse_format_boolean("--italic", value))
            .transpose()?,
        underline: parsed
            .underline
            .as_deref()
            .map(parse_underline)
            .transpose()?,
        script: parsed
            .script
            .as_deref()
            .map(parse_text_script)
            .transpose()?,
        strikethrough: parsed
            .strikethrough
            .as_deref()
            .map(|value| parse_format_boolean("--strikethrough", value))
            .transpose()?,
        double_strikethrough: parsed
            .double_strikethrough
            .as_deref()
            .map(|value| parse_format_boolean("--double-strikethrough", value))
            .transpose()?,
        text_case: parsed
            .text_case
            .as_deref()
            .map(parse_text_case)
            .transpose()?,
        highlight: parsed
            .highlight
            .as_deref()
            .map(parse_highlight)
            .transpose()?,
        language: parsed.language.clone(),
        font_family: parsed.font_family.clone(),
        font_size_centipoints: parsed
            .font_size
            .as_deref()
            .map(parse_font_size)
            .transpose()?,
        text_color: parsed
            .text_color
            .as_deref()
            .map(parse_text_color)
            .transpose()?,
        alignment: parsed
            .alignment
            .as_deref()
            .map(parse_alignment)
            .transpose()?,
    };
    Ok((!format.is_empty()).then_some(format))
}

pub(super) fn parse_cell_format(
    parsed: &ParsedArguments,
) -> UseResult<Option<NativeSpreadsheetCellFormat>> {
    let format = NativeSpreadsheetCellFormat {
        number_format: parsed.number_format.clone(),
        fill: parsed.fill.as_deref().map(parse_cell_fill).transpose()?,
        border: border::parse(parsed)?,
        vertical_alignment: parsed
            .vertical_alignment
            .as_deref()
            .map(parse_vertical_alignment)
            .transpose()?,
        wrap_text: parsed
            .wrap_text
            .as_deref()
            .map(|value| parse_format_boolean("--wrap-text", value))
            .transpose()?,
        text_rotation: parsed
            .text_rotation
            .as_deref()
            .map(|value| parse_integer_option::<u16>("--text-rotation", value))
            .transpose()?,
        indent: parsed
            .indent
            .as_deref()
            .map(|value| parse_integer_option::<u8>("--indent", value))
            .transpose()?,
        shrink_to_fit: parsed
            .shrink_to_fit
            .as_deref()
            .map(|value| parse_format_boolean("--shrink-to-fit", value))
            .transpose()?,
        reading_order: parsed
            .reading_order
            .as_deref()
            .map(parse_reading_order)
            .transpose()?,
    };
    Ok((!format.is_empty()).then_some(format))
}

fn parse_cell_fill(value: &str) -> UseResult<NativeSpreadsheetFill> {
    if value.eq_ignore_ascii_case("none") {
        Ok(NativeSpreadsheetFill::None)
    } else {
        Ok(NativeSpreadsheetFill::Solid {
            color: parse_rgb_color("--fill", value)?,
        })
    }
}

fn parse_vertical_alignment(value: &str) -> UseResult<NativeSpreadsheetVerticalAlignment> {
    match value.to_ascii_lowercase().as_str() {
        "top" => Ok(NativeSpreadsheetVerticalAlignment::Top),
        "center" | "centre" => Ok(NativeSpreadsheetVerticalAlignment::Center),
        "bottom" => Ok(NativeSpreadsheetVerticalAlignment::Bottom),
        "justify" | "justified" => Ok(NativeSpreadsheetVerticalAlignment::Justify),
        "distributed" => Ok(NativeSpreadsheetVerticalAlignment::Distributed),
        _ => Err(usage_error(format!(
            "--vertical-align requires top, center, bottom, justify, or distributed, received '{value}'"
        ))),
    }
}

fn parse_reading_order(value: &str) -> UseResult<NativeSpreadsheetReadingOrder> {
    match value.to_ascii_lowercase().replace(['-', '_'], "").as_str() {
        "0" | "context" | "contextdependent" => Ok(NativeSpreadsheetReadingOrder::Context),
        "1" | "ltr" | "lefttoright" => Ok(NativeSpreadsheetReadingOrder::LeftToRight),
        "2" | "rtl" | "righttoleft" => Ok(NativeSpreadsheetReadingOrder::RightToLeft),
        _ => Err(usage_error(format!(
            "--reading-order requires context, ltr, or rtl, received '{value}'"
        ))),
    }
}

fn parse_integer_option<T>(option: &str, value: &str) -> UseResult<T>
where
    T: std::str::FromStr,
{
    value.parse::<T>().map_err(|_| {
        usage_error(format!(
            "{option} requires a non-negative integer in its documented range, received '{value}'"
        ))
    })
}

pub(super) fn parse_hyperlink(
    parsed: &ParsedArguments,
    display: Option<&str>,
) -> UseResult<Option<NativeOfficeHyperlink>> {
    if parsed.url.is_some() && parsed.location.is_some() {
        return Err(usage_error(
            "native Office hyperlink accepts exactly one of --url or --location",
        ));
    }
    let mut hyperlink = if let Some(uri) = &parsed.url {
        Some(NativeOfficeHyperlink::external(uri)?)
    } else if let Some(location) = &parsed.location {
        Some(NativeOfficeHyperlink::internal(location)?)
    } else {
        None
    };
    if hyperlink.is_none() && (display.is_some() || parsed.tooltip.is_some()) {
        return Err(usage_error(
            "--display and --tooltip require --url or --location",
        ));
    }
    if let Some(value) = hyperlink.as_mut() {
        if let Some(display) = display {
            *value = value.clone().with_display(display);
        }
        if let Some(tooltip) = &parsed.tooltip {
            *value = value.clone().with_tooltip(tooltip);
        }
    }
    Ok(hyperlink)
}

pub(super) fn parse_format_boolean(option: &str, value: &str) -> UseResult<bool> {
    match value.to_ascii_lowercase().as_str() {
        "true" => Ok(true),
        "false" => Ok(false),
        _ => Err(usage_error(format!(
            "{option} requires true or false, received '{value}'"
        ))),
    }
}

fn parse_font_size(value: &str) -> UseResult<u32> {
    let normalized = value
        .strip_suffix("pt")
        .or_else(|| value.strip_suffix("pT"))
        .or_else(|| value.strip_suffix("Pt"))
        .or_else(|| value.strip_suffix("PT"))
        .unwrap_or(value);
    let (whole, fraction) = normalized
        .split_once('.')
        .map_or((normalized, ""), |(whole, fraction)| (whole, fraction));
    if whole.is_empty()
        || !whole.bytes().all(|byte| byte.is_ascii_digit())
        || fraction.len() > 2
        || !fraction.bytes().all(|byte| byte.is_ascii_digit())
    {
        return Err(usage_error(format!(
            "--font-size requires points with at most two decimals, received '{value}'"
        )));
    }
    let whole = whole.parse::<u32>().map_err(|_| {
        usage_error(format!(
            "--font-size requires points with at most two decimals, received '{value}'"
        ))
    })?;
    let fraction = match fraction.len() {
        0 => 0,
        1 => fraction.parse::<u32>().unwrap_or_default() * 10,
        _ => fraction.parse::<u32>().unwrap_or_default(),
    };
    let centipoints = whole
        .checked_mul(100)
        .and_then(|value| value.checked_add(fraction))
        .ok_or_else(|| usage_error("--font-size is too large"))?;
    if !(100..=40_000).contains(&centipoints) {
        return Err(usage_error("--font-size must be from 1 through 400 points"));
    }
    Ok(centipoints)
}

fn parse_text_color(value: &str) -> UseResult<NativeOfficeRgbColor> {
    parse_rgb_color("--text-color", value)
}

pub(super) fn parse_rgb_color(option: &str, value: &str) -> UseResult<NativeOfficeRgbColor> {
    let value = value.strip_prefix('#').unwrap_or(value);
    if value.len() != 6 || !value.bytes().all(|byte| byte.is_ascii_hexdigit()) {
        return Err(usage_error(format!(
            "{option} requires exactly six hexadecimal RGB digits, received '{value}'"
        )));
    }
    let component = |range: std::ops::Range<usize>| {
        u8::from_str_radix(&value[range], 16).map_err(|_| {
            usage_error(format!(
                "{option} requires exactly six hexadecimal RGB digits, received '{value}'"
            ))
        })
    };
    Ok(NativeOfficeRgbColor::new(
        component(0..2)?,
        component(2..4)?,
        component(4..6)?,
    ))
}

fn parse_alignment(value: &str) -> UseResult<NativeOfficeHorizontalAlignment> {
    match value.to_ascii_lowercase().as_str() {
        "left" => Ok(NativeOfficeHorizontalAlignment::Left),
        "center" | "centre" => Ok(NativeOfficeHorizontalAlignment::Center),
        "right" => Ok(NativeOfficeHorizontalAlignment::Right),
        "justify" | "justified" => Ok(NativeOfficeHorizontalAlignment::Justify),
        _ => Err(usage_error(format!(
            "--align requires left, center, right, or justify, received '{value}'"
        ))),
    }
}

fn parse_underline(value: &str) -> UseResult<NativeOfficeUnderline> {
    match value.to_ascii_lowercase().as_str() {
        "none" => Ok(NativeOfficeUnderline::None),
        "single" => Ok(NativeOfficeUnderline::Single),
        "double" => Ok(NativeOfficeUnderline::Double),
        _ => Err(usage_error(format!(
            "--underline requires none, single, or double, received '{value}'"
        ))),
    }
}

fn parse_text_script(value: &str) -> UseResult<NativeOfficeTextScript> {
    match value.to_ascii_lowercase().as_str() {
        "baseline" => Ok(NativeOfficeTextScript::Baseline),
        "superscript" => Ok(NativeOfficeTextScript::Superscript),
        "subscript" => Ok(NativeOfficeTextScript::Subscript),
        _ => Err(usage_error(format!(
            "--script requires baseline, superscript, or subscript, received '{value}'"
        ))),
    }
}

fn parse_text_case(value: &str) -> UseResult<NativeOfficeTextCase> {
    match value.to_ascii_lowercase().replace(['-', '_'], "").as_str() {
        "none" => Ok(NativeOfficeTextCase::None),
        "small" | "smallcaps" => Ok(NativeOfficeTextCase::SmallCaps),
        "all" | "allcaps" => Ok(NativeOfficeTextCase::AllCaps),
        _ => Err(usage_error(format!(
            "--text-case requires none, small-caps, or all-caps, received '{value}'"
        ))),
    }
}

fn parse_highlight(value: &str) -> UseResult<NativeOfficeHighlightColor> {
    let normalized = value.to_ascii_lowercase().replace(['-', '_'], "");
    match normalized.as_str() {
        "none" => Ok(NativeOfficeHighlightColor::None),
        "black" => Ok(NativeOfficeHighlightColor::Black),
        "blue" => Ok(NativeOfficeHighlightColor::Blue),
        "cyan" => Ok(NativeOfficeHighlightColor::Cyan),
        "darkblue" => Ok(NativeOfficeHighlightColor::DarkBlue),
        "darkcyan" => Ok(NativeOfficeHighlightColor::DarkCyan),
        "darkgray" | "darkgrey" => Ok(NativeOfficeHighlightColor::DarkGray),
        "darkgreen" => Ok(NativeOfficeHighlightColor::DarkGreen),
        "darkmagenta" => Ok(NativeOfficeHighlightColor::DarkMagenta),
        "darkred" => Ok(NativeOfficeHighlightColor::DarkRed),
        "darkyellow" => Ok(NativeOfficeHighlightColor::DarkYellow),
        "green" => Ok(NativeOfficeHighlightColor::Green),
        "lightgray" | "lightgrey" => Ok(NativeOfficeHighlightColor::LightGray),
        "magenta" => Ok(NativeOfficeHighlightColor::Magenta),
        "red" => Ok(NativeOfficeHighlightColor::Red),
        "white" => Ok(NativeOfficeHighlightColor::White),
        "yellow" => Ok(NativeOfficeHighlightColor::Yellow),
        _ => Err(usage_error(format!(
            "--highlight requires a portable Word highlight color or none, received '{value}'"
        ))),
    }
}
