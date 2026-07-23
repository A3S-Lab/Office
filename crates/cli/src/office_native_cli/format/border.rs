use a3s_office::{
    NativeOfficeRgbColor, NativeSpreadsheetBorder, NativeSpreadsheetBorderLine,
    NativeSpreadsheetBorderStyle,
};
use a3s_use_core::UseResult;

use super::{parse_format_boolean, parse_rgb_color, usage_error, ParsedArguments};

#[derive(Clone, Copy)]
enum ParsedBorderStyle {
    None,
    Line(NativeSpreadsheetBorderStyle),
}

pub(super) fn parse(parsed: &ParsedArguments) -> UseResult<Option<NativeSpreadsheetBorder>> {
    let all_style = parsed
        .border
        .all
        .as_deref()
        .map(|value| parse_style("--border-all", value))
        .transpose()?;
    let all_color = parsed
        .border
        .color
        .as_deref()
        .map(|value| parse_rgb_color("--border-color", value))
        .transpose()?;

    let left = parse_line(
        "--border-left",
        parsed
            .border
            .left
            .as_deref()
            .map(|value| parse_style("--border-left", value))
            .transpose()?
            .or(all_style),
        "--border-left-color",
        parsed
            .border
            .left_color
            .as_deref()
            .map(|value| parse_rgb_color("--border-left-color", value))
            .transpose()?,
        all_color,
    )?;
    let right = parse_line(
        "--border-right",
        parsed
            .border
            .right
            .as_deref()
            .map(|value| parse_style("--border-right", value))
            .transpose()?
            .or(all_style),
        "--border-right-color",
        parsed
            .border
            .right_color
            .as_deref()
            .map(|value| parse_rgb_color("--border-right-color", value))
            .transpose()?,
        all_color,
    )?;
    let top = parse_line(
        "--border-top",
        parsed
            .border
            .top
            .as_deref()
            .map(|value| parse_style("--border-top", value))
            .transpose()?
            .or(all_style),
        "--border-top-color",
        parsed
            .border
            .top_color
            .as_deref()
            .map(|value| parse_rgb_color("--border-top-color", value))
            .transpose()?,
        all_color,
    )?;
    let bottom = parse_line(
        "--border-bottom",
        parsed
            .border
            .bottom
            .as_deref()
            .map(|value| parse_style("--border-bottom", value))
            .transpose()?
            .or(all_style),
        "--border-bottom-color",
        parsed
            .border
            .bottom_color
            .as_deref()
            .map(|value| parse_rgb_color("--border-bottom-color", value))
            .transpose()?,
        all_color,
    )?;
    if all_color.is_some()
        && ![left, right, top, bottom]
            .into_iter()
            .any(|line| matches!(line, Some(NativeSpreadsheetBorderLine::Line { .. })))
    {
        return Err(usage_error(
            "--border-color requires --border-all or at least one non-none side border",
        ));
    }

    let diagonal = parse_line(
        "--border-diagonal",
        parsed
            .border
            .diagonal
            .as_deref()
            .map(|value| parse_style("--border-diagonal", value))
            .transpose()?,
        "--border-diagonal-color",
        parsed
            .border
            .diagonal_color
            .as_deref()
            .map(|value| parse_rgb_color("--border-diagonal-color", value))
            .transpose()?,
        None,
    )?;
    let border = NativeSpreadsheetBorder {
        left,
        right,
        top,
        bottom,
        diagonal,
        diagonal_up: parsed
            .border
            .diagonal_up
            .as_deref()
            .map(|value| parse_format_boolean("--border-diagonal-up", value))
            .transpose()?,
        diagonal_down: parsed
            .border
            .diagonal_down
            .as_deref()
            .map(|value| parse_format_boolean("--border-diagonal-down", value))
            .transpose()?,
    };
    Ok((!border.is_empty()).then_some(border))
}

fn parse_line(
    style_option: &str,
    style: Option<ParsedBorderStyle>,
    color_option: &str,
    color: Option<NativeOfficeRgbColor>,
    fallback_color: Option<NativeOfficeRgbColor>,
) -> UseResult<Option<NativeSpreadsheetBorderLine>> {
    match style {
        None if color.is_some() => Err(usage_error(format!(
            "{color_option} requires {style_option} or --border-all"
        ))),
        None => Ok(None),
        Some(ParsedBorderStyle::None) if color.is_some() => Err(usage_error(format!(
            "{color_option} cannot be combined with a none border"
        ))),
        Some(ParsedBorderStyle::None) => Ok(Some(NativeSpreadsheetBorderLine::None)),
        Some(ParsedBorderStyle::Line(style)) => Ok(Some(NativeSpreadsheetBorderLine::Line {
            style,
            color: color.or(fallback_color),
        })),
    }
}

fn parse_style(option: &str, value: &str) -> UseResult<ParsedBorderStyle> {
    let normalized = value.to_ascii_lowercase().replace(['-', '_'], "");
    let style = match normalized.as_str() {
        "none" => return Ok(ParsedBorderStyle::None),
        "thin" => NativeSpreadsheetBorderStyle::Thin,
        "medium" => NativeSpreadsheetBorderStyle::Medium,
        "thick" => NativeSpreadsheetBorderStyle::Thick,
        "double" => NativeSpreadsheetBorderStyle::Double,
        "dashed" => NativeSpreadsheetBorderStyle::Dashed,
        "dotted" => NativeSpreadsheetBorderStyle::Dotted,
        "dashdot" => NativeSpreadsheetBorderStyle::DashDot,
        "dashdotdot" => NativeSpreadsheetBorderStyle::DashDotDot,
        "hair" => NativeSpreadsheetBorderStyle::Hair,
        "mediumdashed" => NativeSpreadsheetBorderStyle::MediumDashed,
        "mediumdashdot" => NativeSpreadsheetBorderStyle::MediumDashDot,
        "mediumdashdotdot" => NativeSpreadsheetBorderStyle::MediumDashDotDot,
        "slantdashdot" => NativeSpreadsheetBorderStyle::SlantDashDot,
        _ => {
            return Err(usage_error(format!(
                "{option} requires none or a native Excel border style, received '{value}'"
            )))
        }
    };
    Ok(ParsedBorderStyle::Line(style))
}
