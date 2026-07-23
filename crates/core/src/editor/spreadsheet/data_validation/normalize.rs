use a3s_use_core::UseResult;

use super::super::editor_error;
use crate::editor::{NativeSpreadsheetDataValidation, NativeSpreadsheetDataValidationType};
use crate::spreadsheet_reference::CellRange;
use crate::xml_edit::index_xml;
use crate::NativeOfficePackage;

pub(super) fn reject_operator(validation: &NativeSpreadsheetDataValidation) -> UseResult<()> {
    if validation.operator.is_some() {
        return Err(editor_error(
            "use.office.spreadsheet_validation_operator_unsupported",
            "List and custom data validation do not accept a comparison operator.",
        ));
    }
    Ok(())
}

pub(super) fn workbook_uses_1904_date_system(package: &NativeOfficePackage) -> UseResult<bool> {
    let part_name = "xl/workbook.xml";
    let part = package.xml_part(part_name)?;
    let workbook = index_xml(&part)?;
    let properties = workbook
        .children
        .iter()
        .filter(|child| child.local_name == "workbookPr" && child.namespace == workbook.namespace)
        .collect::<Vec<_>>();
    if properties.len() > 1 {
        return Err(editor_error(
            "use.office.spreadsheet_validation_date_system_invalid",
            "Spreadsheet workbook.xml contains multiple workbookPr elements.",
        )
        .with_detail("part", part_name));
    }
    match properties
        .first()
        .and_then(|properties| properties.qualified_attributes.get("date1904"))
        .map(String::as_str)
    {
        None | Some("0" | "false") => Ok(false),
        Some("1" | "true") => Ok(true),
        Some(value) => Err(editor_error(
            "use.office.spreadsheet_validation_date_system_invalid",
            format!("Spreadsheet workbook.xml has invalid date1904='{value}'."),
        )
        .with_detail("part", part_name)
        .with_detail("date1904", value)),
    }
}

pub(super) fn normalize_list_formula(value: &str) -> UseResult<String> {
    let value = value.strip_prefix('=').unwrap_or(value);
    if value.is_empty() {
        return Err(editor_error(
            "use.office.spreadsheet_validation_formula_required",
            "List data validation requires a non-empty formula1.",
        ));
    }
    if value.starts_with('"') || value.ends_with('"') {
        if value.len() < 2 || !value.starts_with('"') || !value.ends_with('"') {
            return Err(invalid_list_formula());
        }
        let inner = &value[1..value.len() - 1];
        if inner.is_empty() || inner.contains('"') {
            return Err(invalid_list_formula());
        }
        return Ok(value.to_string());
    }
    if looks_like_list_formula(value) {
        return Ok(value.to_string());
    }
    if value.contains('"') {
        return Err(invalid_list_formula());
    }
    if value.contains(',') {
        return Ok(format!("\"{value}\""));
    }
    Ok(format!("\"{value}\""))
}

fn looks_like_list_formula(value: &str) -> bool {
    if value.contains([':', '!', '(']) || is_defined_name(value) {
        return true;
    }
    let reference = value.strip_suffix('#').unwrap_or(value);
    let reference = reference.replace('$', "");
    CellRange::parse(&reference).is_ok()
}

fn invalid_list_formula() -> a3s_use_core::UseError {
    editor_error(
        "use.office.spreadsheet_validation_list_invalid",
        "List validation values cannot contain embedded or unmatched double quotes; use a cell range for those values.",
    )
}

fn is_defined_name(value: &str) -> bool {
    let mut characters = value.chars();
    characters
        .next()
        .is_some_and(|character| character.is_ascii_alphabetic() || matches!(character, '_' | '\\'))
        && characters.all(|character| {
            character.is_ascii_alphanumeric() || matches!(character, '_' | '.' | '\\')
        })
}

pub(super) fn normalize_comparison_formula(
    validation_type: NativeSpreadsheetDataValidationType,
    value: &str,
    date_1904: bool,
) -> UseResult<String> {
    let value = strip_formula_equals(value)?;
    match validation_type {
        NativeSpreadsheetDataValidationType::Date if looks_like_iso_date(&value) => {
            excel_date_serial(&value, date_1904).map(|serial| serial.to_string())
        }
        NativeSpreadsheetDataValidationType::Time if looks_like_clock_time(&value) => {
            excel_time_fraction(&value)
        }
        _ => Ok(value),
    }
}

pub(super) fn strip_formula_equals(value: &str) -> UseResult<String> {
    let value = value.strip_prefix('=').unwrap_or(value);
    if value.is_empty() {
        return Err(editor_error(
            "use.office.spreadsheet_validation_formula_required",
            "Data-validation formulas cannot be empty.",
        ));
    }
    Ok(value.to_string())
}

fn looks_like_iso_date(value: &str) -> bool {
    value.len() == 10
        && value.as_bytes()[4] == b'-'
        && value.as_bytes()[7] == b'-'
        && value
            .bytes()
            .enumerate()
            .all(|(index, byte)| matches!(index, 4 | 7) || byte.is_ascii_digit())
}

fn excel_date_serial(value: &str, date_1904: bool) -> UseResult<i64> {
    let year = value[0..4].parse::<i32>().unwrap_or_default();
    let month = value[5..7].parse::<u32>().unwrap_or_default();
    let day = value[8..10].parse::<u32>().unwrap_or_default();
    if !(1900..=9999).contains(&year)
        || !(1..=12).contains(&month)
        || day == 0
        || day > days_in_month(year, month)
    {
        return Err(editor_error(
            "use.office.spreadsheet_validation_date_invalid",
            format!("Data-validation date '{value}' is not a valid YYYY-MM-DD date from 1900 through 9999."),
        ));
    }
    let baseline = days_from_civil(1899, 12, 31);
    let mut serial = days_from_civil(year, month, day) - baseline;
    if (year, month, day) >= (1900, 3, 1) {
        serial += 1;
    }
    if date_1904 {
        serial -= 1_462;
    }
    Ok(serial)
}

fn days_in_month(year: i32, month: u32) -> u32 {
    match month {
        1 | 3 | 5 | 7 | 8 | 10 | 12 => 31,
        4 | 6 | 9 | 11 => 30,
        2 if year % 400 == 0 || (year % 4 == 0 && year % 100 != 0) => 29,
        2 => 28,
        _ => 0,
    }
}

fn days_from_civil(year: i32, month: u32, day: u32) -> i64 {
    let year = year - i32::from(month <= 2);
    let era = if year >= 0 { year } else { year - 399 } / 400;
    let year_of_era = year - era * 400;
    let month = i32::try_from(month).unwrap_or_default();
    let day = i32::try_from(day).unwrap_or_default();
    let day_of_year = (153 * (month + if month > 2 { -3 } else { 9 }) + 2) / 5 + day - 1;
    let day_of_era = year_of_era * 365 + year_of_era / 4 - year_of_era / 100 + day_of_year;
    i64::from(era * 146_097 + day_of_era)
}

fn looks_like_clock_time(value: &str) -> bool {
    let parts = value.split(':').collect::<Vec<_>>();
    matches!(parts.len(), 2 | 3)
        && parts.iter().all(|part| {
            !part.is_empty() && part.len() <= 2 && part.bytes().all(|byte| byte.is_ascii_digit())
        })
}

fn excel_time_fraction(value: &str) -> UseResult<String> {
    let mut parts = value.split(':');
    let hour = parts.next().and_then(|value| value.parse::<u32>().ok());
    let minute = parts.next().and_then(|value| value.parse::<u32>().ok());
    let second = parts
        .next()
        .map_or(Some(0), |value| value.parse::<u32>().ok());
    let (Some(hour), Some(minute), Some(second)) = (hour, minute, second) else {
        return Err(invalid_time(value));
    };
    if hour > 23 || minute > 59 || second > 59 {
        return Err(invalid_time(value));
    }
    let seconds = hour * 3_600 + minute * 60 + second;
    let mut result = format!("{:.15}", f64::from(seconds) / 86_400.0);
    while result.ends_with('0') {
        result.pop();
    }
    if result.ends_with('.') {
        result.push('0');
    }
    Ok(result)
}

fn invalid_time(value: &str) -> a3s_use_core::UseError {
    editor_error(
        "use.office.spreadsheet_validation_time_invalid",
        format!("Data-validation time '{value}' must be HH:MM or HH:MM:SS from 00:00:00 through 23:59:59."),
    )
}

pub(super) fn validate_optional_text(
    value: Option<&str>,
    field: &str,
    max: usize,
) -> UseResult<()> {
    if let Some(value) = value {
        validate_xml_text(value, field, max)?;
    }
    Ok(())
}

pub(super) fn validate_xml_text(value: &str, field: &str, max: usize) -> UseResult<()> {
    let characters = value.chars().count();
    if value.is_empty() || characters > max {
        return Err(editor_error(
            "use.office.spreadsheet_validation_text_invalid",
            format!("Data-validation {field} must contain 1-{max} characters."),
        )
        .with_detail("field", field)
        .with_detail("characters", characters));
    }
    if let Some(character) = value.chars().find(|character| {
        !matches!(*character, '\u{9}' | '\u{a}' | '\u{d}')
            && (*character < '\u{20}' || matches!(*character, '\u{fffe}' | '\u{ffff}'))
    }) {
        return Err(editor_error(
            "use.office.spreadsheet_validation_text_invalid",
            format!(
                "Data-validation {field} contains XML-forbidden character U+{:04X}.",
                u32::from(character)
            ),
        )
        .with_detail("field", field));
    }
    Ok(())
}
