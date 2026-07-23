use a3s_use_core::UseResult;

use crate::editor::{
    NativeSpreadsheetDelimitedImport, SpreadsheetCellValue, MAX_NATIVE_SPREADSHEET_IMPORT_CELLS,
};

const MAX_CELL_UTF16_UNITS: usize = 32_767;

#[derive(Debug)]
pub(super) struct ParsedImport {
    pub(super) rows: Vec<Vec<ParsedField>>,
    pub(super) max_columns: usize,
}

#[derive(Debug)]
pub(super) enum ParsedField {
    Empty,
    Value {
        value: SpreadsheetCellValue,
        date: bool,
    },
}

impl ParsedField {
    pub(super) fn is_empty(&self) -> bool {
        matches!(self, Self::Empty)
    }

    pub(super) fn value(&self) -> Option<(&SpreadsheetCellValue, bool)> {
        match self {
            Self::Empty => None,
            Self::Value { value, date } => Some((value, *date)),
        }
    }
}

pub(super) fn parse(
    request: &NativeSpreadsheetDelimitedImport,
    date_1904: bool,
) -> UseResult<ParsedImport> {
    let rows = parse_delimited(&request.content, request.format.delimiter())?;
    let max_columns = rows.iter().map(Vec::len).max().unwrap_or(0);
    let rows = rows
        .into_iter()
        .map(|row| {
            row.into_iter()
                .map(|field| parse_field(field, date_1904))
                .collect::<UseResult<Vec<_>>>()
        })
        .collect::<UseResult<Vec<_>>>()?;
    Ok(ParsedImport { rows, max_columns })
}

fn parse_delimited(content: &str, delimiter: char) -> UseResult<Vec<Vec<String>>> {
    let content = content.strip_prefix('\u{feff}').unwrap_or(content);
    if content.is_empty() {
        return Ok(Vec::new());
    }
    let mut rows = Vec::new();
    let mut row = Vec::new();
    let mut field = String::new();
    let mut in_quotes = false;
    let mut field_started = false;
    let mut quote_closed = false;
    let mut fields = 0_usize;
    let mut maximum_columns = 0_usize;
    let mut characters = content.chars().peekable();
    while let Some(character) = characters.next() {
        if in_quotes {
            if character == '"' {
                if characters.peek() == Some(&'"') {
                    field.push('"');
                    characters.next();
                } else {
                    in_quotes = false;
                    quote_closed = true;
                }
            } else {
                field.push(character);
            }
            continue;
        }
        if character == '"' && !field_started {
            in_quotes = true;
            field_started = true;
        } else if character == delimiter {
            row.push(std::mem::take(&mut field));
            fields = checked_field_count(fields)?;
            field_started = false;
            quote_closed = false;
        } else if character == '\r' {
            row.push(std::mem::take(&mut field));
            fields = checked_field_count(fields)?;
            push_row(&mut rows, std::mem::take(&mut row), &mut maximum_columns)?;
            field_started = false;
            quote_closed = false;
            if characters.peek() == Some(&'\n') {
                characters.next();
            }
        } else if character == '\n' {
            row.push(std::mem::take(&mut field));
            fields = checked_field_count(fields)?;
            push_row(&mut rows, std::mem::take(&mut row), &mut maximum_columns)?;
            field_started = false;
            quote_closed = false;
        } else if character == '"' {
            return Err(delimited_invalid(
                &rows,
                &row,
                "A quote may appear only at the beginning of a delimited field.",
            ));
        } else if quote_closed {
            return Err(delimited_invalid(
                &rows,
                &row,
                "Only a delimiter or record boundary may follow a closing quote.",
            ));
        } else {
            field.push(character);
            field_started = true;
        }
    }
    if in_quotes {
        return Err(delimited_invalid(
            &rows,
            &row,
            "Delimited input ended before a quoted field was closed.",
        ));
    }
    if field_started || !row.is_empty() {
        row.push(field);
        checked_field_count(fields)?;
        push_row(&mut rows, row, &mut maximum_columns)?;
    }
    Ok(rows)
}

fn checked_field_count(fields: usize) -> UseResult<usize> {
    let fields = fields.checked_add(1).ok_or_else(import_cell_count_limit)?;
    if fields > MAX_NATIVE_SPREADSHEET_IMPORT_CELLS {
        return Err(import_cell_count_limit().with_detail("fields", fields));
    }
    Ok(fields)
}

fn push_row(
    rows: &mut Vec<Vec<String>>,
    row: Vec<String>,
    maximum_columns: &mut usize,
) -> UseResult<()> {
    *maximum_columns = (*maximum_columns).max(row.len());
    let row_count = rows
        .len()
        .checked_add(1)
        .ok_or_else(import_cell_count_limit)?;
    let cells = row_count
        .checked_mul(*maximum_columns)
        .ok_or_else(import_cell_count_limit)?;
    if cells > MAX_NATIVE_SPREADSHEET_IMPORT_CELLS {
        return Err(import_cell_count_limit().with_detail("cells", cells));
    }
    rows.push(row);
    Ok(())
}

fn parse_field(value: String, date_1904: bool) -> UseResult<ParsedField> {
    if value.is_empty() {
        return Ok(ParsedField::Empty);
    }
    validate_cell_text(&value)?;
    let (value, date) = if let Some(expression) = value.strip_prefix('=') {
        (
            super::super::normalize_cell_value(&SpreadsheetCellValue::Formula {
                expression: expression.to_string(),
            })?,
            false,
        )
    } else if let Some(number) = normalize_number(&value) {
        (
            super::super::normalize_cell_value(&SpreadsheetCellValue::Number { value: number })?,
            false,
        )
    } else if let Some(serial) = parse_iso_date(&value, date_1904) {
        (
            super::super::normalize_cell_value(&SpreadsheetCellValue::Number { value: serial })?,
            true,
        )
    } else if value.eq_ignore_ascii_case("true") {
        (SpreadsheetCellValue::Boolean { value: true }, false)
    } else if value.eq_ignore_ascii_case("false") {
        (SpreadsheetCellValue::Boolean { value: false }, false)
    } else {
        (SpreadsheetCellValue::Text { value }, false)
    };
    Ok(ParsedField::Value { value, date })
}

fn validate_cell_text(value: &str) -> UseResult<()> {
    let units = value.encode_utf16().count();
    if units > MAX_CELL_UTF16_UNITS {
        return Err(import_error(
            "use.office.spreadsheet_import_cell_limit",
            format!(
                "Spreadsheet import fields cannot exceed {MAX_CELL_UTF16_UNITS} UTF-16 code units."
            ),
        )
        .with_detail("utf16Units", units));
    }
    if let Some(character) = value.chars().find(|character| {
        !matches!(
            u32::from(*character),
            0x9 | 0xA | 0xD | 0x20..=0xD7FF | 0xE000..=0xFFFD | 0x10000..=0x10FFFF
        )
    }) {
        return Err(import_error(
            "use.office.spreadsheet_import_cell_invalid",
            format!(
                "Spreadsheet import field contains XML-forbidden character U+{:04X}.",
                u32::from(character)
            ),
        ));
    }
    Ok(())
}

fn normalize_number(value: &str) -> Option<String> {
    let parsed = value.parse::<f64>().ok().filter(|value| value.is_finite());
    if let Some(parsed) = parsed {
        return Some(if is_canonical_number(value) {
            value.to_string()
        } else {
            parsed.to_string()
        });
    }

    let mut candidate = value.trim();
    let parenthesized = candidate.starts_with('(') && candidate.ends_with(')');
    if parenthesized {
        candidate = &candidate[1..candidate.len() - 1];
    }
    candidate = candidate.trim();
    if let Some(rest) = candidate.strip_prefix('$') {
        candidate = rest;
    }
    let normalized = candidate.replace(',', "");
    let parsed = normalized
        .parse::<f64>()
        .ok()
        .filter(|value| value.is_finite())?;
    let parsed = if parenthesized { -parsed } else { parsed };
    Some(parsed.to_string())
}

fn is_canonical_number(value: &str) -> bool {
    let value = value.strip_prefix('-').unwrap_or(value);
    if value.is_empty() {
        return false;
    }
    let mut exponent_parts = value.split(['e', 'E']);
    let mantissa = exponent_parts.next().unwrap_or_default();
    let exponent = exponent_parts.next();
    if exponent_parts.next().is_some()
        || exponent.is_some_and(|value| {
            let digits = value
                .strip_prefix('+')
                .or_else(|| value.strip_prefix('-'))
                .unwrap_or(value);
            digits.is_empty() || !digits.bytes().all(|byte| byte.is_ascii_digit())
        })
    {
        return false;
    }
    let mut decimal_parts = mantissa.split('.');
    let whole = decimal_parts.next().unwrap_or_default();
    let fraction = decimal_parts.next();
    if decimal_parts.next().is_some() {
        return false;
    }
    let whole_valid = whole.bytes().all(|byte| byte.is_ascii_digit());
    let fraction_valid = fraction.is_none_or(|digits| {
        digits.bytes().all(|byte| byte.is_ascii_digit())
            && (!whole.is_empty() || !digits.is_empty())
    });
    whole_valid && fraction_valid && (!whole.is_empty() || fraction.is_some())
}

fn parse_iso_date(value: &str, date_1904: bool) -> Option<String> {
    let (date, time) = if value.len() == 10 {
        (value, None)
    } else if value.len() >= 19 && matches!(value.as_bytes().get(10), Some(b'T' | b' ')) {
        (&value[..10], Some(&value[11..]))
    } else {
        return None;
    };
    let (year, month, day) = parse_date(date)?;
    let (hour, minute, second, millis) = match time {
        None => (0, 0, 0, 0),
        Some(time) => parse_time(time)?,
    };
    if year < 100 {
        return None;
    }
    let baseline = days_from_civil(1899, 12, 30);
    let mut serial = (days_from_civil(year, month, day) - baseline) as f64;
    serial += f64::from(hour * 3_600 + minute * 60 + second) / 86_400.0;
    serial += f64::from(millis) / 86_400_000.0;
    if (year, month, day) < (1900, 3, 1) && serial >= 2.0 {
        serial -= 1.0;
    }
    if date_1904 {
        serial -= 1_462.0;
    }
    Some(serial.to_string())
}

fn parse_date(value: &str) -> Option<(i32, u32, u32)> {
    if value.len() != 10 || value.as_bytes()[4] != b'-' || value.as_bytes()[7] != b'-' {
        return None;
    }
    let year = value[0..4].parse::<i32>().ok()?;
    let month = value[5..7].parse::<u32>().ok()?;
    let day = value[8..10].parse::<u32>().ok()?;
    ((1..=12).contains(&month) && day >= 1 && day <= days_in_month(year, month))
        .then_some((year, month, day))
}

fn parse_time(value: &str) -> Option<(u32, u32, u32, u32)> {
    let value = value.strip_suffix('Z').unwrap_or(value);
    let (clock, millis) = value.split_once('.').map_or((value, 0), |(clock, millis)| {
        if millis.len() != 3 {
            return (clock, u32::MAX);
        }
        (clock, millis.parse::<u32>().unwrap_or(u32::MAX))
    });
    if clock.len() != 8 || clock.as_bytes()[2] != b':' || clock.as_bytes()[5] != b':' {
        return None;
    }
    let hour = clock[0..2].parse::<u32>().ok()?;
    let minute = clock[3..5].parse::<u32>().ok()?;
    let second = clock[6..8].parse::<u32>().ok()?;
    (hour < 24 && minute < 60 && second < 60 && millis < 1_000)
        .then_some((hour, minute, second, millis))
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

fn import_error(code: &str, message: impl Into<String>) -> a3s_use_core::UseError {
    super::super::editor_error(code, message)
}

fn import_cell_count_limit() -> a3s_use_core::UseError {
    import_error(
        "use.office.spreadsheet_import_cell_count_limit",
        format!(
            "Spreadsheet import accepts at most {MAX_NATIVE_SPREADSHEET_IMPORT_CELLS} rectangular target cells."
        ),
    )
}

fn delimited_invalid(
    rows: &[Vec<String>],
    row: &[String],
    message: impl Into<String>,
) -> a3s_use_core::UseError {
    import_error("use.office.spreadsheet_import_delimited_invalid", message)
        .with_detail("row", rows.len() + 1)
        .with_detail("column", row.len() + 1)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn parser_matches_delimited_quote_and_blank_line_semantics() {
        assert_eq!(
            parse_delimited("a,\"b,c\"\r\n\r\n\"d\nq\",\"x\"\"y\"", ',').unwrap(),
            [
                vec!["a".to_string(), "b,c".to_string()],
                vec![String::new()],
                vec!["d\nq".to_string(), "x\"y".to_string()]
            ]
        );
        assert!(parse_delimited("", ',').unwrap().is_empty());
        assert_eq!(parse_delimited("\"\"", ',').unwrap(), [vec![String::new()]]);
    }

    #[test]
    fn parser_rejects_ambiguous_or_unclosed_quotes() {
        for input in ["\"unclosed", "\"closed\"suffix", "unquoted\"quote"] {
            let error = parse_delimited(input, ',').unwrap_err();
            assert_eq!(
                error.code,
                "use.office.spreadsheet_import_delimited_invalid"
            );
            assert_eq!(error.details["row"], 1);
            assert_eq!(error.details["column"], 1);
        }
    }

    #[test]
    fn iso_dates_follow_the_declared_excel_date_system() {
        let standard = parse_iso_date("2026-07-17T12:30:15.250Z", false)
            .unwrap()
            .parse::<f64>()
            .unwrap();
        let date_1904 = parse_iso_date("2026-07-17T12:30:15.250Z", true)
            .unwrap()
            .parse::<f64>()
            .unwrap();
        assert_eq!(standard - date_1904, 1_462.0);
        assert_eq!(parse_iso_date("1900-02-28", false).as_deref(), Some("59"));
        assert_eq!(parse_iso_date("1900-03-01", false).as_deref(), Some("61"));
        assert!(parse_iso_date("2026-02-29", false).is_none());
    }
}
