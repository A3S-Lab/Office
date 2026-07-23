use a3s_use_core::{UseError, UseResult};

use crate::discovery::office_error;
use crate::spreadsheet_reference::CellReference;

pub(crate) const MAX_SPREADSHEET_NAMED_RANGES: usize = 65_536;
pub(crate) const MAX_SPREADSHEET_NAMED_RANGE_NAME_CHARS: usize = 255;
pub(crate) const MAX_SPREADSHEET_NAMED_RANGE_REF_CHARS: usize = 8_192;
pub(crate) const MAX_SPREADSHEET_NAMED_RANGE_COMMENT_CHARS: usize = 255;

#[derive(Debug, Clone, PartialEq, Eq)]
pub(crate) enum NamedRangePathSelector {
    Collection,
    Position(usize),
    Name { name: String, scope: Option<String> },
}

pub(crate) fn canonical_named_range_path(name: &str, scope: &str) -> String {
    format!(
        "/namedrange[@name={}][@scope={}]",
        encode_path_value(name),
        encode_path_value(scope)
    )
}

pub(crate) fn named_range_scope_label(scope: &str, worksheet_local: bool) -> String {
    if worksheet_local && scope.eq_ignore_ascii_case("workbook") {
        format!("worksheet:{scope}")
    } else {
        scope.to_string()
    }
}

pub(crate) fn parse_named_range_path(path: &str) -> UseResult<Option<NamedRangePathSelector>> {
    let Some(segment) = path.strip_prefix('/') else {
        return Ok(None);
    };
    if segment.contains('/') {
        return Ok(None);
    }
    let base_end = segment.find('[').unwrap_or(segment.len());
    let base = &segment[..base_end];
    if !matches!(
        base.to_ascii_lowercase().as_str(),
        "namedrange" | "definedname" | "name"
    ) {
        return Ok(None);
    }
    let suffix = &segment[base_end..];
    if suffix.is_empty() {
        return Ok(Some(NamedRangePathSelector::Collection));
    }
    let selectors = bracket_values(suffix)?;
    if selectors.len() == 1 && !selectors[0].starts_with('@') {
        if let Ok(position) = selectors[0].parse::<usize>() {
            if position == 0 {
                return Err(path_error("Named-range positions are one-based."));
            }
            return Ok(Some(NamedRangePathSelector::Position(position)));
        }
        return Ok(Some(NamedRangePathSelector::Name {
            name: decode_path_value(&selectors[0])?,
            scope: None,
        }));
    }

    let mut name = None;
    let mut scope = None;
    for selector in selectors {
        let Some((key, value)) = selector
            .strip_prefix('@')
            .and_then(|value| value.split_once('='))
        else {
            return Err(path_error(
                "Named-range stable selectors use [@name=NAME] and optional [@scope=SCOPE].",
            ));
        };
        if value.is_empty() {
            return Err(path_error("Named-range selector values cannot be empty."));
        }
        match key.to_ascii_lowercase().as_str() {
            "name" if name.is_none() => name = Some(decode_path_value(value)?),
            "scope" if scope.is_none() => scope = Some(decode_path_value(value)?),
            _ => {
                return Err(path_error(format!(
                    "Named-range selector attribute '@{key}' is unsupported or repeated."
                )))
            }
        }
    }
    let name = name.ok_or_else(|| path_error("Named-range stable selectors require @name."))?;
    Ok(Some(NamedRangePathSelector::Name { name, scope }))
}

pub(crate) fn validate_named_range_name(name: &str) -> UseResult<()> {
    let mut characters = name.chars();
    let first = characters.next();
    let valid_first =
        first.is_some_and(|character| character.is_alphabetic() || matches!(character, '_' | '\\'));
    let valid_rest = characters
        .all(|character| character.is_alphanumeric() || matches!(character, '_' | '.' | '\\'));
    if !valid_first
        || !valid_rest
        || name.chars().count() > MAX_SPREADSHEET_NAMED_RANGE_NAME_CHARS
        || CellReference::parse(name).is_ok()
        || looks_like_r1c1_reference(name)
    {
        return Err(office_error(
            "use.office.spreadsheet_named_range_name_invalid",
            format!(
                "Spreadsheet defined name '{name}' must contain 1-{MAX_SPREADSHEET_NAMED_RANGE_NAME_CHARS} letters, digits, underscores, periods, or backslashes, begin with a letter, underscore, or backslash, and not resemble A1 or R1C1 notation."
            ),
        )
        .with_detail("name", name));
    }
    Ok(())
}

pub(crate) fn validate_named_range_reference(reference: &str) -> UseResult<()> {
    if reference.is_empty()
        || reference.trim() != reference
        || reference.starts_with('=')
        || reference.chars().count() > MAX_SPREADSHEET_NAMED_RANGE_REF_CHARS
        || reference.chars().any(char::is_control)
    {
        return Err(office_error(
            "use.office.spreadsheet_named_range_ref_invalid",
            format!(
                "Spreadsheet named-range refs must contain 1-{MAX_SPREADSHEET_NAMED_RANGE_REF_CHARS} non-control characters without surrounding whitespace or a leading '='."
            ),
        )
        .with_detail("ref", reference));
    }
    let probe = reference.trim_start_matches('\'');
    if probe.starts_with('[')
        && probe.find(']').is_some_and(|end| {
            let book = &probe[1..end];
            book.bytes().all(|byte| byte.is_ascii_digit())
                || [".xls", ".xlsx", ".xlsm", ".xlsb"]
                    .iter()
                    .any(|suffix| book.to_ascii_lowercase().ends_with(suffix))
        })
    {
        return Err(office_error(
            "use.office.spreadsheet_named_range_ref_invalid",
            "Cross-workbook named-range refs require externalLinks parts and are not supported by the typed native editor.",
        )
        .with_detail("ref", reference));
    }
    Ok(())
}

pub(crate) fn validate_named_range_comment(comment: Option<&str>) -> UseResult<()> {
    let Some(comment) = comment else {
        return Ok(());
    };
    if comment.chars().count() > MAX_SPREADSHEET_NAMED_RANGE_COMMENT_CHARS
        || comment.chars().any(char::is_control)
    {
        return Err(office_error(
            "use.office.spreadsheet_named_range_comment_invalid",
            format!(
                "Spreadsheet named-range comments may contain at most {MAX_SPREADSHEET_NAMED_RANGE_COMMENT_CHARS} non-control characters."
            ),
        ));
    }
    Ok(())
}

pub(crate) fn is_protected_named_range(name: &str) -> bool {
    let normalized = name.to_ascii_lowercase();
    normalized.starts_with("_xlnm.") || normalized.starts_with("slicer_")
}

pub(crate) fn protected_named_range_error(name: &str) -> UseError {
    office_error(
        "use.office.spreadsheet_named_range_reserved",
        format!(
            "Spreadsheet defined name '{name}' is reserved for an Office-managed print, filter, or slicer feature."
        ),
    )
    .with_suggestion("Manage the owning Office feature instead of editing its internal defined name.")
    .with_detail("name", name)
}

pub(crate) fn quote_sheet_name(name: &str) -> String {
    format!("'{}'", name.replace('\'', "''"))
}

fn bracket_values(suffix: &str) -> UseResult<Vec<String>> {
    let mut values = Vec::new();
    let mut remaining = suffix;
    while !remaining.is_empty() {
        let Some(content) = remaining.strip_prefix('[') else {
            return Err(path_error("Named-range selector brackets are malformed."));
        };
        let Some(end) = content.find(']') else {
            return Err(path_error("Named-range selector is missing ']'."));
        };
        if end == 0 {
            return Err(path_error("Named-range selectors cannot be empty."));
        }
        values.push(content[..end].to_string());
        remaining = &content[end + 1..];
    }
    Ok(values)
}

fn encode_path_value(value: &str) -> String {
    let mut output = String::with_capacity(value.len());
    for byte in value.as_bytes() {
        if byte.is_ascii_alphanumeric() || matches!(*byte, b'-' | b'_' | b'.') {
            output.push(char::from(*byte));
        } else {
            output.push('%');
            output.push(hex(byte >> 4));
            output.push(hex(byte & 0x0f));
        }
    }
    output
}

fn decode_path_value(value: &str) -> UseResult<String> {
    let bytes = value.as_bytes();
    let mut decoded = Vec::with_capacity(bytes.len());
    let mut index = 0;
    while index < bytes.len() {
        if bytes[index] != b'%' {
            decoded.push(bytes[index]);
            index += 1;
            continue;
        }
        let Some((high, low)) = bytes.get(index + 1).zip(bytes.get(index + 2)) else {
            return Err(path_error(
                "Named-range selector has incomplete percent encoding.",
            ));
        };
        let byte = unhex(*high)
            .and_then(|high| unhex(*low).map(|low| high << 4 | low))
            .ok_or_else(|| path_error("Named-range selector has invalid percent encoding."))?;
        decoded.push(byte);
        index += 3;
    }
    String::from_utf8(decoded)
        .map_err(|_| path_error("Named-range selector is not valid percent-encoded UTF-8."))
}

const fn hex(value: u8) -> char {
    match value {
        0..=9 => (b'0' + value) as char,
        _ => (b'A' + value - 10) as char,
    }
}

const fn unhex(value: u8) -> Option<u8> {
    match value {
        b'0'..=b'9' => Some(value - b'0'),
        b'a'..=b'f' => Some(value - b'a' + 10),
        b'A'..=b'F' => Some(value - b'A' + 10),
        _ => None,
    }
}

fn looks_like_r1c1_reference(name: &str) -> bool {
    let upper = name.to_ascii_uppercase();
    if matches!(upper.as_str(), "R" | "C") {
        return true;
    }
    let Some(rest) = upper.strip_prefix('R') else {
        return false;
    };
    let Some((row, column)) = rest.split_once('C') else {
        return false;
    };
    !row.is_empty()
        && !column.is_empty()
        && row.bytes().all(|byte| byte.is_ascii_digit())
        && column.bytes().all(|byte| byte.is_ascii_digit())
}

fn path_error(message: impl Into<String>) -> UseError {
    office_error("use.office.path_invalid", message)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn stable_paths_round_trip_unicode_and_reserved_characters() {
        let path = canonical_named_range_path("收入\\率", "Sales Q1's");
        assert_eq!(
            parse_named_range_path(&path).unwrap(),
            Some(NamedRangePathSelector::Name {
                name: "收入\\率".to_string(),
                scope: Some("Sales Q1's".to_string()),
            })
        );
    }

    #[test]
    fn name_validation_rejects_reference_notation() {
        for invalid in ["A1", "XFD1048576", "R", "C", "R1C1", "bad name"] {
            assert!(validate_named_range_name(invalid).is_err(), "{invalid}");
        }
        for valid in ["Revenue", "_value", "收入", "Tax.Rate", "\\hidden"] {
            validate_named_range_name(valid).unwrap();
        }
    }
}
