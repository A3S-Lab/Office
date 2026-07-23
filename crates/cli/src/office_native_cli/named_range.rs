use a3s_office::{DocumentNode, NativeSpreadsheetNamedRange, NativeSpreadsheetNamedRangeScope};
use a3s_use_core::UseResult;

use super::arguments::{parse_boolean_option, ParsedArguments};
use super::usage_error;

pub(super) fn build_new(
    parent: &str,
    parsed: &ParsedArguments,
) -> UseResult<NativeSpreadsheetNamedRange> {
    let name = parsed
        .name
        .as_deref()
        .ok_or_else(|| usage_error("native named-range add requires --name <name>"))?;
    let reference = parsed
        .named_range_ref
        .as_deref()
        .ok_or_else(|| usage_error("native named-range add requires --ref <expression>"))?;
    let scope = parsed
        .named_range_scope
        .as_deref()
        .map(parse_scope)
        .transpose()?
        .unwrap_or_else(|| default_scope(parent));
    let mut named_range = NativeSpreadsheetNamedRange::new(name, reference).with_scope(scope);
    if let Some(comment) = optional_comment(parsed.named_range_comment.as_deref()) {
        named_range = named_range.with_comment(comment);
    }
    if let Some(volatile) = parsed.named_range_volatile.as_deref() {
        named_range = named_range.with_volatile(parse_boolean_option(volatile)?);
    }
    Ok(named_range)
}

pub(super) fn merge_existing(
    node: &DocumentNode,
    parsed: &ParsedArguments,
) -> UseResult<NativeSpreadsheetNamedRange> {
    let existing = |key: &str| {
        node.format.get(key).map(String::as_str).ok_or_else(|| {
            usage_error(format!(
                "existing named range '{}' has no {key} property",
                node.path
            ))
        })
    };
    let name = parsed.name.as_deref().unwrap_or(existing("name")?);
    let reference = parsed
        .named_range_ref
        .as_deref()
        .unwrap_or(existing("ref")?);
    let scope = parsed
        .named_range_scope
        .as_deref()
        .unwrap_or(existing("scope")?);
    let mut named_range =
        NativeSpreadsheetNamedRange::new(name, reference).with_scope(parse_scope(scope)?);
    let comment = if parsed.named_range_comment.is_some() {
        optional_comment(parsed.named_range_comment.as_deref())
    } else {
        node.format.get("comment").map(String::as_str)
    };
    if let Some(comment) = comment {
        named_range = named_range.with_comment(comment);
    }
    let volatile = parsed
        .named_range_volatile
        .as_deref()
        .map(parse_boolean_option)
        .transpose()?
        .unwrap_or_else(|| {
            node.format
                .get("volatile")
                .is_some_and(|value| value == "true")
        });
    Ok(named_range.with_volatile(volatile))
}

pub(super) fn is_path(path: &str) -> bool {
    let normalized = path.trim_start_matches('/').to_ascii_lowercase();
    ["namedrange", "definedname", "name"]
        .iter()
        .any(|prefix| normalized == *prefix || normalized.starts_with(&format!("{prefix}[")))
}

fn default_scope(parent: &str) -> NativeSpreadsheetNamedRangeScope {
    let parent = parent.trim_matches('/');
    if parent.is_empty()
        || parent.eq_ignore_ascii_case("workbook")
        || ["namedrange", "definedname", "name"]
            .iter()
            .any(|prefix| parent.to_ascii_lowercase().starts_with(prefix))
    {
        NativeSpreadsheetNamedRangeScope::Workbook
    } else {
        NativeSpreadsheetNamedRangeScope::worksheet(parent)
    }
}

fn parse_scope(scope: &str) -> UseResult<NativeSpreadsheetNamedRangeScope> {
    NativeSpreadsheetNamedRangeScope::try_from(scope.to_string()).map_err(usage_error)
}

fn optional_comment(comment: Option<&str>) -> Option<&str> {
    comment.filter(|comment| !comment.is_empty() && !comment.eq_ignore_ascii_case("none"))
}
