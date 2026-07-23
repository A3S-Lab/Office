use a3s_use_core::UseResult;

use crate::{
    DocumentNode, NativeSpreadsheetAutoFilter, NativeSpreadsheetDynamicFilter,
    NativeSpreadsheetFilterColumn, NativeSpreadsheetFilterCriteria, OfficeNodeType,
};

impl NativeSpreadsheetAutoFilter {
    /// Reconstructs one complete typed AutoFilter from a mutable semantic node.
    pub fn from_semantic_node(node: &DocumentNode) -> UseResult<Self> {
        if node.node_type != OfficeNodeType::AutoFilter
            || node.tag != "autofilter"
            || node.format.get("nativeMutable").map(String::as_str) != Some("true")
        {
            return Err(semantic_filter_error(
                node,
                "is not a natively mutable Spreadsheet AutoFilter node",
            ));
        }
        let mut columns = Vec::with_capacity(node.children.len());
        for child in &node.children {
            if child.node_type != OfficeNodeType::FilterColumn || child.tag != "filtercolumn" {
                return Err(semantic_filter_error(
                    node,
                    "contains an unsupported filter child node",
                ));
            }
            columns.push(NativeSpreadsheetFilterColumn {
                column: parse(child, "column")?,
                criteria: criteria_from_node(child)?,
            });
        }
        let filter = Self {
            range: required(node, "ref")?.to_string(),
            columns,
        };
        filter.validate()?;
        Ok(filter)
    }
}

fn criteria_from_node(node: &DocumentNode) -> UseResult<NativeSpreadsheetFilterCriteria> {
    Ok(match required(node, "criteriaType")? {
        "values" => NativeSpreadsheetFilterCriteria::Values {
            values: node
                .children
                .iter()
                .map(|value| {
                    if value.node_type == OfficeNodeType::FilterValue
                        && value.tag == "filtervalue"
                        && value.children.is_empty()
                    {
                        Ok(value.text.clone())
                    } else {
                        Err(semantic_filter_error(
                            node,
                            "contains an unsupported value-filter child",
                        ))
                    }
                })
                .collect::<UseResult<Vec<_>>>()?,
            include_blanks: boolean(node, "includeBlanks")?,
        },
        "equals" => NativeSpreadsheetFilterCriteria::Equals {
            value: required(node, "value")?.to_string(),
        },
        "not-equals" => NativeSpreadsheetFilterCriteria::NotEquals {
            value: required(node, "value")?.to_string(),
        },
        "contains" => NativeSpreadsheetFilterCriteria::Contains {
            value: required(node, "value")?.to_string(),
        },
        "does-not-contain" => NativeSpreadsheetFilterCriteria::DoesNotContain {
            value: required(node, "value")?.to_string(),
        },
        "begins-with" => NativeSpreadsheetFilterCriteria::BeginsWith {
            value: required(node, "value")?.to_string(),
        },
        "ends-with" => NativeSpreadsheetFilterCriteria::EndsWith {
            value: required(node, "value")?.to_string(),
        },
        "greater-than" => NativeSpreadsheetFilterCriteria::GreaterThan {
            value: required(node, "value")?.to_string(),
        },
        "greater-than-or-equal" => NativeSpreadsheetFilterCriteria::GreaterThanOrEqual {
            value: required(node, "value")?.to_string(),
        },
        "less-than" => NativeSpreadsheetFilterCriteria::LessThan {
            value: required(node, "value")?.to_string(),
        },
        "less-than-or-equal" => NativeSpreadsheetFilterCriteria::LessThanOrEqual {
            value: required(node, "value")?.to_string(),
        },
        "between" => NativeSpreadsheetFilterCriteria::Between {
            lower: required(node, "lower")?.to_string(),
            upper: required(node, "upper")?.to_string(),
        },
        "not-between" => NativeSpreadsheetFilterCriteria::NotBetween {
            lower: required(node, "lower")?.to_string(),
            upper: required(node, "upper")?.to_string(),
        },
        "blanks" => NativeSpreadsheetFilterCriteria::Blanks,
        "non-blanks" => NativeSpreadsheetFilterCriteria::NonBlanks,
        "top" => NativeSpreadsheetFilterCriteria::Top {
            count: parse(node, "count")?,
        },
        "top-percent" => NativeSpreadsheetFilterCriteria::TopPercent {
            percent: parse(node, "percent")?,
        },
        "bottom" => NativeSpreadsheetFilterCriteria::Bottom {
            count: parse(node, "count")?,
        },
        "bottom-percent" => NativeSpreadsheetFilterCriteria::BottomPercent {
            percent: parse(node, "percent")?,
        },
        "dynamic" => NativeSpreadsheetFilterCriteria::Dynamic {
            kind: NativeSpreadsheetDynamicFilter::from_ooxml_name(required(node, "dynamicKind")?)
                .ok_or_else(|| {
                semantic_filter_error(node, "has an unsupported dynamic filter kind")
            })?,
        },
        value => {
            return Err(semantic_filter_error(
                node,
                format!("has unsupported criteria type '{value}'"),
            ))
        }
    })
}

pub(crate) fn criteria_type(criteria: &NativeSpreadsheetFilterCriteria) -> &'static str {
    match criteria {
        NativeSpreadsheetFilterCriteria::Values { .. } => "values",
        NativeSpreadsheetFilterCriteria::Equals { .. } => "equals",
        NativeSpreadsheetFilterCriteria::NotEquals { .. } => "not-equals",
        NativeSpreadsheetFilterCriteria::Contains { .. } => "contains",
        NativeSpreadsheetFilterCriteria::DoesNotContain { .. } => "does-not-contain",
        NativeSpreadsheetFilterCriteria::BeginsWith { .. } => "begins-with",
        NativeSpreadsheetFilterCriteria::EndsWith { .. } => "ends-with",
        NativeSpreadsheetFilterCriteria::GreaterThan { .. } => "greater-than",
        NativeSpreadsheetFilterCriteria::GreaterThanOrEqual { .. } => "greater-than-or-equal",
        NativeSpreadsheetFilterCriteria::LessThan { .. } => "less-than",
        NativeSpreadsheetFilterCriteria::LessThanOrEqual { .. } => "less-than-or-equal",
        NativeSpreadsheetFilterCriteria::Between { .. } => "between",
        NativeSpreadsheetFilterCriteria::NotBetween { .. } => "not-between",
        NativeSpreadsheetFilterCriteria::Blanks => "blanks",
        NativeSpreadsheetFilterCriteria::NonBlanks => "non-blanks",
        NativeSpreadsheetFilterCriteria::Top { .. } => "top",
        NativeSpreadsheetFilterCriteria::TopPercent { .. } => "top-percent",
        NativeSpreadsheetFilterCriteria::Bottom { .. } => "bottom",
        NativeSpreadsheetFilterCriteria::BottomPercent { .. } => "bottom-percent",
        NativeSpreadsheetFilterCriteria::Dynamic { .. } => "dynamic",
    }
}

pub(crate) fn escape_wildcard_literal(value: &str) -> String {
    let mut output = String::with_capacity(value.len());
    for character in value.chars() {
        if matches!(character, '~' | '*' | '?') {
            output.push('~');
        }
        output.push(character);
    }
    output
}

pub(crate) fn decode_custom_pattern(
    value: &str,
    negative: bool,
) -> Option<NativeSpreadsheetFilterCriteria> {
    let tokens = wildcard_tokens(value);
    let leading = tokens
        .first()
        .is_some_and(|token| *token == WildcardToken::Many);
    let trailing = tokens
        .last()
        .is_some_and(|token| *token == WildcardToken::Many);
    let interior_wildcard = tokens.iter().enumerate().any(|(index, token)| {
        matches!(token, WildcardToken::Many | WildcardToken::One)
            && !(index == 0 && leading)
            && !(index + 1 == tokens.len() && trailing)
    });
    let literal = tokens
        .iter()
        .filter_map(|token| match token {
            WildcardToken::Literal(character) => Some(*character),
            WildcardToken::Many | WildcardToken::One => None,
        })
        .collect::<String>();
    if !interior_wildcard && leading && trailing && tokens.len() >= 2 {
        if negative {
            Some(NativeSpreadsheetFilterCriteria::DoesNotContain { value: literal })
        } else {
            Some(NativeSpreadsheetFilterCriteria::Contains { value: literal })
        }
    } else if !interior_wildcard && trailing && !leading {
        if negative {
            None
        } else {
            Some(NativeSpreadsheetFilterCriteria::BeginsWith { value: literal })
        }
    } else if !interior_wildcard && leading && !trailing {
        if negative {
            None
        } else {
            Some(NativeSpreadsheetFilterCriteria::EndsWith { value: literal })
        }
    } else if interior_wildcard {
        None
    } else if negative {
        Some(NativeSpreadsheetFilterCriteria::NotEquals {
            value: unescape_wildcards(value),
        })
    } else {
        Some(NativeSpreadsheetFilterCriteria::Equals {
            value: unescape_wildcards(value),
        })
    }
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
enum WildcardToken {
    Literal(char),
    Many,
    One,
}

fn wildcard_tokens(value: &str) -> Vec<WildcardToken> {
    let mut output = Vec::new();
    let mut escaped = false;
    for character in value.chars() {
        if escaped {
            output.push(WildcardToken::Literal(character));
            escaped = false;
        } else {
            match character {
                '~' => escaped = true,
                '*' => output.push(WildcardToken::Many),
                '?' => output.push(WildcardToken::One),
                _ => output.push(WildcardToken::Literal(character)),
            }
        }
    }
    if escaped {
        output.push(WildcardToken::Literal('~'));
    }
    output
}

fn unescape_wildcards(value: &str) -> String {
    wildcard_tokens(value)
        .into_iter()
        .map(|token| match token {
            WildcardToken::Literal(character) => character,
            WildcardToken::Many => '*',
            WildcardToken::One => '?',
        })
        .collect()
}

fn required<'a>(node: &'a DocumentNode, key: &str) -> UseResult<&'a str> {
    node.format
        .get(key)
        .map(String::as_str)
        .ok_or_else(|| semantic_filter_error(node, format!("has no '{key}' property")))
}

fn boolean(node: &DocumentNode, key: &str) -> UseResult<bool> {
    match required(node, key)? {
        "true" => Ok(true),
        "false" => Ok(false),
        value => Err(semantic_filter_error(
            node,
            format!("has non-boolean '{key}' value '{value}'"),
        )),
    }
}

fn parse<T>(node: &DocumentNode, key: &str) -> UseResult<T>
where
    T: std::str::FromStr,
    T::Err: std::fmt::Display,
{
    required(node, key)?
        .parse::<T>()
        .map_err(|error| semantic_filter_error(node, format!("has invalid '{key}' value: {error}")))
}

fn semantic_filter_error(node: &DocumentNode, reason: impl Into<String>) -> a3s_use_core::UseError {
    crate::discovery::office_error(
        "use.office.spreadsheet_filter_semantic_invalid",
        format!(
            "Spreadsheet AutoFilter node '{}' {}.",
            node.path,
            reason.into()
        ),
    )
    .with_detail("path", node.path.clone())
}
