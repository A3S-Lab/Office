use std::cmp::Ordering;

use a3s_use_core::UseResult;
use serde_json::Value as JsonValue;

use super::invalid_shared_comments;

const MAX_JSON_DEPTH: usize = 128;

pub(in crate::collaboration) fn canonical_json(value: &JsonValue) -> UseResult<String> {
    let mut output = String::new();
    write_canonical_json(value, &mut output, 0)?;
    Ok(output)
}

fn write_canonical_json(value: &JsonValue, output: &mut String, depth: usize) -> UseResult<()> {
    if depth > MAX_JSON_DEPTH {
        return Err(invalid_shared_comments(
            "The shared Document comment JSON exceeds the nesting limit.",
        ));
    }
    match value {
        JsonValue::Null => output.push_str("null"),
        JsonValue::Bool(value) => output.push_str(if *value { "true" } else { "false" }),
        JsonValue::Number(value) => {
            let value = value.as_f64().ok_or_else(|| {
                invalid_shared_comments("A shared Document comment JSON number is unsupported.")
            })?;
            if !value.is_finite() {
                return Err(invalid_shared_comments(
                    "A shared Document comment JSON number is not finite.",
                ));
            }
            output.push_str(ryu_js::Buffer::new().format_finite(value));
        }
        JsonValue::String(value) => {
            output.push_str(&serde_json::to_string(value).map_err(|error| {
                invalid_shared_comments(format!(
                    "Failed to encode shared Document comment JSON: {error}"
                ))
            })?);
        }
        JsonValue::Array(values) => {
            output.push('[');
            for (index, value) in values.iter().enumerate() {
                if index > 0 {
                    output.push(',');
                }
                write_canonical_json(value, output, depth + 1)?;
            }
            output.push(']');
        }
        JsonValue::Object(values) => {
            output.push('{');
            let mut keys = values.keys().collect::<Vec<_>>();
            keys.sort_by(|left, right| utf16_cmp(left, right));
            for (index, key) in keys.into_iter().enumerate() {
                if index > 0 {
                    output.push(',');
                }
                output.push_str(&serde_json::to_string(key).map_err(|error| {
                    invalid_shared_comments(format!(
                        "Failed to encode shared Document comment JSON key: {error}"
                    ))
                })?);
                output.push(':');
                write_canonical_json(&values[key], output, depth + 1)?;
            }
            output.push('}');
        }
    }
    Ok(())
}

fn utf16_cmp(left: &str, right: &str) -> Ordering {
    let mut left = left.encode_utf16();
    let mut right = right.encode_utf16();
    loop {
        match (left.next(), right.next()) {
            (Some(left), Some(right)) => match left.cmp(&right) {
                Ordering::Equal => {}
                ordering => return ordering,
            },
            (None, Some(_)) => return Ordering::Less,
            (Some(_), None) => return Ordering::Greater,
            (None, None) => return Ordering::Equal,
        }
    }
}
