use std::cmp::Ordering;
use std::collections::BTreeSet;
use std::sync::Arc;

use a3s_use_core::UseResult;
use serde_json::{Map as JsonMap, Number as JsonNumber, Value as JsonValue};
use yrs::Any;

use super::{
    invalid_presentation_mutation, invalid_shared_presentation, presentation_match_conflict,
};

const MAX_PRESENTATION_ELEMENT_JSON_BYTES: usize = 8 * 1024 * 1024;
const MAX_JSON_DEPTH: usize = 128;

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct PresentationElementIdentity {
    pub(super) id: String,
    pub(super) element_type: String,
}

pub(super) enum PresentationElementPatch {
    Remove(String),
    Set(String, Any),
}

pub(super) fn validate_element_input(
    value: &JsonValue,
    label: &str,
) -> UseResult<PresentationElementIdentity> {
    validate_element(value, label, false)
}

pub(super) fn validate_shared_element(
    value: &JsonValue,
    label: &str,
) -> UseResult<PresentationElementIdentity> {
    validate_element(value, label, true)
}

fn validate_element(
    value: &JsonValue,
    label: &str,
    shared: bool,
) -> UseResult<PresentationElementIdentity> {
    validate_json(value, label, 0, shared)?;
    let object = value.as_object().ok_or_else(|| {
        invalid_element(
            shared,
            format!("A Presentation {label} must be a JSON object."),
        )
    })?;
    if object.contains_key("tombstone") {
        return Err(invalid_element(
            shared,
            format!("A Presentation {label} contains the reserved tombstone field."),
        ));
    }
    let id = object
        .get("id")
        .and_then(JsonValue::as_str)
        .ok_or_else(|| {
            invalid_element(
                shared,
                format!("A Presentation {label} must contain a string ID."),
            )
        })?;
    validate_presentation_identifier(id, label, shared)?;
    let element_type = object
        .get("type")
        .and_then(JsonValue::as_str)
        .ok_or_else(|| {
            invalid_element(
                shared,
                format!("A Presentation {label} must contain a string type."),
            )
        })?;
    if !matches!(
        element_type,
        "text" | "shape" | "image" | "table" | "chart" | "line"
    ) {
        return Err(invalid_element(
            shared,
            format!("A Presentation {label} contains an unsupported type."),
        ));
    }
    for key in ["x", "y", "width", "height", "fontSize"] {
        if object
            .get(key)
            .and_then(JsonValue::as_f64)
            .is_none_or(|value| !value.is_finite())
        {
            return Err(invalid_element(
                shared,
                format!("A Presentation {label} must contain a finite numeric '{key}' field."),
            ));
        }
    }
    for key in ["text", "color", "fill"] {
        if !object.get(key).is_some_and(JsonValue::is_string) {
            return Err(invalid_element(
                shared,
                format!("A Presentation {label} must contain a string '{key}' field."),
            ));
        }
    }
    if !object.get("bold").is_some_and(JsonValue::is_boolean) {
        return Err(invalid_element(
            shared,
            format!("A Presentation {label} must contain a boolean 'bold' field."),
        ));
    }
    if !matches!(
        object.get("align").and_then(JsonValue::as_str),
        Some("left" | "center" | "right")
    ) {
        return Err(invalid_element(
            shared,
            format!("A Presentation {label} contains an unsupported text alignment."),
        ));
    }
    Ok(PresentationElementIdentity {
        id: id.to_owned(),
        element_type: element_type.to_owned(),
    })
}

pub(super) fn validate_presentation_identifier(
    value: &str,
    label: &str,
    shared: bool,
) -> UseResult<()> {
    let length = value.encode_utf16().count();
    let trimmed = value
        .chars()
        .next()
        .is_some_and(is_ecmascript_trim_character)
        || value
            .chars()
            .next_back()
            .is_some_and(is_ecmascript_trim_character);
    if (1..=256).contains(&length) && !trimmed {
        return Ok(());
    }
    Err(invalid_element(
        shared,
        format!(
            "A Presentation {label} ID must contain 1 to 256 UTF-16 code units without leading or trailing whitespace."
        ),
    ))
}

pub(super) fn assert_compatible_element(
    previous: &JsonValue,
    next: &JsonValue,
    shared: &JsonValue,
    label: &str,
) -> UseResult<()> {
    let previous = previous.as_object().expect("validated expected element");
    let next = next.as_object().expect("validated next element");
    let shared = shared.as_object().expect("validated shared element");
    let keys = previous
        .keys()
        .chain(next.keys())
        .cloned()
        .collect::<BTreeSet<_>>();
    for key in keys {
        let before = previous.get(&key);
        let after = next.get(&key);
        let current = shared.get(&key);
        if optional_json_equal(before, after) {
            continue;
        }
        let field_label = format!("{label} field '{key}'");
        match (before, after, current) {
            (Some(before), None, Some(current)) if !json_equal(before, current) => {
                return Err(presentation_match_conflict(format!(
                    "The {field_label} changed before it could be removed."
                )))
            }
            (_, None, _) => {}
            (None, Some(after), Some(current)) if !json_equal(after, current) => {
                return Err(presentation_match_conflict(format!(
                    "The {field_label} was added concurrently with a different value."
                )))
            }
            (None, Some(_), _) => {}
            (Some(_), Some(_), None) => {
                return Err(presentation_match_conflict(format!(
                    "The {field_label} was removed before this change could be applied."
                )))
            }
            (Some(before), Some(after), Some(current))
                if !json_equal(before, current) && !json_equal(after, current) =>
            {
                return Err(presentation_match_conflict(format!(
                    "The {field_label} changed concurrently."
                )))
            }
            _ => {}
        }
    }
    Ok(())
}

pub(super) fn element_field_patches(
    previous: &JsonValue,
    shared: &JsonValue,
    next: &JsonValue,
) -> UseResult<Vec<PresentationElementPatch>> {
    let previous = previous.as_object().expect("validated expected element");
    let shared = shared.as_object().expect("validated shared element");
    let next = next.as_object().expect("validated next element");
    let keys = previous
        .keys()
        .chain(next.keys())
        .cloned()
        .collect::<BTreeSet<_>>();
    let mut patches = Vec::new();
    for key in keys {
        if optional_json_equal(previous.get(&key), next.get(&key))
            || optional_json_equal(shared.get(&key), next.get(&key))
        {
            continue;
        }
        patches.push(match next.get(&key) {
            Some(value) => PresentationElementPatch::Set(key, json_to_any(value, 0)?),
            None => PresentationElementPatch::Remove(key),
        });
    }
    Ok(patches)
}

pub(super) fn any_to_json(value: Any, depth: usize, label: &str) -> UseResult<JsonValue> {
    if depth > MAX_JSON_DEPTH {
        return Err(invalid_shared_presentation(format!(
            "The shared Presentation {label} exceeds the JSON nesting limit."
        )));
    }
    match value {
        Any::Null => Ok(JsonValue::Null),
        Any::Bool(value) => Ok(JsonValue::Bool(value)),
        Any::Number(value) => JsonNumber::from_f64(value)
            .map(JsonValue::Number)
            .ok_or_else(|| {
                invalid_shared_presentation(format!(
                    "The shared Presentation {label} contains a non-finite number."
                ))
            }),
        Any::String(value) => Ok(JsonValue::String(value.to_string())),
        Any::Array(values) => values
            .iter()
            .cloned()
            .map(|value| any_to_json(value, depth + 1, label))
            .collect::<UseResult<Vec<_>>>()
            .map(JsonValue::Array),
        Any::Map(values) => {
            let mut object = JsonMap::new();
            for (key, value) in values.iter() {
                if unsafe_json_key(key) {
                    return Err(invalid_shared_presentation(format!(
                        "The shared Presentation {label} contains an unsafe JSON key."
                    )));
                }
                object.insert(key.clone(), any_to_json(value.clone(), depth + 1, label)?);
            }
            Ok(JsonValue::Object(object))
        }
        Any::Undefined | Any::BigInt(_) | Any::Buffer(_) => Err(invalid_shared_presentation(
            format!("The shared Presentation {label} contains a non-JSON value."),
        )),
    }
}

pub(super) fn json_to_any(value: &JsonValue, depth: usize) -> UseResult<Any> {
    if depth > MAX_JSON_DEPTH {
        return Err(invalid_presentation_mutation(
            "A Presentation scene element exceeds the JSON nesting limit.",
        ));
    }
    match value {
        JsonValue::Null => Ok(Any::Null),
        JsonValue::Bool(value) => Ok(Any::Bool(*value)),
        JsonValue::Number(value) => value.as_f64().map(Any::Number).ok_or_else(|| {
            invalid_presentation_mutation(
                "A Presentation scene element contains a non-finite number.",
            )
        }),
        JsonValue::String(value) => Ok(Any::String(Arc::from(value.as_str()))),
        JsonValue::Array(values) => values
            .iter()
            .map(|value| json_to_any(value, depth + 1))
            .collect::<UseResult<Vec<_>>>()
            .map(|values| Any::Array(Arc::from(values))),
        JsonValue::Object(values) => {
            let mut object = std::collections::HashMap::with_capacity(values.len());
            for (key, value) in values {
                object.insert(key.clone(), json_to_any(value, depth + 1)?);
            }
            Ok(Any::Map(Arc::new(object)))
        }
    }
}

pub(super) fn canonical_json(value: &JsonValue) -> UseResult<String> {
    let mut output = String::new();
    write_canonical_json(value, &mut output, 0)?;
    Ok(output)
}

pub(super) fn json_equal(left: &JsonValue, right: &JsonValue) -> bool {
    match (left, right) {
        (JsonValue::Null, JsonValue::Null) => true,
        (JsonValue::Bool(left), JsonValue::Bool(right)) => left == right,
        (JsonValue::Number(left), JsonValue::Number(right)) => left
            .as_f64()
            .zip(right.as_f64())
            .is_some_and(|(left, right)| left.to_bits() == right.to_bits()),
        (JsonValue::String(left), JsonValue::String(right)) => left == right,
        (JsonValue::Array(left), JsonValue::Array(right)) => {
            left.len() == right.len()
                && left
                    .iter()
                    .zip(right)
                    .all(|(left, right)| json_equal(left, right))
        }
        (JsonValue::Object(left), JsonValue::Object(right)) => {
            left.len() == right.len()
                && left
                    .iter()
                    .all(|(key, left)| right.get(key).is_some_and(|right| json_equal(left, right)))
        }
        _ => false,
    }
}

fn validate_json(value: &JsonValue, label: &str, depth: usize, shared: bool) -> UseResult<()> {
    if depth == 0 {
        let bytes = serde_json::to_vec(value)
            .map_err(|error| invalid_element(shared, format!("Failed to encode {label}: {error}")))?
            .len();
        if bytes > MAX_PRESENTATION_ELEMENT_JSON_BYTES {
            return Err(invalid_element(
                shared,
                format!(
                    "A Presentation {label} is {bytes} bytes; the limit is {MAX_PRESENTATION_ELEMENT_JSON_BYTES} bytes."
                ),
            ));
        }
    }
    if depth > MAX_JSON_DEPTH {
        return Err(invalid_element(
            shared,
            format!("A Presentation {label} exceeds the JSON nesting limit."),
        ));
    }
    match value {
        JsonValue::Array(values) => {
            for value in values {
                validate_json(value, label, depth + 1, shared)?;
            }
        }
        JsonValue::Object(values) => {
            for (key, value) in values {
                if unsafe_json_key(key) {
                    return Err(invalid_element(
                        shared,
                        format!("A Presentation {label} contains an unsafe JSON key."),
                    ));
                }
                validate_json(value, label, depth + 1, shared)?;
            }
        }
        JsonValue::Number(value) if value.as_f64().is_none_or(|value| !value.is_finite()) => {
            return Err(invalid_element(
                shared,
                format!("A Presentation {label} contains a non-finite number."),
            ));
        }
        _ => {}
    }
    Ok(())
}

fn write_canonical_json(value: &JsonValue, output: &mut String, depth: usize) -> UseResult<()> {
    if depth > MAX_JSON_DEPTH {
        return Err(invalid_shared_presentation(
            "The shared Presentation canonical JSON exceeds the nesting limit.",
        ));
    }
    match value {
        JsonValue::Null => output.push_str("null"),
        JsonValue::Bool(value) => output.push_str(if *value { "true" } else { "false" }),
        JsonValue::Number(value) => {
            let value = value.as_f64().ok_or_else(|| {
                invalid_shared_presentation(
                    "The shared Presentation canonical JSON number is unsupported.",
                )
            })?;
            if !value.is_finite() {
                return Err(invalid_shared_presentation(
                    "The shared Presentation canonical JSON number is not finite.",
                ));
            }
            output.push_str(ryu_js::Buffer::new().format_finite(value));
        }
        JsonValue::String(value) => {
            output.push_str(&serde_json::to_string(value).map_err(|error| {
                invalid_shared_presentation(format!(
                    "Failed to encode shared Presentation JSON: {error}"
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
                    invalid_shared_presentation(format!(
                        "Failed to encode shared Presentation JSON: {error}"
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

fn optional_json_equal(left: Option<&JsonValue>, right: Option<&JsonValue>) -> bool {
    match (left, right) {
        (Some(left), Some(right)) => json_equal(left, right),
        (None, None) => true,
        _ => false,
    }
}

fn invalid_element(shared: bool, message: String) -> a3s_use_core::UseError {
    if shared {
        invalid_shared_presentation(message)
    } else {
        invalid_presentation_mutation(message)
    }
}

fn unsafe_json_key(value: &str) -> bool {
    matches!(value, "__proto__" | "constructor" | "prototype")
}

fn is_ecmascript_trim_character(character: char) -> bool {
    matches!(
        character,
        '\u{0009}'
            | '\u{000A}'
            | '\u{000B}'
            | '\u{000C}'
            | '\u{000D}'
            | '\u{0020}'
            | '\u{00A0}'
            | '\u{1680}'
            | '\u{2000}'
            ..='\u{200A}'
                | '\u{2028}'
                | '\u{2029}'
                | '\u{202F}'
                | '\u{205F}'
                | '\u{3000}'
                | '\u{FEFF}'
    )
}
