use std::collections::{BTreeMap, BTreeSet, HashMap, HashSet};
use std::sync::Arc;

use a3s_use_core::UseResult;
use serde_json::{Map as JsonMap, Number as JsonNumber, Value as JsonValue};
use yrs::Any;

use super::{invalid_shared_spreadsheet, invalid_spreadsheet_mutation, spreadsheet_match_conflict};

const MAX_SPREADSHEET_CELL_JSON_BYTES: usize = 8 * 1024 * 1024;
const MAX_JSON_DEPTH: usize = 128;

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum FlatJsonEntryKind {
    Object,
    Value,
}

impl FlatJsonEntryKind {
    const fn as_str(self) -> &'static str {
        match self {
            Self::Object => "object",
            Self::Value => "value",
        }
    }
}

pub(super) struct DecodedFlatJsonEntry {
    pub(super) kind: FlatJsonEntryKind,
    pub(super) path: Vec<String>,
    pub(super) value: JsonValue,
}

pub(super) enum FlatJsonPatch {
    Remove(String),
    Set(String, Any),
}

pub(super) fn validate_cell_json(value: &JsonValue, label: &str) -> UseResult<()> {
    if !value.is_object() {
        return Err(invalid_spreadsheet_mutation(format!(
            "A Spreadsheet {label} must be a JSON object."
        )));
    }
    let bytes = serde_json::to_vec(value)
        .map_err(|error| {
            invalid_spreadsheet_mutation(format!(
                "Failed to encode the Spreadsheet {label}: {error}"
            ))
        })?
        .len();
    if bytes > MAX_SPREADSHEET_CELL_JSON_BYTES {
        return Err(invalid_spreadsheet_mutation(format!(
            "A Spreadsheet {label} is {bytes} bytes; the per-cell limit is {MAX_SPREADSHEET_CELL_JSON_BYTES} bytes."
        )));
    }
    validate_json_value(value, label, 0)
}

pub(super) fn validate_shared_cell_json(value: &JsonValue) -> UseResult<()> {
    let bytes = serde_json::to_vec(value)
        .map_err(|error| {
            invalid_shared_spreadsheet(format!(
                "Failed to encode a shared Spreadsheet cell value: {error}"
            ))
        })?
        .len();
    if bytes > MAX_SPREADSHEET_CELL_JSON_BYTES {
        return Err(invalid_shared_spreadsheet(format!(
            "A shared Spreadsheet cell is {bytes} bytes; the supported limit is {MAX_SPREADSHEET_CELL_JSON_BYTES} bytes."
        )));
    }
    Ok(())
}

pub(super) fn assert_compatible_cell(
    previous: &JsonValue,
    next: &JsonValue,
    shared: &JsonValue,
    label: &str,
) -> UseResult<()> {
    if json_equal(previous, next) {
        return Ok(());
    }
    if let (Some(previous), Some(next), Some(shared)) =
        (previous.as_object(), next.as_object(), shared.as_object())
    {
        let keys = previous
            .keys()
            .chain(next.keys())
            .cloned()
            .collect::<BTreeSet<_>>();
        for key in keys {
            let before = previous.get(&key);
            let after = next.get(&key);
            let current = shared.get(&key);
            let field_label = format!("{label} field '{key}'");
            match (before, after, current) {
                (Some(before), None, Some(current)) if !json_equal(before, current) => {
                    return Err(spreadsheet_match_conflict(format!(
                        "The {field_label} changed before it could be removed."
                    )))
                }
                (_, None, _) => {}
                (None, Some(after), Some(current)) if !json_equal(after, current) => {
                    return Err(spreadsheet_match_conflict(format!(
                        "The {field_label} was added concurrently with a different value."
                    )))
                }
                (None, Some(_), _) => {}
                (Some(before), Some(after), _) if json_equal(before, after) => {}
                (Some(_), Some(_), None) => {
                    return Err(spreadsheet_match_conflict(format!(
                        "The {field_label} was removed before this change could be applied."
                    )))
                }
                (Some(before), Some(after), Some(current)) => {
                    assert_compatible_cell(before, after, current, &field_label)?;
                }
            }
        }
        return Ok(());
    }
    if !json_equal(previous, shared) && !json_equal(next, shared) {
        return Err(spreadsheet_match_conflict(format!(
            "The {label} changed concurrently."
        )));
    }
    Ok(())
}

pub(super) fn cell_field_patches(
    previous: &JsonValue,
    shared: &JsonValue,
    next: &JsonValue,
) -> UseResult<Vec<FlatJsonPatch>> {
    let before = flattened_cell(previous)?;
    let current = flattened_cell(shared)?;
    let after = flattened_cell(next)?;
    let keys = before
        .keys()
        .chain(after.keys())
        .cloned()
        .collect::<BTreeSet<_>>();
    let mut patches = Vec::new();
    for key in keys {
        if optional_json_equal(before.get(&key), after.get(&key))
            || optional_json_equal(current.get(&key), after.get(&key))
        {
            continue;
        }
        patches.push(match after.get(&key) {
            Some(value) => FlatJsonPatch::Set(key, json_to_any(value, 0)?),
            None => FlatJsonPatch::Remove(key),
        });
    }
    Ok(patches)
}

pub(super) fn flattened_cell(value: &JsonValue) -> UseResult<BTreeMap<String, JsonValue>> {
    let object = value.as_object().ok_or_else(|| {
        invalid_spreadsheet_mutation("A Spreadsheet cell mutation value must be a JSON object.")
    })?;
    let mut fields = BTreeMap::new();
    for (key, value) in object {
        flatten_json_value(&mut fields, &mut vec![key.clone()], value, 0)?;
    }
    Ok(fields)
}

pub(super) fn decode_flat_json_key(encoded: &str) -> UseResult<(FlatJsonEntryKind, Vec<String>)> {
    let identity = serde_json::from_str::<Vec<String>>(encoded).map_err(|_| {
        invalid_shared_spreadsheet("A shared Spreadsheet cell field identity is not valid JSON.")
    })?;
    let kind = match identity.first().map(String::as_str) {
        Some("object") => FlatJsonEntryKind::Object,
        Some("value") => FlatJsonEntryKind::Value,
        _ => {
            return Err(invalid_shared_spreadsheet(
                "A shared Spreadsheet cell field identity has an unsupported kind.",
            ))
        }
    };
    let path = identity[1..].to_vec();
    if path.is_empty()
        || path.len() > MAX_JSON_DEPTH
        || path.iter().any(|part| invalid_json_key(part))
        || encode_flat_json_key(kind, &path)? != encoded
    {
        return Err(invalid_shared_spreadsheet(
            "A shared Spreadsheet cell field identity is invalid.",
        ));
    }
    Ok((kind, path))
}

pub(super) fn reconstruct_cell(mut fields: Vec<DecodedFlatJsonEntry>) -> UseResult<JsonValue> {
    fields.sort_by(|left, right| {
        left.path.len().cmp(&right.path.len()).then_with(|| {
            if left.kind == right.kind {
                std::cmp::Ordering::Equal
            } else if left.kind == FlatJsonEntryKind::Object {
                std::cmp::Ordering::Less
            } else {
                std::cmp::Ordering::Greater
            }
        })
    });
    let object_paths = fields
        .iter()
        .filter(|field| field.kind == FlatJsonEntryKind::Object)
        .map(|field| field.path.clone())
        .collect::<HashSet<_>>();
    for field in &fields {
        for depth in 1..field.path.len() {
            if !object_paths.contains(&field.path[..depth]) {
                return Err(invalid_shared_spreadsheet(
                    "A shared Spreadsheet cell value is missing an object marker.",
                ));
            }
        }
    }

    let mut result = JsonMap::new();
    for field in fields {
        match field.kind {
            FlatJsonEntryKind::Object => ensure_object_path(&mut result, &field.path)?,
            FlatJsonEntryKind::Value => set_value_at_path(&mut result, &field.path, field.value)?,
        }
    }
    Ok(JsonValue::Object(result))
}

pub(super) fn any_to_json(value: Any, depth: usize) -> UseResult<JsonValue> {
    if depth > MAX_JSON_DEPTH {
        return Err(invalid_shared_spreadsheet(
            "A shared Spreadsheet cell value exceeds the JSON nesting limit.",
        ));
    }
    match value {
        Any::Null => Ok(JsonValue::Null),
        Any::Bool(value) => Ok(JsonValue::Bool(value)),
        Any::Number(value) => JsonNumber::from_f64(value)
            .map(JsonValue::Number)
            .ok_or_else(|| {
                invalid_shared_spreadsheet(
                    "A shared Spreadsheet cell value contains a non-finite number.",
                )
            }),
        Any::String(value) => Ok(JsonValue::String(value.to_string())),
        Any::Array(values) => values
            .iter()
            .cloned()
            .map(|value| any_to_json(value, depth + 1))
            .collect::<UseResult<Vec<_>>>()
            .map(JsonValue::Array),
        Any::Map(values) => {
            let mut object = JsonMap::new();
            for (key, value) in values.iter() {
                if invalid_json_key(key) {
                    return Err(invalid_shared_spreadsheet(
                        "A shared Spreadsheet cell value contains an empty or unsafe JSON key.",
                    ));
                }
                object.insert(key.clone(), any_to_json(value.clone(), depth + 1)?);
            }
            Ok(JsonValue::Object(object))
        }
        Any::Undefined | Any::BigInt(_) | Any::Buffer(_) => Err(invalid_shared_spreadsheet(
            "A shared Spreadsheet cell value contains a non-JSON value.",
        )),
    }
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

fn validate_json_value(value: &JsonValue, label: &str, depth: usize) -> UseResult<()> {
    if depth > MAX_JSON_DEPTH {
        return Err(invalid_spreadsheet_mutation(format!(
            "A Spreadsheet {label} exceeds the JSON nesting limit."
        )));
    }
    match value {
        JsonValue::Array(values) => {
            for value in values {
                validate_json_value(value, label, depth + 1)?;
            }
        }
        JsonValue::Object(values) => {
            for (key, value) in values {
                if invalid_json_key(key) {
                    return Err(invalid_spreadsheet_mutation(format!(
                        "A Spreadsheet {label} contains an empty or unsafe JSON key."
                    )));
                }
                validate_json_value(value, label, depth + 1)?;
            }
        }
        JsonValue::Number(value) if value.as_f64().is_none_or(|value| !value.is_finite()) => {
            return Err(invalid_spreadsheet_mutation(format!(
                "A Spreadsheet {label} contains a non-finite number."
            )))
        }
        _ => {}
    }
    Ok(())
}

fn flatten_json_value(
    fields: &mut BTreeMap<String, JsonValue>,
    path: &mut Vec<String>,
    value: &JsonValue,
    depth: usize,
) -> UseResult<()> {
    if depth > MAX_JSON_DEPTH {
        return Err(invalid_spreadsheet_mutation(
            "A Spreadsheet cell mutation exceeds the JSON nesting limit.",
        ));
    }
    if let JsonValue::Object(object) = value {
        fields.insert(
            encode_flat_json_key(FlatJsonEntryKind::Object, path)?,
            JsonValue::Bool(true),
        );
        for (key, value) in object {
            path.push(key.clone());
            flatten_json_value(fields, path, value, depth + 1)?;
            path.pop();
        }
    } else {
        fields.insert(
            encode_flat_json_key(FlatJsonEntryKind::Value, path)?,
            value.clone(),
        );
    }
    Ok(())
}

fn encode_flat_json_key(kind: FlatJsonEntryKind, path: &[String]) -> UseResult<String> {
    if path.is_empty() || path.iter().any(|part| invalid_json_key(part)) {
        return Err(invalid_spreadsheet_mutation(
            "A Spreadsheet cell field path contains an empty or unsafe key.",
        ));
    }
    let mut identity = Vec::with_capacity(path.len() + 1);
    identity.push(kind.as_str().to_owned());
    identity.extend_from_slice(path);
    serde_json::to_string(&identity).map_err(|error| {
        invalid_spreadsheet_mutation(format!(
            "Failed to encode a Spreadsheet cell field identity: {error}"
        ))
    })
}

fn json_to_any(value: &JsonValue, depth: usize) -> UseResult<Any> {
    if depth > MAX_JSON_DEPTH {
        return Err(invalid_spreadsheet_mutation(
            "A Spreadsheet cell mutation exceeds the JSON nesting limit.",
        ));
    }
    match value {
        JsonValue::Null => Ok(Any::Null),
        JsonValue::Bool(value) => Ok(Any::Bool(*value)),
        JsonValue::Number(value) => value.as_f64().map(Any::Number).ok_or_else(|| {
            invalid_spreadsheet_mutation(
                "A Spreadsheet cell mutation contains a non-finite number.",
            )
        }),
        JsonValue::String(value) => Ok(Any::String(Arc::from(value.as_str()))),
        JsonValue::Array(values) => values
            .iter()
            .map(|value| json_to_any(value, depth + 1))
            .collect::<UseResult<Vec<_>>>()
            .map(|values| Any::Array(Arc::from(values))),
        JsonValue::Object(values) => {
            let mut object = HashMap::with_capacity(values.len());
            for (key, value) in values {
                object.insert(key.clone(), json_to_any(value, depth + 1)?);
            }
            Ok(Any::Map(Arc::new(object)))
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

fn ensure_object_path(object: &mut JsonMap<String, JsonValue>, path: &[String]) -> UseResult<()> {
    let Some((key, rest)) = path.split_first() else {
        return Err(invalid_shared_spreadsheet(
            "A shared Spreadsheet cell contains an empty object path.",
        ));
    };
    let value = object
        .entry(key.clone())
        .or_insert_with(|| JsonValue::Object(JsonMap::new()));
    let child = value.as_object_mut().ok_or_else(|| {
        invalid_shared_spreadsheet(
            "A shared Spreadsheet cell contains overlapping object and value fields.",
        )
    })?;
    if rest.is_empty() {
        Ok(())
    } else {
        ensure_object_path(child, rest)
    }
}

fn set_value_at_path(
    object: &mut JsonMap<String, JsonValue>,
    path: &[String],
    value: JsonValue,
) -> UseResult<()> {
    let Some((key, rest)) = path.split_first() else {
        return Err(invalid_shared_spreadsheet(
            "A shared Spreadsheet cell contains an empty value path.",
        ));
    };
    if rest.is_empty() {
        if object.insert(key.clone(), value).is_some() {
            return Err(invalid_shared_spreadsheet(
                "A shared Spreadsheet cell contains duplicate object/value fields.",
            ));
        }
        return Ok(());
    }
    let child = object
        .get_mut(key)
        .and_then(JsonValue::as_object_mut)
        .ok_or_else(|| {
            invalid_shared_spreadsheet(
                "A shared Spreadsheet cell contains a value without an object parent.",
            )
        })?;
    set_value_at_path(child, rest, value)
}

fn invalid_json_key(key: &str) -> bool {
    key.is_empty() || matches!(key, "__proto__" | "constructor" | "prototype")
}
