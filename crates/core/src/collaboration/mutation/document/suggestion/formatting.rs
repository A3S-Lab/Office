use std::collections::{HashMap, HashSet};

use a3s_use_core::UseResult;
use serde_json::Value as JsonValue;
use yrs::Any;

use crate::collaboration::collaboration_error;

const MAX_FORMATTING_SNAPSHOT_UTF16: usize = 4_096;
const MAX_FORMATTING_ATTRIBUTE_UTF16: usize = 512;
const FORMATTING_MARKS: &[&str] = &[
    "bold",
    "italic",
    "underline",
    "strike",
    "subscript",
    "superscript",
    "textStyle",
    "highlight",
];

pub(super) fn validate_document_formatting_change(value: &Any) -> UseResult<bool> {
    let Any::Map(fields) = value else {
        return Ok(false);
    };
    if string_field(fields, "kind")?.as_deref() != Some("formatting") {
        return Ok(false);
    }
    let actor_id = optional_identifier(fields, "actorId", "formatting-change actor ID")?;
    let expected_fields = if actor_id.is_some() { 6 } else { 5 };
    if fields.len() != expected_fields
        || fields.keys().any(|key| {
            !matches!(
                key.as_str(),
                "actorId" | "author" | "before" | "date" | "id" | "kind"
            )
        })
    {
        return Err(invalid_formatting(
            "A shared Document formatting-change mark contains unsupported fields.",
        ));
    }
    required_identifier(fields, "id", "formatting-change ID")?;
    required_identifier(fields, "author", "formatting-change author")?;
    let date = required_string(fields, "date", "formatting-change date")?;
    if date != date.trim() || date.encode_utf16().count() > 256 {
        return Err(invalid_formatting(
            "A shared Document formatting-change date is invalid.",
        ));
    }
    let snapshot = required_string(fields, "before", "formatting snapshot")?;
    validate_formatting_snapshot(&snapshot)?;
    Ok(true)
}

fn validate_formatting_snapshot(source: &str) -> UseResult<()> {
    if source.encode_utf16().count() > MAX_FORMATTING_SNAPSHOT_UTF16 {
        return Err(invalid_formatting(
            "A shared Document formatting snapshot is too large.",
        ));
    }
    let JsonValue::Array(marks) = serde_json::from_str(source).map_err(|_| {
        invalid_formatting("A shared Document formatting snapshot is not valid JSON.")
    })?
    else {
        return Err(invalid_formatting(
            "A shared Document formatting snapshot is not an array.",
        ));
    };
    if marks.len() > FORMATTING_MARKS.len() {
        return Err(invalid_formatting(
            "A shared Document formatting snapshot contains too many marks.",
        ));
    }
    let mut seen = HashSet::new();
    for mark in marks {
        let object = mark.as_object().ok_or_else(|| {
            invalid_formatting("A shared Document formatting mark is not an object.")
        })?;
        if object.len() > 2
            || object
                .keys()
                .any(|key| !matches!(key.as_str(), "type" | "attrs"))
        {
            return Err(invalid_formatting(
                "A shared Document formatting mark contains unsupported fields.",
            ));
        }
        let name = object
            .get("type")
            .and_then(JsonValue::as_str)
            .filter(|name| FORMATTING_MARKS.contains(name))
            .ok_or_else(|| {
                invalid_formatting("A shared Document formatting mark type is invalid.")
            })?;
        if !seen.insert(name.to_owned()) {
            return Err(invalid_formatting(
                "A shared Document formatting snapshot repeats one mark type.",
            ));
        }
        let Some(attributes) = object.get("attrs") else {
            continue;
        };
        let attributes = attributes.as_object().ok_or_else(|| {
            invalid_formatting("A shared Document formatting mark attributes value is invalid.")
        })?;
        for (key, value) in attributes {
            if !allowed_attribute(name, key) || !valid_attribute_value(value) {
                return Err(invalid_formatting(
                    "A shared Document formatting mark attribute is unsupported.",
                ));
            }
        }
    }
    Ok(())
}

fn allowed_attribute(mark: &str, attribute: &str) -> bool {
    match mark {
        "textStyle" => matches!(
            attribute,
            "color"
                | "fontFamily"
                | "fontSize"
                | "themeColor"
                | "wordLineHeightFactor"
                | "wordSnapToGrid"
        ),
        "highlight" => matches!(attribute, "color" | "themeFill"),
        _ => false,
    }
}

fn valid_attribute_value(value: &JsonValue) -> bool {
    match value {
        JsonValue::String(value) => {
            !value.is_empty() && value.encode_utf16().count() <= MAX_FORMATTING_ATTRIBUTE_UTF16
        }
        JsonValue::Bool(_) => true,
        JsonValue::Number(value) => value.as_f64().is_some_and(f64::is_finite),
        _ => false,
    }
}

fn optional_identifier(
    fields: &HashMap<String, Any>,
    key: &str,
    label: &str,
) -> UseResult<Option<String>> {
    match string_field(fields, key)? {
        Some(value) => {
            validate_identifier(&value, label)?;
            Ok(Some(value))
        }
        None => Ok(None),
    }
}

fn required_identifier(fields: &HashMap<String, Any>, key: &str, label: &str) -> UseResult<String> {
    let value = required_string(fields, key, label)?;
    validate_identifier(&value, label)?;
    Ok(value)
}

fn required_string(fields: &HashMap<String, Any>, key: &str, label: &str) -> UseResult<String> {
    string_field(fields, key)?
        .ok_or_else(|| invalid_formatting(format!("A shared Document {label} is missing.")))
}

fn string_field(fields: &HashMap<String, Any>, key: &str) -> UseResult<Option<String>> {
    match fields.get(key) {
        Some(Any::String(value)) => Ok(Some(value.to_string())),
        Some(_) => Err(invalid_formatting(format!(
            "A shared Document formatting-change field '{key}' is not a string."
        ))),
        None => Ok(None),
    }
}

fn validate_identifier(value: &str, label: &str) -> UseResult<()> {
    if !value.is_empty() && value == value.trim() && value.encode_utf16().count() <= 256 {
        return Ok(());
    }
    Err(invalid_formatting(format!(
        "A shared Document {label} must contain 1 to 256 non-padded characters."
    )))
}

fn invalid_formatting(message: impl Into<String>) -> a3s_use_core::UseError {
    collaboration_error("office.collaboration.permission_denied", message)
}
