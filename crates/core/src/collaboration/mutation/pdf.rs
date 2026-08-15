use std::collections::{HashMap, HashSet};

use a3s_use_core::UseResult;
use yrs::{Any, Array, ArrayRef, Map, MapRef, Out, Transact};

use super::super::{
    collaboration_error, NativeOfficeCollaborationManifest, NativeOfficeCollaborationMutation,
};

mod annotation;
mod records;
mod review;

use annotation::{apply_pdf_annotation_mutation, validate_pdf_annotation_mutation};
use review::{apply_pdf_review_mutation, validate_pdf_review_mutation};

const MAX_PDF_IDENTIFIER_UTF16: usize = 512;
pub(super) const MAX_PDF_RECORDS: u32 = 1_000_000;

pub(super) fn validate_pdf_mutation(mutation: &NativeOfficeCollaborationMutation) -> UseResult<()> {
    match mutation {
        NativeOfficeCollaborationMutation::PdfCreateAnnotation { .. }
        | NativeOfficeCollaborationMutation::PdfUpdateAnnotation { .. }
        | NativeOfficeCollaborationMutation::PdfDeleteAnnotation { .. } => {
            validate_pdf_annotation_mutation(mutation)
        }
        NativeOfficeCollaborationMutation::PdfSetFormValue { field_id, .. } => {
            validate_pdf_identifier(field_id, "fieldId", "PDF form field")
        }
        NativeOfficeCollaborationMutation::PdfProposeRedaction { .. }
        | NativeOfficeCollaborationMutation::PdfProposePageRotation { .. }
        | NativeOfficeCollaborationMutation::PdfProposePageDeletion { .. }
        | NativeOfficeCollaborationMutation::PdfProposePageReorder { .. }
        | NativeOfficeCollaborationMutation::PdfDecideReview { .. } => {
            validate_pdf_review_mutation(mutation)
        }
        _ => Err(collaboration_error(
            "office.collaboration.mutation_invalid",
            "The supplied mutation is not a PDF mutation.",
        )),
    }
}

pub(super) fn apply_pdf_mutation(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    match mutation {
        NativeOfficeCollaborationMutation::PdfCreateAnnotation { .. }
        | NativeOfficeCollaborationMutation::PdfUpdateAnnotation { .. }
        | NativeOfficeCollaborationMutation::PdfDeleteAnnotation { .. } => {
            apply_pdf_annotation_mutation(doc, manifest, mutation)
        }
        NativeOfficeCollaborationMutation::PdfSetFormValue { field_id, value } => {
            set_pdf_form_value(doc, manifest, field_id, value)
        }
        NativeOfficeCollaborationMutation::PdfProposeRedaction { .. }
        | NativeOfficeCollaborationMutation::PdfProposePageRotation { .. }
        | NativeOfficeCollaborationMutation::PdfProposePageDeletion { .. }
        | NativeOfficeCollaborationMutation::PdfProposePageReorder { .. }
        | NativeOfficeCollaborationMutation::PdfDecideReview { .. } => {
            apply_pdf_review_mutation(doc, manifest, mutation)
        }
        _ => Err(collaboration_error(
            "office.collaboration.mutation_invalid",
            "The supplied mutation is not a PDF mutation.",
        )),
    }
}

fn set_pdf_form_value(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    field_id: &str,
    value: &str,
) -> UseResult<()> {
    let roots = PdfRecordCollectionRoots::new(doc, manifest, "form-values");
    let current = read_pdf_form_values(doc, &roots)?;
    if current
        .get(field_id)
        .is_some_and(|current| current == value)
    {
        return Ok(());
    }

    let id_key = encoded_record_field_key(field_id, &["id"])?;
    let value_key = encoded_record_field_key(field_id, &["value"])?;
    let mut transaction = doc.transact_mut();
    if !current.contains_key(field_id) {
        roots.presence.insert(&mut transaction, field_id, true);
        roots.fields.insert(&mut transaction, id_key, field_id);
        roots.order.push_back(&mut transaction, field_id);
    }
    roots.fields.insert(&mut transaction, value_key, value);
    drop(transaction);

    read_pdf_form_values(doc, &roots)?;
    Ok(())
}

pub(super) struct PdfRecordCollectionRoots {
    pub(super) presence: MapRef,
    pub(super) fields: MapRef,
    pub(super) order: ArrayRef,
}

impl PdfRecordCollectionRoots {
    pub(super) fn new(
        doc: &yrs::Doc,
        manifest: &NativeOfficeCollaborationManifest,
        collection: &str,
    ) -> Self {
        let root = |suffix: &str| format!("{}.pdf.{collection}.{suffix}", manifest.namespace);
        Self {
            presence: doc.get_or_insert_map(root("presence")),
            fields: doc.get_or_insert_map(root("fields")),
            order: doc.get_or_insert_array(root("order")),
        }
    }
}

fn read_pdf_form_values(
    doc: &yrs::Doc,
    roots: &PdfRecordCollectionRoots,
) -> UseResult<HashMap<String, String>> {
    let transaction = doc.transact();
    if roots.presence.len(&transaction) > MAX_PDF_RECORDS {
        return Err(invalid_shared_pdf(
            "The shared PDF contains too many form values.",
        ));
    }

    let mut ordered_ids = HashSet::new();
    for index in 0..roots.order.len(&transaction) {
        let id = match roots.order.get(&transaction, index) {
            Some(Out::Any(Any::String(value))) => value.to_string(),
            _ => {
                return Err(invalid_shared_pdf(
                    "The shared PDF form-value order contains a non-string identity.",
                ))
            }
        };
        validate_shared_pdf_identifier(&id, "form field")?;
        ordered_ids.insert(id);
    }

    let mut present_ids = HashSet::new();
    for (id, presence) in roots.presence.iter(&transaction) {
        validate_shared_pdf_identifier(id, "form field")?;
        if !matches!(presence, Out::Any(Any::Bool(true))) || !ordered_ids.contains(id) {
            return Err(invalid_shared_pdf(
                "The shared PDF form-value presence and order roots disagree.",
            ));
        }
        present_ids.insert(id.to_owned());
    }
    if ordered_ids.len() != present_ids.len() || ordered_ids != present_ids {
        return Err(invalid_shared_pdf(
            "The shared PDF form-value presence and order roots disagree.",
        ));
    }

    let mut ids = HashMap::new();
    let mut values = HashMap::new();
    for (encoded, value) in roots.fields.iter(&transaction) {
        let (id, path) = decoded_record_field_key(encoded, "form value")?;
        if !present_ids.contains(&id) {
            return Err(invalid_shared_pdf(
                "The shared PDF contains an orphan form-value field.",
            ));
        }
        if path.len() != 1 {
            return Err(invalid_shared_pdf(
                "The shared PDF form-value record contains a nested field.",
            ));
        }
        let value = match value {
            Out::Any(Any::String(value)) => value.to_string(),
            _ => {
                return Err(invalid_shared_pdf(
                    "The shared PDF form-value record contains a non-string field.",
                ))
            }
        };
        let previous = if path[0] == "id" {
            if value != id {
                return Err(invalid_shared_pdf(
                    "The shared PDF form-value identity field does not match its record.",
                ));
            }
            ids.insert(id, value)
        } else if path[0] == "value" {
            values.insert(id, value)
        } else {
            return Err(invalid_shared_pdf(
                "The shared PDF form-value record contains an unsupported field.",
            ));
        };
        if previous.is_some() {
            return Err(invalid_shared_pdf(
                "The shared PDF form-value record contains duplicate fields.",
            ));
        }
    }

    if ids.len() != present_ids.len()
        || values.len() != present_ids.len()
        || ids.keys().any(|id| !present_ids.contains(id))
        || values.keys().any(|id| !present_ids.contains(id))
    {
        return Err(invalid_shared_pdf(
            "The shared PDF form-value records are incomplete.",
        ));
    }
    Ok(values)
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub(super) enum PdfRecordFieldKind {
    Object,
    Value,
}

impl PdfRecordFieldKind {
    const fn as_str(self) -> &'static str {
        match self {
            Self::Object => "object",
            Self::Value => "value",
        }
    }
}

pub(super) struct PdfRecordFieldIdentity {
    pub(super) record_id: String,
    pub(super) kind: PdfRecordFieldKind,
    pub(super) path: Vec<String>,
}

pub(super) fn encoded_record_field_key(record_id: &str, path: &[&str]) -> UseResult<String> {
    let path = path
        .iter()
        .map(|part| (*part).to_owned())
        .collect::<Vec<_>>();
    encoded_record_entry_key(record_id, PdfRecordFieldKind::Value, &path)
}

pub(super) fn encoded_record_entry_key(
    record_id: &str,
    kind: PdfRecordFieldKind,
    path: &[String],
) -> UseResult<String> {
    if path.is_empty() || path.iter().any(|part| invalid_json_path_part(part)) {
        return Err(collaboration_error(
            "office.collaboration.mutation_invalid",
            "A PDF record field path contains an empty or unsafe key.",
        ));
    }
    let mut identity = Vec::with_capacity(path.len() + 1);
    identity.push(kind.as_str().to_owned());
    identity.extend_from_slice(path);
    let field = serde_json::to_string(&identity).map_err(|error| {
        collaboration_error(
            "office.collaboration.mutation_invalid",
            format!("Failed to encode the PDF record field identity: {error}"),
        )
    })?;
    serde_json::to_string(&[record_id, field.as_str()]).map_err(|error| {
        collaboration_error(
            "office.collaboration.mutation_invalid",
            format!("Failed to encode the PDF record identity: {error}"),
        )
    })
}

pub(super) fn decoded_record_field_key(
    encoded: &str,
    label: &str,
) -> UseResult<(String, Vec<String>)> {
    let identity = decoded_record_field_identity(encoded, label)?;
    if identity.kind != PdfRecordFieldKind::Value {
        return Err(invalid_shared_pdf(format!(
            "The shared PDF {label} contains an unexpected object marker."
        )));
    }
    Ok((identity.record_id, identity.path))
}

pub(super) fn decoded_record_field_identity(
    encoded: &str,
    label: &str,
) -> UseResult<PdfRecordFieldIdentity> {
    let identity = serde_json::from_str::<Vec<String>>(encoded).map_err(|_| {
        invalid_shared_pdf(format!(
            "The shared PDF {label} field identity is not valid JSON."
        ))
    })?;
    if identity.len() != 2
        || serde_json::to_string(&identity).ok().as_deref() != Some(encoded)
        || identity[0].is_empty()
    {
        return Err(invalid_shared_pdf(format!(
            "The shared PDF {label} field identity is invalid."
        )));
    }
    validate_shared_pdf_identifier(&identity[0], label)?;
    let field = serde_json::from_str::<Vec<String>>(&identity[1]).map_err(|_| {
        invalid_shared_pdf(format!(
            "The shared PDF {label} property identity is not valid JSON."
        ))
    })?;
    let kind = match field.first().map(String::as_str) {
        Some("object") => PdfRecordFieldKind::Object,
        Some("value") => PdfRecordFieldKind::Value,
        _ => {
            return Err(invalid_shared_pdf(format!(
                "The shared PDF {label} property identity has an unsupported kind."
            )))
        }
    };
    if field.len() < 2
        || field[1..].iter().any(|part| invalid_json_path_part(part))
        || serde_json::to_string(&field).ok().as_deref() != Some(identity[1].as_str())
    {
        return Err(invalid_shared_pdf(format!(
            "The shared PDF {label} property identity is invalid."
        )));
    }
    Ok(PdfRecordFieldIdentity {
        record_id: identity[0].clone(),
        kind,
        path: field[1..].to_vec(),
    })
}

pub(super) fn validate_pdf_identifier(value: &str, field: &str, label: &str) -> UseResult<()> {
    if valid_pdf_identifier(value) {
        return Ok(());
    }
    Err(collaboration_error(
        "office.collaboration.mutation_invalid",
        format!(
            "A {label} ID must contain 1 to 512 UTF-16 code units without leading or trailing whitespace."
        ),
    )
    .with_detail(field, value.to_owned()))
}

pub(super) fn validate_shared_pdf_identifier(value: &str, label: &str) -> UseResult<()> {
    if valid_pdf_identifier(value) {
        return Ok(());
    }
    Err(invalid_shared_pdf(format!(
        "The shared PDF contains an invalid {label} identity."
    )))
}

fn valid_pdf_identifier(value: &str) -> bool {
    let utf16_len = value.encode_utf16().count();
    let has_trimmed_edge = value
        .chars()
        .next()
        .is_some_and(is_ecmascript_trim_character)
        || value
            .chars()
            .next_back()
            .is_some_and(is_ecmascript_trim_character);
    (1..=MAX_PDF_IDENTIFIER_UTF16).contains(&utf16_len) && !has_trimmed_edge
}

fn invalid_json_path_part(part: &str) -> bool {
    part.is_empty() || matches!(part, "__proto__" | "constructor" | "prototype")
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

pub(super) fn invalid_shared_pdf(message: impl Into<String>) -> a3s_use_core::UseError {
    collaboration_error("office.collaboration.content_invalid", message)
}
