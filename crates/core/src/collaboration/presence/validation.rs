use std::collections::BTreeSet;

use a3s_use_core::UseResult;

use super::super::{
    collaboration_error, validate_client_id, NativeOfficeCollaborationArtifactKind,
    NativeOfficeCollaborationManifest, NativeOfficeCollaborationPresenceLocation,
    NativeOfficeCollaborationPresenceProfile, NativeOfficeCollaborationPresenceState,
    NativeOfficeCollaborationSpreadsheetPresenceCell,
    NativeOfficeCollaborationSpreadsheetPresenceRange,
    MAX_NATIVE_OFFICE_COLLABORATION_PRESENCE_ELEMENT_IDS,
    MAX_NATIVE_OFFICE_COLLABORATION_PRESENCE_RANGES, NATIVE_OFFICE_COLLABORATION_PROTOCOL,
    NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION,
};

const MAX_JAVASCRIPT_SAFE_INTEGER: u64 = (1_u64 << 53) - 1;
const MAX_PRESENCE_IDENTIFIER_UNITS: usize = 256;
const MAX_PRESENCE_COLOR_UNITS: usize = 64;
const MAX_PRESENCE_AVATAR_URL_UNITS: usize = 2_048;

pub(super) fn validate_manifest(manifest: &NativeOfficeCollaborationManifest) -> UseResult<()> {
    if manifest.protocol != NATIVE_OFFICE_COLLABORATION_PROTOCOL
        || manifest.protocol_version != NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION
    {
        return Err(collaboration_error(
            "office.collaboration.presence_identity_mismatch",
            "The replica manifest does not use the supported Office collaboration protocol.",
        ));
    }
    validate_client_id(manifest.client_id)
}

pub(super) fn validate_profile(
    profile: &NativeOfficeCollaborationPresenceProfile,
) -> UseResult<()> {
    validate_bounded_string(&profile.name, MAX_PRESENCE_IDENTIFIER_UNITS, "actor name")?;
    if let Some(color) = &profile.color {
        validate_bounded_string(color, MAX_PRESENCE_COLOR_UNITS, "actor color")?;
    }
    if let Some(avatar_url) = &profile.avatar_url {
        validate_bounded_string(
            avatar_url,
            MAX_PRESENCE_AVATAR_URL_UNITS,
            "actor avatar URL",
        )?;
    }
    Ok(())
}

pub(super) fn validate_presence_state(
    state: &NativeOfficeCollaborationPresenceState,
    manifest: &NativeOfficeCollaborationManifest,
) -> UseResult<()> {
    if state.protocol != NATIVE_OFFICE_COLLABORATION_PROTOCOL
        || state.version != NATIVE_OFFICE_COLLABORATION_PROTOCOL_VERSION
        || state.artifact_id != manifest.artifact_id
        || state.artifact_kind != manifest.kind
        || state.namespace != manifest.namespace
    {
        return Err(collaboration_error(
            "office.collaboration.presence_identity_mismatch",
            "The Office Presence state belongs to another protocol, artifact, kind, or namespace.",
        ));
    }
    validate_bounded_string(
        &state.presence_id,
        MAX_PRESENCE_IDENTIFIER_UNITS,
        "presence ID",
    )?;
    validate_bounded_string(&state.actor.id, MAX_PRESENCE_IDENTIFIER_UNITS, "actor ID")?;
    validate_bounded_string(
        &state.actor.name,
        MAX_PRESENCE_IDENTIFIER_UNITS,
        "actor name",
    )?;
    if let Some(color) = &state.actor.color {
        validate_bounded_string(color, MAX_PRESENCE_COLOR_UNITS, "actor color")?;
    }
    if let Some(avatar_url) = &state.actor.avatar_url {
        validate_bounded_string(
            avatar_url,
            MAX_PRESENCE_AVATAR_URL_UNITS,
            "actor avatar URL",
        )?;
    }
    if let Some(location) = &state.location {
        validate_location(location, manifest.kind)?;
    }
    Ok(())
}

pub(super) fn validate_location(
    location: &NativeOfficeCollaborationPresenceLocation,
    artifact_kind: NativeOfficeCollaborationArtifactKind,
) -> UseResult<()> {
    if location.artifact_kind() != artifact_kind {
        return Err(invalid_presence(format!(
            "The Presence location kind '{}' does not match the '{}' artifact.",
            location.artifact_kind().as_str(),
            artifact_kind.as_str()
        )));
    }
    match location {
        NativeOfficeCollaborationPresenceLocation::Document { anchor, head }
        | NativeOfficeCollaborationPresenceLocation::Markdown { anchor, head, .. } => {
            validate_position(*anchor, "selection anchor")?;
            validate_position(*head, "selection head")?;
        }
        NativeOfficeCollaborationPresenceLocation::Spreadsheet {
            sheet_id,
            ranges,
            active_cell,
        } => {
            validate_bounded_string(sheet_id, MAX_PRESENCE_IDENTIFIER_UNITS, "sheet ID")?;
            if ranges.is_empty() || ranges.len() > MAX_NATIVE_OFFICE_COLLABORATION_PRESENCE_RANGES {
                return Err(invalid_presence(format!(
                    "Spreadsheet Presence requires 1 to {MAX_NATIVE_OFFICE_COLLABORATION_PRESENCE_RANGES} ranges."
                )));
            }
            for range in ranges {
                validate_range(range)?;
            }
            if let Some(active_cell) = active_cell {
                validate_cell(active_cell)?;
            }
        }
        NativeOfficeCollaborationPresenceLocation::Presentation {
            slide_id,
            element_ids,
        } => {
            validate_bounded_string(slide_id, MAX_PRESENCE_IDENTIFIER_UNITS, "slide ID")?;
            if element_ids.len() > MAX_NATIVE_OFFICE_COLLABORATION_PRESENCE_ELEMENT_IDS {
                return Err(invalid_presence(format!(
                    "Presentation Presence supports at most {MAX_NATIVE_OFFICE_COLLABORATION_PRESENCE_ELEMENT_IDS} element IDs."
                )));
            }
            let mut unique = BTreeSet::new();
            for element_id in element_ids {
                validate_bounded_string(element_id, MAX_PRESENCE_IDENTIFIER_UNITS, "element ID")?;
                if !unique.insert(element_id) {
                    return Err(invalid_presence(
                        "Presentation Presence element IDs must be unique.",
                    ));
                }
            }
        }
        NativeOfficeCollaborationPresenceLocation::Pdf {
            page_index,
            annotation_id,
        } => {
            validate_position(*page_index, "PDF page index")?;
            if let Some(annotation_id) = annotation_id {
                validate_bounded_string(
                    annotation_id,
                    MAX_PRESENCE_IDENTIFIER_UNITS,
                    "annotation ID",
                )?;
            }
        }
    }
    Ok(())
}

fn validate_range(range: &NativeOfficeCollaborationSpreadsheetPresenceRange) -> UseResult<()> {
    for (value, label) in [
        (range.start_row, "range start row"),
        (range.start_column, "range start column"),
        (range.end_row, "range end row"),
        (range.end_column, "range end column"),
    ] {
        validate_position(value, label)?;
    }
    if range.start_row > range.end_row || range.start_column > range.end_column {
        return Err(invalid_presence(
            "Spreadsheet Presence ranges must be normalized from start to end.",
        ));
    }
    Ok(())
}

fn validate_cell(cell: &NativeOfficeCollaborationSpreadsheetPresenceCell) -> UseResult<()> {
    validate_position(cell.row, "active-cell row")?;
    validate_position(cell.column, "active-cell column")
}

fn validate_position(value: u64, label: &str) -> UseResult<()> {
    if value <= MAX_JAVASCRIPT_SAFE_INTEGER {
        return Ok(());
    }
    Err(invalid_presence(format!(
        "The Presence {label} must fit in a JavaScript safe integer."
    )))
}

fn validate_bounded_string(value: &str, maximum_units: usize, label: &str) -> UseResult<()> {
    if !value.is_empty() && value.trim() == value && value.encode_utf16().count() <= maximum_units {
        return Ok(());
    }
    Err(invalid_presence(format!(
        "The Presence {label} must contain 1 to {maximum_units} non-padded UTF-16 units."
    )))
}

fn invalid_presence(message: impl Into<String>) -> a3s_use_core::UseError {
    collaboration_error("office.collaboration.presence_invalid", message)
}
