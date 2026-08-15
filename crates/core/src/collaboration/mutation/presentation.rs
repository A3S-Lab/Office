use a3s_use_core::UseResult;

use super::super::{
    collaboration_error, NativeOfficeCollaborationManifest, NativeOfficeCollaborationMutation,
    NativeOfficeCollaborationPresentationContainerKind,
};

mod element;
mod json;
mod order;
mod state;

pub(super) fn validate_presentation_mutation(
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    match mutation {
        NativeOfficeCollaborationMutation::PresentationMoveElement { .. } => {
            order::validate_element_order_mutation(mutation)
        }
        _ => element::validate_element_mutation(mutation),
    }
}

pub(super) fn apply_presentation_mutation(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    match mutation {
        NativeOfficeCollaborationMutation::PresentationMoveElement { .. } => {
            order::apply_element_order_mutation(doc, manifest, mutation)
        }
        _ => element::apply_element_mutation(doc, manifest, mutation),
    }
}

fn invalid_presentation_mutation(message: impl Into<String>) -> a3s_use_core::UseError {
    collaboration_error("office.collaboration.mutation_invalid", message)
}

fn presentation_match_conflict(message: impl Into<String>) -> a3s_use_core::UseError {
    collaboration_error("office.collaboration.mutation_match_conflict", message).with_suggestion(
        "Read the latest collaborative Presentation scene element and z-order, then retry.",
    )
}

fn invalid_shared_presentation(message: impl Into<String>) -> a3s_use_core::UseError {
    collaboration_error("office.collaboration.content_invalid", message)
}
