mod create;
mod decide;
mod shared;

use a3s_use_core::UseResult;

use crate::collaboration::{
    collaboration_error, NativeOfficeCollaborationManifest, NativeOfficeCollaborationMutation,
};

pub(in crate::collaboration) fn validate_suggestion_mutation(
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    match mutation {
        NativeOfficeCollaborationMutation::DocumentSuggestionCreate { .. } => {
            create::validate_create_mutation(mutation)
        }
        NativeOfficeCollaborationMutation::DocumentSuggestionDecide { .. } => {
            decide::validate_decide_mutation(mutation)
        }
        _ => Err(collaboration_error(
            "office.collaboration.mutation_invalid",
            "The supplied mutation is not a Document suggestion mutation.",
        )),
    }
}

pub(in crate::collaboration) fn apply_suggestion_mutation(
    doc: &yrs::Doc,
    manifest: &NativeOfficeCollaborationManifest,
    mutation: &NativeOfficeCollaborationMutation,
) -> UseResult<()> {
    match mutation {
        NativeOfficeCollaborationMutation::DocumentSuggestionCreate { .. } => {
            create::apply_create_mutation(doc, manifest, mutation)
        }
        NativeOfficeCollaborationMutation::DocumentSuggestionDecide { .. } => {
            decide::apply_decide_mutation(doc, manifest, mutation)
        }
        _ => Err(collaboration_error(
            "office.collaboration.mutation_invalid",
            "The supplied mutation is not a Document suggestion mutation.",
        )),
    }
}
