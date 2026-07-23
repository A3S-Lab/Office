use a3s_use_core::UseResult;

use super::part::{
    dialect, ensure_namespace as ensure_part_namespace, relationship_part, relative_target,
};
use super::{editor_error, NativeOfficeComment, NativeOfficeCommentUpdate};
use crate::{
    DocumentKind, LosslessXmlPart, NativeOfficePackage, RelationshipSource, RelationshipTarget,
};

mod presentation;
mod spreadsheet;
mod word;

const WORD_COMMENTS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml";
const SPREADSHEET_COMMENTS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
const VML_CONTENT_TYPE: &str = "application/vnd.openxmlformats-officedocument.vmlDrawing";
const PRESENTATION_COMMENTS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.comments+xml";
const PRESENTATION_AUTHORS_CONTENT_TYPE: &str =
    "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml";

pub(super) fn add(
    package: &mut NativeOfficePackage,
    parent: &str,
    comment: &NativeOfficeComment,
) -> UseResult<String> {
    match package.kind() {
        DocumentKind::Word => word::add(package, parent, comment),
        DocumentKind::Spreadsheet => spreadsheet::add(package, parent, comment),
        DocumentKind::Presentation => presentation::add(package, parent, comment),
    }
}

pub(super) fn set(
    package: &mut NativeOfficePackage,
    path: &str,
    update: &NativeOfficeCommentUpdate,
) -> UseResult<String> {
    match package.kind() {
        DocumentKind::Word => word::set(package, path, update),
        DocumentKind::Spreadsheet => spreadsheet::set(package, path, update),
        DocumentKind::Presentation => presentation::set(package, path, update),
    }
}

pub(super) fn remove(package: &mut NativeOfficePackage, path: &str) -> UseResult<()> {
    match package.kind() {
        DocumentKind::Word => word::remove(package, path),
        DocumentKind::Spreadsheet => spreadsheet::remove(package, path),
        DocumentKind::Presentation => presentation::remove(package, path),
    }
}

pub(super) fn remove_presentation_slide_comments(
    package: &mut NativeOfficePackage,
    slide_part: &str,
) -> UseResult<()> {
    presentation::remove_for_slide(package, slide_part)
}

pub(super) fn remove_word_owner_comments(
    package: &mut NativeOfficePackage,
    path: &str,
) -> UseResult<()> {
    word::remove_owned(package, path)
}

pub(super) fn remove_spreadsheet_range_comments(
    package: &mut NativeOfficePackage,
    path: &str,
) -> UseResult<()> {
    spreadsheet::remove_for_range(package, path)
}

#[derive(Debug, Clone, PartialEq, Eq)]
struct RelatedPart {
    relationship_id: String,
    part_name: String,
}

fn related_part(
    package: &NativeOfficePackage,
    owner: &str,
    relationship_name: &str,
) -> UseResult<Option<RelatedPart>> {
    let source = RelationshipSource::Part {
        part_name: owner.to_string(),
    };
    let opc = package.opc_model()?;
    let Some(relationship) =
        opc.relationships()
            .relationships_from(&source)
            .iter()
            .find(|relationship| {
                relationship
                    .relationship_type
                    .ends_with(&format!("/{relationship_name}"))
            })
    else {
        return Ok(None);
    };
    let RelationshipTarget::Internal { part_name, .. } = &relationship.target else {
        return Err(comment_error(
            "use.office.comment_relationship_invalid",
            format!(
                "Office {relationship_name} relationship '{}' from '{owner}' must be internal.",
                relationship.id
            ),
        ));
    };
    Ok(Some(RelatedPart {
        relationship_id: relationship.id.clone(),
        part_name: part_name.clone(),
    }))
}

struct CreateRelatedPart<'a> {
    owner: &'a str,
    directory: &'a str,
    stem: &'a str,
    extension: &'a str,
    content_type: &'a str,
    relationship_name: &'a str,
    xml: String,
}

fn ensure_related_part(
    package: &mut NativeOfficePackage,
    request: CreateRelatedPart<'_>,
) -> UseResult<RelatedPart> {
    if let Some(existing) = related_part(package, request.owner, request.relationship_name)? {
        return Ok(existing);
    }
    let part_name = allocate_part(package, request.directory, request.stem, request.extension)?;
    LosslessXmlPart::parse(part_name.clone(), request.xml.as_bytes().to_vec())?;
    crate::opc_edit::add_content_type_override(package, &part_name, request.content_type)?;
    package.set_part(&part_name, request.xml.into_bytes())?;
    let relationship_id = crate::opc_edit::add_relationship(
        package,
        &relationship_part(request.owner),
        &dialect(package)?.relationship_type(request.relationship_name),
        &relative_target(request.owner, &part_name),
    )?;
    Ok(RelatedPart {
        relationship_id,
        part_name,
    })
}

fn allocate_part(
    package: &NativeOfficePackage,
    directory: &str,
    stem: &str,
    extension: &str,
) -> UseResult<String> {
    let direct = format!("{directory}/{stem}.{extension}");
    if !package.contains_part(&direct) {
        return Ok(direct);
    }
    (1..=package.limits().max_entries.saturating_add(1))
        .map(|number| format!("{directory}/{stem}{number}.{extension}"))
        .find(|candidate| !package.contains_part(candidate))
        .ok_or_else(|| {
            comment_error(
                "use.office.comment_part_name_exhausted",
                format!("No available '{stem}' Office comment part name remains."),
            )
        })
}

fn remove_related_part(
    package: &mut NativeOfficePackage,
    owner: &str,
    related: &RelatedPart,
) -> UseResult<()> {
    crate::opc_edit::remove_relationship(
        package,
        &relationship_part(owner),
        &related.relationship_id,
    )?;
    remove_content_type_override_if_present(package, &related.part_name)?;
    let related_relationships = relationship_part(&related.part_name);
    package.remove_part(&related_relationships)?;
    if !package.remove_part(&related.part_name)? {
        return Err(comment_error(
            "use.office.comment_part_missing",
            format!(
                "Office comment part '{}' does not exist.",
                related.part_name
            ),
        ));
    }
    Ok(())
}

fn remove_content_type_override_if_present(
    package: &mut NativeOfficePackage,
    part_name: &str,
) -> UseResult<()> {
    let has_override = package
        .opc_model()?
        .content_types()
        .overrides()
        .any(|(part, _)| part.eq_ignore_ascii_case(part_name.trim_start_matches('/')));
    if has_override {
        crate::opc_edit::remove_content_type_override(package, part_name)?;
    }
    Ok(())
}

fn derive_initials(author: &str) -> String {
    let words = author.split_whitespace().collect::<Vec<_>>();
    let initials = if words.len() <= 1 {
        author
            .chars()
            .filter(|character| character.is_alphanumeric())
            .take(2)
            .flat_map(char::to_uppercase)
            .collect::<String>()
    } else {
        words
            .into_iter()
            .filter_map(|word| word.chars().find(|character| character.is_alphanumeric()))
            .take(3)
            .flat_map(char::to_uppercase)
            .collect::<String>()
    };
    if initials.is_empty() {
        "?".to_string()
    } else {
        initials
    }
}

fn ensure_namespace(
    part: &LosslessXmlPart,
    preferred: &str,
    namespace: &str,
) -> UseResult<(Vec<u8>, String)> {
    ensure_part_namespace(
        part,
        preferred,
        namespace,
        "use.office.comment_namespace_exhausted",
    )
}

fn comment_error(code: &str, message: impl Into<String>) -> a3s_use_core::UseError {
    editor_error(code, message)
}
