use a3s_use_core::{UseError, UseResult};

use crate::discovery::office_error;

#[derive(Debug, Clone, PartialEq, Eq)]
pub(super) struct PathSegment {
    pub(super) name: String,
    pub(super) position: Option<usize>,
}

pub(super) fn parse_segments(path: &str) -> UseResult<Vec<PathSegment>> {
    validate_mutation_path(path)?;
    path.trim_start_matches('/')
        .split('/')
        .map(|segment| {
            if let Some((name, position)) = segment.split_once('[') {
                let position = position.strip_suffix(']').ok_or_else(|| {
                    editor_error(
                        "use.office.path_invalid",
                        format!("Office path segment '{segment}' is missing ']'."),
                    )
                })?;
                let position = position.parse::<usize>().map_err(|_| {
                    editor_error(
                        "use.office.path_invalid",
                        format!("Office mutation path segment '{segment}' is not numeric."),
                    )
                })?;
                if name.is_empty() || position == 0 {
                    return Err(editor_error(
                        "use.office.path_invalid",
                        "Office mutation paths use non-empty, one-based segments.",
                    ));
                }
                Ok(PathSegment {
                    name: name.to_ascii_lowercase(),
                    position: Some(position),
                })
            } else if segment.is_empty() {
                Err(editor_error(
                    "use.office.path_invalid",
                    "Office mutation paths cannot contain empty segments.",
                ))
            } else {
                Ok(PathSegment {
                    name: segment.to_ascii_lowercase(),
                    position: None,
                })
            }
        })
        .collect()
}

pub(super) fn validate_mutation_path(path: &str) -> UseResult<()> {
    if path.is_empty()
        || path == "/"
        || !path.starts_with('/')
        || path.len() > 4_096
        || path.contains('\\')
        || path.chars().any(char::is_control)
        || path.split('/').any(|segment| matches!(segment, "." | ".."))
    {
        return Err(editor_error(
            "use.office.path_invalid",
            "Office mutation path must be absolute, bounded, and traversal-free.",
        ));
    }
    Ok(())
}

pub(super) fn prefix(qualified_name: &str) -> Option<&str> {
    qualified_name.rsplit_once(':').map(|(prefix, _)| prefix)
}

pub(super) fn qualified(prefix: Option<&str>, local_name: &str) -> String {
    prefix.map_or_else(
        || local_name.to_string(),
        |prefix| format!("{prefix}:{local_name}"),
    )
}

pub(super) fn preserve_space_attribute(text: &str) -> &'static str {
    if text.starts_with(char::is_whitespace) || text.ends_with(char::is_whitespace) {
        " xml:space=\"preserve\""
    } else {
        ""
    }
}

pub(super) fn escape_attribute(value: &str) -> String {
    quick_xml::escape::escape(value).into_owned()
}

pub(super) fn node_not_found(path: &str) -> UseError {
    editor_error(
        "use.office.node_not_found",
        format!("Office semantic path '{path}' does not exist."),
    )
    .with_detail("path", path)
}

pub(super) fn editor_error(code: &str, message: impl Into<String>) -> UseError {
    office_error(code, message)
}
