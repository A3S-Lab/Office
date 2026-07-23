use a3s_use_core::UseResult;

use super::{editor_error, table};
use crate::editor::{NativeOfficeInsertPosition, NativeOfficeSwapResult};
use crate::xml_edit::IndexedXmlElement;
use crate::NativeOfficePackage;

mod object;
mod slide;

pub(in crate::editor) fn move_node(
    package: &mut NativeOfficePackage,
    path: &str,
    target_parent: Option<&str>,
    position: Option<&NativeOfficeInsertPosition>,
) -> UseResult<String> {
    if table::is_column_path(path) {
        table::move_column(package, path, target_parent, position)
    } else if parse_slide_path(path).is_some() {
        slide::move_slide(package, path, target_parent, position)
    } else {
        object::move_object(package, path, target_parent, position)
    }
}

pub(in crate::editor) fn copy_node(
    package: &mut NativeOfficePackage,
    path: &str,
    target_parent: Option<&str>,
    position: Option<&NativeOfficeInsertPosition>,
    name: Option<&str>,
) -> UseResult<String> {
    if name.is_some() {
        return Err(editor_error(
            "use.office.mutation_option_unsupported",
            "Presentation copy does not accept a worksheet name.",
        ));
    }
    if table::is_column_path(path) {
        table::copy_column(package, path, target_parent, position)
    } else if parse_slide_path(path).is_some() {
        slide::copy_slide(package, path, target_parent, position)
    } else {
        object::copy_object(package, path, target_parent, position)
    }
}

pub(in crate::editor) fn swap_nodes(
    package: &mut NativeOfficePackage,
    path: &str,
    with: &str,
) -> UseResult<NativeOfficeSwapResult> {
    match (table::is_column_path(path), table::is_column_path(with)) {
        (true, true) => return table::swap_columns(package, path, with),
        (true, false) | (false, true) => {
            return Err(editor_error(
                "use.office.mutation_type_unsupported",
                "Presentation swap requires two table columns, two slides, or two top-level objects on the same slide.",
            ));
        }
        (false, false) => {}
    }
    match (parse_slide_path(path), parse_slide_path(with)) {
        (Some(_), Some(_)) => slide::swap_slides(package, path, with),
        (None, None) => object::swap_objects(package, path, with),
        _ => Err(editor_error(
            "use.office.mutation_type_unsupported",
            "Presentation swap requires two table columns, two slides, or two top-level objects on the same slide.",
        )),
    }
}

pub(super) fn parse_slide_path(path: &str) -> Option<usize> {
    path.strip_prefix("/slide[")?
        .strip_suffix(']')?
        .parse::<usize>()
        .ok()
        .filter(|position| *position > 0)
}

pub(super) fn local_name(name: &str) -> &str {
    name.rsplit_once(':').map_or(name, |(_, local)| local)
}

pub(super) fn container_end(container: &IndexedXmlElement) -> usize {
    container
        .children
        .iter()
        .find(|child| child.local_name == "extLst")
        .map_or(container.content_range.end, |extension| {
            extension.full_range.start
        })
}
