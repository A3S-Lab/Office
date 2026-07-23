use a3s_use_core::UseResult;

use super::{
    editor_error, mark_workbook_for_recalculation, node_not_found, update_dimension,
    validate_mutation_path,
};
use crate::semantic::NativeOfficeDocument;
use crate::xml_edit::index_xml;
use crate::{DocumentKind, NativeOfficePackage, NativeSpreadsheetSort};

mod metadata;
mod rows;
mod scope;
mod state;
mod values;

pub(super) fn is_path(path: &str) -> bool {
    state::is_path(path)
}

pub(super) fn sort(
    package: &mut NativeOfficePackage,
    path: &str,
    request: &NativeSpreadsheetSort,
) -> UseResult<String> {
    require_spreadsheet(package)?;
    validate_mutation_path(path)?;
    let snapshot = NativeOfficeDocument::from_package(package.clone())?;
    let resolved = scope::resolve(&snapshot, path)?;
    let keys = request.validate(resolved.range)?;
    let sheet = snapshot.get(&resolved.sheet_path, 2)?;
    let data_range = scope::validate_geometry(package, &sheet, &resolved, request.header)?;
    scope::validate_existing_state(&sheet)?;
    scope::validate_formula_boundary(package)?;

    let old_to_new = values::permutation(&sheet, data_range, &keys, request.case_sensitive)?;
    let worksheet = package.xml_part(&resolved.part_name)?;
    let root = index_xml(&worksheet)?;
    let sheet_data = root.child("sheetData", 1).ok_or_else(|| {
        sort_error(
            "use.office.spreadsheet_sheet_data_missing",
            format!(
                "Worksheet '{}' has no sheetData element.",
                resolved.sheet_path
            ),
        )
    })?;
    let edited = rows::rebuild(&worksheet, sheet_data, data_range, &old_to_new)?;
    let edited = update_dimension(&resolved.part_name, edited)?;
    package.set_part(&resolved.part_name, edited)?;

    let worksheet = package.xml_part(&resolved.part_name)?;
    let edited = metadata::rewrite_worksheet(&worksheet, data_range, &old_to_new)?;
    package.set_part(&resolved.part_name, edited)?;
    metadata::rewrite_related(package, &resolved.part_name, data_range, &old_to_new)?;
    metadata::clear_chart_caches(package)?;
    state::replace(
        package,
        &resolved.part_name,
        resolved.range,
        data_range,
        &keys,
        request,
    )?;
    mark_workbook_for_recalculation(package)?;
    Ok(format!("{}/sort", resolved.sheet_path))
}

pub(super) fn remove(package: &mut NativeOfficePackage, path: &str) -> UseResult<()> {
    state::remove(package, path)
}

fn require_spreadsheet(package: &NativeOfficePackage) -> UseResult<()> {
    if package.kind() == DocumentKind::Spreadsheet {
        Ok(())
    } else {
        Err(sort_error(
            "use.office.mutation_type_unsupported",
            "Spreadsheet sorting is available only for Spreadsheet documents.",
        ))
    }
}

fn sort_error(code: &str, message: impl Into<String>) -> a3s_use_core::UseError {
    editor_error(code, message)
}
