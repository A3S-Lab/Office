mod planning;
mod write;

use std::collections::BTreeMap;

use a3s_use_core::{UseError, UseResult};

use crate::semantic::{DocumentNode, NativeOfficeDocument};
use crate::spreadsheet_reference::{CellRange, CellReference};
use crate::xml_edit::{index_xml, IndexedXmlElement};
use crate::{
    NativeOfficePackage, SpreadsheetFormulaCalculation, SpreadsheetFormulaValue,
    MAX_SPREADSHEET_FORMULA_CELLS, MAX_SPREADSHEET_FORMULA_SPILL_CELLS,
};

use super::{editor_error, update_dimension};
use planning::{plan_writes, worksheet_cells};
use write::{apply_cell_writes, mark_workbook_calculated};

const MAX_CALCULATION_WRITES: usize =
    MAX_SPREADSHEET_FORMULA_CELLS + MAX_SPREADSHEET_FORMULA_SPILL_CELLS;

#[derive(Debug, Clone)]
enum CellWrite {
    Clear,
    Cached(SpreadsheetFormulaValue),
    Formula {
        expression: String,
        value: SpreadsheetFormulaValue,
        spill_range: Option<String>,
    },
}

pub(super) fn prepare_for_value_write(
    part: &crate::LosslessXmlPart,
    sheet_data: &IndexedXmlElement,
    sheet: &DocumentNode,
    target: CellRange,
) -> UseResult<Vec<u8>> {
    prepare_spill_edit(part, sheet_data, sheet, target, false)
}

pub(super) fn prepare_for_remove(
    part: &crate::LosslessXmlPart,
    sheet_data: &IndexedXmlElement,
    sheet: &DocumentNode,
    target: CellRange,
) -> UseResult<Vec<u8>> {
    prepare_spill_edit(part, sheet_data, sheet, target, true)
}

fn prepare_spill_edit(
    part: &crate::LosslessXmlPart,
    sheet_data: &IndexedXmlElement,
    sheet: &DocumentNode,
    target: CellRange,
    remove: bool,
) -> UseResult<Vec<u8>> {
    let cells = worksheet_cells(sheet)?;
    let mut writes = BTreeMap::new();
    for (anchor, cell) in &cells {
        let Some(reference) = cell.format.get("formulaRef") else {
            continue;
        };
        let spill = CellRange::parse(reference)?;
        if !target.intersects(spill) {
            continue;
        }
        if !target.contains(*anchor)
            || (!remove && !intersection_is_anchor_only(target, spill, *anchor))
        {
            return Err(calculation_storage_error(
                "use.office.spreadsheet_formula_spill_cell_read_only",
                format!(
                    "Cell range '{}' intersects spill '{}' outside formula anchor '{}'.",
                    target.a1(),
                    spill.a1(),
                    anchor.a1()
                ),
            )
            .with_suggestion(
                "Edit or remove the spill formula anchor; spilled result cells are read-only.",
            ));
        }
        let spill_cells = spill.cell_count()?;
        if spill_cells > MAX_SPREADSHEET_FORMULA_SPILL_CELLS {
            return Err(calculation_write_limit().with_detail("cells", spill_cells));
        }
        for row in spill.start.row..=spill.end.row {
            for column in spill.start.column..=spill.end.column {
                let reference = CellReference { column, row };
                if reference != *anchor {
                    insert_clear_write(&mut writes, reference)?;
                }
            }
        }
    }
    if writes.is_empty() {
        Ok(part.raw().to_vec())
    } else {
        apply_cell_writes(part, sheet_data, &writes)
    }
}

fn intersection_is_anchor_only(left: CellRange, right: CellRange, anchor: CellReference) -> bool {
    let start_column = left.start.column.max(right.start.column);
    let start_row = left.start.row.max(right.start.row);
    let end_column = left.end.column.min(right.end.column);
    let end_row = left.end.row.min(right.end.row);
    start_column == anchor.column
        && end_column == anchor.column
        && start_row == anchor.row
        && end_row == anchor.row
}

pub(super) fn recalculate(
    package: &mut NativeOfficePackage,
) -> UseResult<SpreadsheetFormulaCalculation> {
    let document = NativeOfficeDocument::from_package(package.clone())?;
    let calculation = document.calculate_spreadsheet_formulas()?;
    let plans = plan_writes(&document, &calculation)?;
    for (part_name, writes) in plans {
        if writes.is_empty() {
            continue;
        }
        let part = package.xml_part(&part_name)?;
        let index = index_xml(&part)?;
        let sheet_data = index.descendant("sheetData").ok_or_else(|| {
            calculation_storage_error(
                "use.office.spreadsheet_formula_storage_invalid",
                format!("Worksheet part '{part_name}' has no sheetData element."),
            )
        })?;
        let edited = apply_cell_writes(&part, sheet_data, &writes)?;
        let edited = update_dimension(&part_name, edited)?;
        package.set_part(&part_name, edited)?;
    }
    mark_workbook_calculated(package)?;
    Ok(calculation)
}

fn calculation_write_limit() -> UseError {
    calculation_storage_error(
        "use.office.spreadsheet_formula_write_limit",
        format!("Native formula recalculation writes at most {MAX_CALCULATION_WRITES} cells."),
    )
}

fn insert_clear_write(
    writes: &mut BTreeMap<CellReference, CellWrite>,
    reference: CellReference,
) -> UseResult<()> {
    if writes.contains_key(&reference) {
        return Ok(());
    }
    let cells = writes
        .len()
        .checked_add(1)
        .ok_or_else(calculation_write_limit)?;
    if cells > MAX_CALCULATION_WRITES {
        return Err(calculation_write_limit().with_detail("cells", cells));
    }
    writes.insert(reference, CellWrite::Clear);
    Ok(())
}

fn calculation_storage_error(code: &str, message: impl Into<String>) -> UseError {
    editor_error(code, message)
}
