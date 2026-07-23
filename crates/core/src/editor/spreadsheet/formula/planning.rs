use std::collections::BTreeMap;

use a3s_use_core::UseResult;

use crate::semantic::{DocumentNode, NativeOfficeDocument, OfficeNodeType};
use crate::spreadsheet_reference::{CellRange, CellReference};
use crate::{
    SpreadsheetFormulaCalculatedCell, SpreadsheetFormulaCalculation, SpreadsheetFormulaValue,
    MAX_SPREADSHEET_FORMULA_SPILL_CELLS,
};

use super::{
    calculation_storage_error, calculation_write_limit, insert_clear_write, CellWrite,
    MAX_CALCULATION_WRITES,
};

pub(super) fn plan_writes(
    document: &NativeOfficeDocument,
    calculation: &SpreadsheetFormulaCalculation,
) -> UseResult<BTreeMap<String, BTreeMap<CellReference, CellWrite>>> {
    let mut plans = BTreeMap::new();
    let mut planned_write_count = 0_usize;
    for sheet in document
        .root()
        .children
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::Worksheet)
    {
        let sheet_name = sheet.path.strip_prefix('/').ok_or_else(|| {
            calculation_storage_error(
                "use.office.spreadsheet_formula_storage_invalid",
                format!("Worksheet path '{}' is invalid.", sheet.path),
            )
        })?;
        let part_name = sheet.format.get("part").cloned().ok_or_else(|| {
            calculation_storage_error(
                "use.office.spreadsheet_formula_storage_invalid",
                format!("Worksheet '{}' has no source part.", sheet.path),
            )
        })?;
        let cells = worksheet_cells(sheet)?;
        let mut writes = BTreeMap::new();
        plan_old_spill_cleanup(&cells, &mut writes)?;
        for calculated in calculation
            .cells
            .iter()
            .filter(|cell| cell.cell.sheet.eq_ignore_ascii_case(sheet_name))
        {
            plan_calculated_cell(calculated, &cells, &mut writes)?;
        }
        validate_planned_writes(&cells, &writes)?;
        if writes.len() > MAX_CALCULATION_WRITES {
            return Err(calculation_storage_error(
                "use.office.spreadsheet_formula_write_limit",
                format!(
                    "Native formula recalculation writes at most {MAX_CALCULATION_WRITES} cells."
                ),
            )
            .with_detail("cells", writes.len()));
        }
        planned_write_count = planned_write_count
            .checked_add(writes.len())
            .ok_or_else(calculation_write_limit)?;
        if planned_write_count > MAX_CALCULATION_WRITES {
            return Err(calculation_write_limit().with_detail("cells", planned_write_count));
        }
        plans.insert(part_name, writes);
    }
    let planned_formulas = plans
        .values()
        .flat_map(BTreeMap::values)
        .filter(|write| matches!(write, CellWrite::Formula { .. }))
        .count();
    if planned_formulas != calculation.formula_count {
        return Err(calculation_storage_error(
            "use.office.spreadsheet_formula_storage_invalid",
            "Calculation results do not match the workbook formula cells.",
        )
        .with_detail("expectedFormulas", calculation.formula_count)
        .with_detail("plannedFormulas", planned_formulas));
    }
    Ok(plans)
}

pub(super) fn worksheet_cells(
    sheet: &DocumentNode,
) -> UseResult<BTreeMap<CellReference, &DocumentNode>> {
    let mut cells = BTreeMap::new();
    for cell in sheet
        .children
        .iter()
        .filter(|node| node.node_type == OfficeNodeType::Row)
        .flat_map(|row| &row.children)
        .filter(|node| node.node_type == OfficeNodeType::Cell)
    {
        let reference = cell
            .path
            .rsplit_once('/')
            .and_then(|(_, reference)| CellReference::parse(reference).ok())
            .ok_or_else(|| {
                calculation_storage_error(
                    "use.office.spreadsheet_formula_storage_invalid",
                    format!("Spreadsheet cell path '{}' is invalid.", cell.path),
                )
            })?;
        if cells.insert(reference, cell).is_some() {
            return Err(calculation_storage_error(
                "use.office.spreadsheet_formula_storage_invalid",
                format!("Worksheet contains duplicate cell '{}'.", reference.a1()),
            ));
        }
    }
    Ok(cells)
}

fn plan_old_spill_cleanup(
    cells: &BTreeMap<CellReference, &DocumentNode>,
    writes: &mut BTreeMap<CellReference, CellWrite>,
) -> UseResult<()> {
    for (anchor, cell) in cells {
        let Some(reference) = cell.format.get("formulaRef") else {
            continue;
        };
        let range = CellRange::parse(reference).map_err(|error| {
            calculation_storage_error(
                "use.office.spreadsheet_formula_storage_invalid",
                format!(
                    "Formula cell '{}' has invalid spill range '{reference}': {error}",
                    cell.path
                ),
            )
        })?;
        if !range.contains(*anchor) {
            return Err(calculation_storage_error(
                "use.office.spreadsheet_formula_storage_invalid",
                format!(
                    "Formula spill range '{}' does not contain anchor '{}'.",
                    range.a1(),
                    anchor.a1()
                ),
            ));
        }
        let spill_cells = range.cell_count()?;
        if spill_cells > MAX_SPREADSHEET_FORMULA_SPILL_CELLS {
            return Err(calculation_storage_error(
                "use.office.spreadsheet_formula_spill_limit",
                format!(
                    "Stored formula spill '{}' exceeds {MAX_SPREADSHEET_FORMULA_SPILL_CELLS} cells.",
                    range.a1()
                ),
            )
            .with_detail("cells", spill_cells));
        }
        for row in range.start.row..=range.end.row {
            for column in range.start.column..=range.end.column {
                let reference = CellReference { column, row };
                if reference != *anchor {
                    insert_clear_write(writes, reference)?;
                }
            }
        }
    }
    Ok(())
}

fn plan_calculated_cell(
    calculated: &SpreadsheetFormulaCalculatedCell,
    cells: &BTreeMap<CellReference, &DocumentNode>,
    writes: &mut BTreeMap<CellReference, CellWrite>,
) -> UseResult<()> {
    let anchor = CellReference {
        column: calculated.cell.column,
        row: calculated.cell.row,
    };
    let cell = cells.get(&anchor).ok_or_else(|| {
        calculation_storage_error(
            "use.office.spreadsheet_formula_storage_invalid",
            format!(
                "Calculated formula cell '{}' is missing.",
                calculated.cell.path()
            ),
        )
    })?;
    let expression = cell.format.get("formula").cloned().ok_or_else(|| {
        calculation_storage_error(
            "use.office.spreadsheet_formula_storage_invalid",
            format!("Calculated cell '{}' has no formula.", cell.path),
        )
    })?;
    if cell.format.get("formulaType").is_some_and(|value| {
        !value.eq_ignore_ascii_case("normal") && !value.eq_ignore_ascii_case("array")
    }) {
        return Err(calculation_storage_error(
            "use.office.spreadsheet_formula_storage_unsupported",
            format!(
                "Formula storage type '{}' at '{}' is not supported by native recalculation.",
                cell.format.get("formulaType").map_or("", String::as_str),
                cell.path
            ),
        ));
    }
    match &calculated.value {
        SpreadsheetFormulaValue::Array { rows } => {
            let spill = calculated.spill_range.as_deref().ok_or_else(|| {
                calculation_storage_error(
                    "use.office.spreadsheet_formula_storage_invalid",
                    format!("Array result '{}' has no spill range.", cell.path),
                )
            })?;
            let range = CellRange::parse(spill)?;
            let height = usize::try_from(range.end.row - range.start.row + 1)
                .map_err(|_| calculation_write_limit())?;
            let width = usize::try_from(range.end.column - range.start.column + 1)
                .map_err(|_| calculation_write_limit())?;
            if range.start != anchor
                || rows.len() != height
                || rows.iter().any(|row| row.len() != width)
            {
                return Err(calculation_storage_error(
                    "use.office.spreadsheet_formula_storage_invalid",
                    format!(
                        "Array result shape does not match spill range '{}' at '{}'.",
                        range.a1(),
                        cell.path
                    ),
                ));
            }
            for (row_offset, row) in rows.iter().enumerate() {
                for (column_offset, value) in row.iter().enumerate() {
                    require_scalar_value(value)?;
                    let reference = CellReference {
                        column: range.start.column
                            + u32::try_from(column_offset)
                                .map_err(|_| calculation_write_limit())?,
                        row: range.start.row
                            + u32::try_from(row_offset).map_err(|_| calculation_write_limit())?,
                    };
                    let write = if reference == anchor {
                        CellWrite::Formula {
                            expression: expression.clone(),
                            value: value.clone(),
                            spill_range: Some(range.a1()),
                        }
                    } else {
                        CellWrite::Cached(value.clone())
                    };
                    insert_planned_write(writes, reference, write)?;
                }
            }
        }
        value => {
            require_scalar_value(value)?;
            if calculated.spill_range.is_some() {
                return Err(calculation_storage_error(
                    "use.office.spreadsheet_formula_storage_invalid",
                    format!(
                        "Scalar result '{}' unexpectedly has a spill range.",
                        cell.path
                    ),
                ));
            }
            insert_planned_write(
                writes,
                anchor,
                CellWrite::Formula {
                    expression,
                    value: value.clone(),
                    spill_range: None,
                },
            )?;
        }
    }
    Ok(())
}

fn insert_planned_write(
    writes: &mut BTreeMap<CellReference, CellWrite>,
    reference: CellReference,
    write: CellWrite,
) -> UseResult<()> {
    match writes.get(&reference) {
        None | Some(CellWrite::Clear) => {
            if !writes.contains_key(&reference) {
                let cells = writes
                    .len()
                    .checked_add(1)
                    .ok_or_else(calculation_write_limit)?;
                if cells > MAX_CALCULATION_WRITES {
                    return Err(calculation_write_limit().with_detail("cells", cells));
                }
            }
            writes.insert(reference, write);
            Ok(())
        }
        Some(_) => Err(calculation_storage_error(
            "use.office.spreadsheet_formula_spill_overlap",
            format!(
                "Calculated formula results overlap at '{}'.",
                reference.a1()
            ),
        )),
    }
}

fn validate_planned_writes(
    cells: &BTreeMap<CellReference, &DocumentNode>,
    writes: &BTreeMap<CellReference, CellWrite>,
) -> UseResult<()> {
    for (reference, write) in writes {
        if matches!(write, CellWrite::Formula { .. }) {
            continue;
        }
        if cells
            .get(reference)
            .is_some_and(|cell| cell.format.contains_key("formula"))
        {
            return Err(calculation_storage_error(
                "use.office.spreadsheet_formula_spill_overlap",
                format!(
                    "Calculated spill at '{}' overlaps another formula cell.",
                    reference.a1()
                ),
            ));
        }
    }
    Ok(())
}

fn require_scalar_value(value: &SpreadsheetFormulaValue) -> UseResult<()> {
    if matches!(value, SpreadsheetFormulaValue::Array { .. }) {
        return Err(calculation_storage_error(
            "use.office.spreadsheet_formula_storage_invalid",
            "Nested Spreadsheet formula arrays cannot be written to OOXML cells.",
        ));
    }
    Ok(())
}
