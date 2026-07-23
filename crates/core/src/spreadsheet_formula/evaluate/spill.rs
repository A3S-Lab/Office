use a3s_use_core::UseResult;

use crate::spreadsheet_reference::{CellReference, MAX_COLUMNS, MAX_ROWS};
use crate::{SpreadsheetFormulaErrorLiteral, SpreadsheetFormulaValue};

use super::{
    calculation_error, calculation_text_limit_error, ensure_formula_text_limit, public_scalar,
    spill_limit_error, EvalValue, EvaluationContext, FormulaCellKey, FormulaReferenceArea,
    ScalarValue, MAX_SPREADSHEET_FORMULA_CALCULATION_TEXT_BYTES,
    MAX_SPREADSHEET_FORMULA_SPILL_CELLS,
};

impl EvaluationContext<'_> {
    pub(super) fn finalize_formula_result(
        &mut self,
        anchor: FormulaCellKey,
        value: EvalValue,
        existing_spill_cells: usize,
        existing_text_bytes: usize,
    ) -> UseResult<(SpreadsheetFormulaValue, Option<String>, usize, usize)> {
        self.clear_previous_spill(anchor);
        let materialized = self.materialize(value)?;
        match materialized {
            EvalValue::Scalar(value) => {
                ensure_formula_text_limit(&value)?;
                let text_bytes = scalar_text_bytes(&value);
                ensure_calculation_text_limit(existing_text_bytes, text_bytes)?;
                self.values.insert(anchor, value.clone());
                Ok((public_scalar(&value), None, 0, text_bytes))
            }
            EvalValue::Array(array) => {
                let mut text_bytes = 0_usize;
                for value in array.rows.iter().flatten() {
                    ensure_formula_text_limit(value)?;
                    text_bytes = text_bytes
                        .checked_add(scalar_text_bytes(value))
                        .ok_or_else(calculation_text_limit_error)?;
                }
                let height = u32::try_from(array.height()).map_err(|_| spill_limit_error())?;
                let width = u32::try_from(array.width()).map_err(|_| spill_limit_error())?;
                let cells = array
                    .height()
                    .checked_mul(array.width())
                    .ok_or_else(spill_limit_error)?;
                if cells > MAX_SPREADSHEET_FORMULA_SPILL_CELLS {
                    return Err(spill_limit_error().with_detail("cells", cells));
                }
                let spill_cells = cells.saturating_sub(1);
                let end_row = anchor
                    .row
                    .checked_add(height.saturating_sub(1))
                    .filter(|row| *row <= MAX_ROWS);
                let end_column = anchor
                    .column
                    .checked_add(width.saturating_sub(1))
                    .filter(|column| *column <= MAX_COLUMNS);
                let (Some(end_row), Some(end_column)) = (end_row, end_column) else {
                    let error = ScalarValue::Error(SpreadsheetFormulaErrorLiteral::Spill);
                    self.values.insert(anchor, error.clone());
                    return Ok((public_scalar(&error), None, 0, 0));
                };
                let area = FormulaReferenceArea {
                    sheet: anchor.sheet,
                    start_column: anchor.column,
                    start_row: anchor.row,
                    end_column,
                    end_row,
                };
                if self.spill_is_blocked(anchor, area) {
                    let error = ScalarValue::Error(SpreadsheetFormulaErrorLiteral::Spill);
                    self.values.insert(anchor, error.clone());
                    return Ok((public_scalar(&error), None, 0, 0));
                }
                let total_spill_cells = existing_spill_cells
                    .checked_add(spill_cells)
                    .ok_or_else(spill_limit_error)?;
                if total_spill_cells > MAX_SPREADSHEET_FORMULA_SPILL_CELLS {
                    return Err(spill_limit_error().with_detail("cells", total_spill_cells));
                }
                ensure_calculation_text_limit(existing_text_bytes, text_bytes)?;
                for (row_offset, row) in array.rows.iter().enumerate() {
                    for (column_offset, value) in row.iter().enumerate() {
                        let key = FormulaCellKey {
                            sheet: anchor.sheet,
                            column: anchor.column
                                + u32::try_from(column_offset).map_err(|_| spill_limit_error())?,
                            row: anchor.row
                                + u32::try_from(row_offset).map_err(|_| spill_limit_error())?,
                        };
                        self.values.insert(key, value.clone());
                        if key != anchor {
                            self.spill_owners.insert(key, anchor);
                        }
                    }
                }
                self.spills.insert(anchor, area);
                let start = CellReference {
                    column: anchor.column,
                    row: anchor.row,
                }
                .a1();
                let end = CellReference {
                    column: end_column,
                    row: end_row,
                }
                .a1();
                Ok((
                    SpreadsheetFormulaValue::Array {
                        rows: array
                            .rows
                            .iter()
                            .map(|row| row.iter().map(public_scalar).collect())
                            .collect(),
                    },
                    Some(if start == end {
                        start
                    } else {
                        format!("{start}:{end}")
                    }),
                    spill_cells,
                    text_bytes,
                ))
            }
            EvalValue::Reference(_) => Err(calculation_error(
                "use.office.spreadsheet_formula_calculation_invalid",
                "Formula result retained an unresolved reference.",
            )),
        }
    }

    fn clear_previous_spill(&mut self, anchor: FormulaCellKey) {
        let Some(area) = self.old_spills.get(&anchor).copied() else {
            return;
        };
        for row in area.start_row..=area.end_row {
            for column in area.start_column..=area.end_column {
                let key = FormulaCellKey {
                    sheet: area.sheet,
                    column,
                    row,
                };
                if key != anchor {
                    self.values.remove(&key);
                }
            }
        }
    }

    fn spill_is_blocked(&self, anchor: FormulaCellKey, area: FormulaReferenceArea) -> bool {
        for row in area.start_row..=area.end_row {
            for column in area.start_column..=area.end_column {
                let key = FormulaCellKey {
                    sheet: area.sheet,
                    column,
                    row,
                };
                if key == anchor {
                    continue;
                }
                if self.formula_cells.contains(&key) || self.spill_owners.contains_key(&key) {
                    return true;
                }
                let owned_before = self
                    .old_spills
                    .get(&anchor)
                    .is_some_and(|old| old.contains(key));
                if self.occupied.contains(&key) && !owned_before {
                    return true;
                }
            }
        }
        false
    }
}

fn scalar_text_bytes(value: &ScalarValue) -> usize {
    match value {
        ScalarValue::Text(value) => value.len(),
        _ => 0,
    }
}

fn ensure_calculation_text_limit(existing: usize, added: usize) -> UseResult<()> {
    let bytes = existing
        .checked_add(added)
        .ok_or_else(calculation_text_limit_error)?;
    if bytes > MAX_SPREADSHEET_FORMULA_CALCULATION_TEXT_BYTES {
        return Err(calculation_text_limit_error().with_detail("bytes", bytes));
    }
    Ok(())
}
