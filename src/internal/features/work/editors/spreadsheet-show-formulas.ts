import type { Cell } from '@fortune-sheet/core';
import type { WorkSpreadsheetSheet } from '../work-types';

/**
 * WPS/Excel "Show Formulas" (Ctrl+`): paint formula text in cells that have
 * native formulas. Projection is display-only and must not be written back.
 */
export function spreadsheetFormulaDisplayText(
  cell: Cell | null | undefined,
): string | null {
  if (typeof cell?.f !== 'string') return null;
  const formula = cell.f.trim();
  if (!formula) return null;
  return formula.startsWith('=') ? formula : `=${formula}`;
}

export function projectSpreadsheetCellShowingFormula(
  cell: Cell | null | undefined,
): Cell | null | undefined {
  const display = spreadsheetFormulaDisplayText(cell);
  if (!display || !cell) return cell;
  if (cell.m === display && cell.v === display) return cell;
  return {
    ...cell,
    m: display,
    v: display,
  };
}

export function projectSpreadsheetSheetShowingFormulas(
  sheet: WorkSpreadsheetSheet,
): WorkSpreadsheetSheet {
  const data = sheet.data;
  if (!data?.length && !sheet.celldata?.length) return sheet;

  let changed = false;
  const nextData = data?.map((row) => {
    if (!row) return row;
    let rowChanged = false;
    const nextRow = row.map((cell) => {
      const nextCell = projectSpreadsheetCellShowingFormula(cell);
      if (nextCell !== cell) rowChanged = true;
      return nextCell ?? null;
    });
    if (!rowChanged) return row;
    changed = true;
    return nextRow;
  });

  // Sparse celldata entries also paint through Fortune when dense data is thin.
  const celldata = sheet.celldata;
  let nextCelldata = celldata;
  if (celldata?.length) {
    let celldataChanged = false;
    nextCelldata = celldata.map((entry) => {
      const nextValue = projectSpreadsheetCellShowingFormula(entry.v);
      if (nextValue === entry.v) return entry;
      celldataChanged = true;
      return { ...entry, v: nextValue ?? null };
    });
    if (celldataChanged) changed = true;
  }

  if (!changed) return sheet;
  return {
    ...sheet,
    ...(nextData ? { data: nextData } : {}),
    ...(nextCelldata === celldata ? {} : { celldata: nextCelldata }),
  };
}

export function projectSpreadsheetSheetsShowingFormulas(
  sheets: readonly WorkSpreadsheetSheet[],
): WorkSpreadsheetSheet[] {
  return sheets.map((sheet) => projectSpreadsheetSheetShowingFormulas(sheet));
}
