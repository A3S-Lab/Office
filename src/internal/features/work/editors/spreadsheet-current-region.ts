import type { Cell } from '@fortune-sheet/core';
import { sparseMatrixColumnCount } from '../spreadsheet-sparse';
import type { WorkSpreadsheetSheet } from '../work-types';
import {
  normalizeSpreadsheetCellRange,
  type SpreadsheetCellRange,
  type SpreadsheetCellRangeInput,
} from './spreadsheet-cell-range';

/**
 * Finds the bounded rectangular current region that contains a seed range.
 * Dense and sparse worksheets share this reader so filter and sort commands do
 * not derive different expansion boundaries from the same controlled sheet.
 */
export function spreadsheetCurrentRegion(
  sheet: WorkSpreadsheetSheet,
  seed: SpreadsheetCellRangeInput,
): SpreadsheetCellRange | null {
  const range = normalizeSpreadsheetCellRange(seed);
  if (!range) return null;
  const cells = spreadsheetSheetCellReader(sheet);
  const bounds = spreadsheetCurrentRegionUsedBounds(sheet);
  let startRow = range.row[0];
  let endRow = range.row[1];
  let startColumn = range.column[0];
  let endColumn = range.column[1];
  if (
    startRow > bounds.lastRow ||
    startColumn > bounds.lastColumn ||
    !spreadsheetCellHasContent(cells(startRow, startColumn))
  ) {
    return null;
  }

  let changed = true;
  while (changed) {
    changed = false;
    while (
      startColumn > 0 &&
      spreadsheetCurrentRegionColumnHasContent(
        cells,
        startColumn - 1,
        startRow,
        endRow,
      )
    ) {
      startColumn -= 1;
      changed = true;
    }
    while (
      endColumn < bounds.lastColumn &&
      spreadsheetCurrentRegionColumnHasContent(
        cells,
        endColumn + 1,
        startRow,
        endRow,
      )
    ) {
      endColumn += 1;
      changed = true;
    }
    while (
      startRow > 0 &&
      spreadsheetRowHasContent(cells, startRow - 1, startColumn, endColumn)
    ) {
      startRow -= 1;
      changed = true;
    }
    while (
      endRow < bounds.lastRow &&
      spreadsheetRowHasContent(cells, endRow + 1, startColumn, endColumn)
    ) {
      endRow += 1;
      changed = true;
    }
  }

  return {
    row: [startRow, endRow],
    column: [startColumn, endColumn],
  };
}

export function spreadsheetRowHasContent(
  cellAt: (row: number, column: number) => Cell | null,
  row: number,
  startColumn: number,
  endColumn: number,
): boolean {
  for (let column = startColumn; column <= endColumn; column += 1) {
    if (spreadsheetCellHasContent(cellAt(row, column))) {
      return true;
    }
  }
  return false;
}

function spreadsheetCurrentRegionColumnHasContent(
  cellAt: (row: number, column: number) => Cell | null,
  column: number,
  startRow: number,
  endRow: number,
): boolean {
  for (let row = startRow; row <= endRow; row += 1) {
    if (spreadsheetCellHasContent(cellAt(row, column))) {
      return true;
    }
  }
  return false;
}

export function spreadsheetCellHasContent(cell: Cell | null): boolean {
  if (!cell) return false;
  return [cell.v, cell.m, cell.f].some(
    (value) => value !== undefined && value !== null && value !== '',
  );
}

export function spreadsheetSheetCellReader(
  sheet: WorkSpreadsheetSheet,
): (row: number, column: number) => Cell | null {
  if (sheet.data) {
    return (row, column) => sheet.data?.[row]?.[column] ?? null;
  }
  const cells = new Map(
    (sheet.celldata ?? []).map((cell) => [`${cell.r}:${cell.c}`, cell.v]),
  );
  return (row, column) => cells.get(`${row}:${column}`) ?? null;
}

function spreadsheetCurrentRegionUsedBounds(sheet: WorkSpreadsheetSheet): {
  lastColumn: number;
  lastRow: number;
} {
  let lastRow = Math.max((sheet.data?.length ?? 1) - 1, 0);
  let lastColumn = Math.max(sparseMatrixColumnCount(sheet.data) - 1, 0);
  for (const cell of sheet.celldata ?? []) {
    lastRow = Math.max(lastRow, cell.r);
    lastColumn = Math.max(lastColumn, cell.c);
  }
  return { lastColumn, lastRow };
}
