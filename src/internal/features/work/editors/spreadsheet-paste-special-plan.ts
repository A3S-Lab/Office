import type { SpreadsheetCellRange } from './spreadsheet-cell-range';
import {
  MAX_SPREADSHEET_COLUMNS,
  MAX_SPREADSHEET_PASTE_SPECIAL_CELLS,
  MAX_SPREADSHEET_ROWS,
  type SpreadsheetClipboardSnapshot,
  type SpreadsheetPastePlan,
  type SpreadsheetPasteSource,
} from './spreadsheet-paste-special-types';

export function planSpreadsheetPaste(
  snapshot: SpreadsheetClipboardSnapshot,
  selection: SpreadsheetCellRange,
  transpose: boolean,
): SpreadsheetPastePlan | null {
  const rowCount = transpose ? snapshot.columnCount : snapshot.rowCount;
  const columnCount = transpose ? snapshot.rowCount : snapshot.columnCount;
  if (!rowCount || !columnCount) return null;

  const selectedRows = selection.row[1] - selection.row[0] + 1;
  const selectedColumns = selection.column[1] - selection.column[0] + 1;
  const singleCell = selectedRows === 1 && selectedColumns === 1;
  if (
    !singleCell &&
    (selectedRows % rowCount !== 0 || selectedColumns % columnCount !== 0)
  ) {
    return null;
  }

  const targetRows = singleCell ? rowCount : selectedRows;
  const targetColumns = singleCell ? columnCount : selectedColumns;
  const endRow = selection.row[0] + targetRows - 1;
  const endColumn = selection.column[0] + targetColumns - 1;
  if (
    selection.row[0] < 0 ||
    selection.column[0] < 0 ||
    endRow >= MAX_SPREADSHEET_ROWS ||
    endColumn >= MAX_SPREADSHEET_COLUMNS ||
    targetRows * targetColumns > MAX_SPREADSHEET_PASTE_SPECIAL_CELLS
  ) {
    return null;
  }

  return {
    rowCount,
    columnCount,
    targetRange: {
      row: [selection.row[0], endRow],
      column: [selection.column[0], endColumn],
    },
  };
}

export function spreadsheetPasteSourceAt(
  snapshot: SpreadsheetClipboardSnapshot,
  plan: SpreadsheetPastePlan,
  row: number,
  column: number,
  transpose: boolean,
): SpreadsheetPasteSource {
  const targetRow = (row - plan.targetRange.row[0]) % plan.rowCount;
  const targetColumn = (column - plan.targetRange.column[0]) % plan.columnCount;
  const sourceRowIndex = transpose ? targetColumn : targetRow;
  const sourceColumnIndex = transpose ? targetRow : targetColumn;
  return {
    cell: snapshot.cells[sourceRowIndex]![sourceColumnIndex]!,
    sourceRow: snapshot.sourceRange.row[0] + sourceRowIndex,
    sourceColumn: snapshot.sourceRange.column[0] + sourceColumnIndex,
  };
}
