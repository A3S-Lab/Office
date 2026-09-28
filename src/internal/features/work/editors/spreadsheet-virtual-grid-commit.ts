import type { Cell, CellMatrix } from '@fortune-sheet/core';
import { cloneSparseMatrix } from '../spreadsheet-sparse';
import type { WorkSpreadsheetContent } from '../work-types';

/**
 * Immutable sparse cell write for the A3S virtual grid edit path.
 * Avoids Fortune `setCellValue` so fortuneReady sheets can commit without the
 * vendor remount/edit surface.
 */
export function spreadsheetVirtualGridCommitCell(
  content: WorkSpreadsheetContent,
  sheetId: string,
  row: number,
  column: number,
  raw: string,
): WorkSpreadsheetContent {
  const sheetIndex = content.sheets.findIndex((sheet) => sheet.id === sheetId);
  if (sheetIndex < 0) return content;
  const sheet = content.sheets[sheetIndex];
  if (!sheet) return content;
  if (
    !Number.isSafeInteger(row) ||
    !Number.isSafeInteger(column) ||
    row < 0 ||
    column < 0
  ) {
    return content;
  }

  const data = cloneSparseMatrix(sheet.data);
  ensureSparseCellSlot(data, row, column);
  const rowCells = data[row];
  if (!rowCells) return content;

  const trimmed = raw.trim();
  if (trimmed.length === 0) {
    rowCells[column] = null;
  } else {
    rowCells[column] = spreadsheetVirtualGridLiteralCell(trimmed);
  }

  const sheets = content.sheets.slice();
  sheets[sheetIndex] = {
    ...sheet,
    data,
    row: Math.max(sheet.row ?? 0, row + 1),
    column: Math.max(sheet.column ?? 0, column + 1),
  };
  return { ...content, sheets };
}

function ensureSparseCellSlot(
  data: CellMatrix,
  row: number,
  column: number,
): void {
  if (data.length < row + 1) data.length = row + 1;
  let rowCells = data[row];
  if (!rowCells) {
    rowCells = [];
    data[row] = rowCells;
  }
  if (rowCells.length < column + 1) rowCells.length = column + 1;
}

function spreadsheetVirtualGridLiteralCell(raw: string): Cell {
  if (/^[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?$/.test(raw)) {
    const value = Number(raw);
    if (Number.isFinite(value)) return { v: value, m: raw };
  }
  return { v: raw, m: raw };
}
