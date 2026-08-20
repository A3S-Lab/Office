import type { Cell } from '@fortune-sheet/core';
import { sparseMatrixColumnCount } from '../spreadsheet-sparse';
import type { WorkSpreadsheetSheet } from '../work-types';
import { isOfficeShortcutBlocked } from './office-shortcuts';
import { isSpreadsheetNativeTextUndoTarget } from './spreadsheet-editor-support';

export type SpreadsheetSelectionMove =
  | 'down'
  | 'left'
  | 'next-cell'
  | 'page-down'
  | 'page-up'
  | 'previous-cell'
  | 'right'
  | 'row-start'
  | 'sheet-start'
  | 'up'
  | 'used-end';

export type SpreadsheetSelectionScope = 'all' | 'column' | 'row';

export interface SpreadsheetKeyboardSelection {
  row: number[];
  column: number[];
  row_focus?: number;
  column_focus?: number;
}

const spreadsheetKeyboardPageRows = 20;
const spreadsheetDefaultRows = 60;
const spreadsheetDefaultColumns = 26;

export function moveSpreadsheetKeyboardSelection(
  sheet: WorkSpreadsheetSheet,
  selection: SpreadsheetKeyboardSelection,
  move: SpreadsheetSelectionMove,
  extend: boolean,
): SpreadsheetKeyboardSelection {
  const bounds = spreadsheetSheetBounds(sheet);
  const current = normalizeSpreadsheetKeyboardSelection(selection, bounds);
  let targetRow = current.rowFocus;
  let targetColumn = current.columnFocus;

  switch (move) {
    case 'up':
      targetRow -= 1;
      break;
    case 'down':
      targetRow += 1;
      break;
    case 'left':
      targetColumn -= 1;
      break;
    case 'right':
      targetColumn += 1;
      break;
    case 'next-cell':
      if (targetColumn < bounds.lastColumn) targetColumn += 1;
      else if (targetRow < bounds.lastRow) {
        targetRow += 1;
        targetColumn = 0;
      }
      break;
    case 'previous-cell':
      if (targetColumn > 0) targetColumn -= 1;
      else if (targetRow > 0) {
        targetRow -= 1;
        targetColumn = bounds.lastColumn;
      }
      break;
    case 'page-up':
      targetRow -= spreadsheetKeyboardPageRows;
      break;
    case 'page-down':
      targetRow += spreadsheetKeyboardPageRows;
      break;
    case 'row-start':
      targetColumn = 0;
      break;
    case 'sheet-start':
      targetRow = 0;
      targetColumn = 0;
      break;
    case 'used-end': {
      const usedEnd = spreadsheetUsedRangeEnd(sheet);
      targetRow = usedEnd.row;
      targetColumn = usedEnd.column;
      break;
    }
  }

  targetRow = clampSpreadsheetIndex(targetRow, bounds.lastRow);
  targetColumn = clampSpreadsheetIndex(targetColumn, bounds.lastColumn);
  if (!extend) {
    return {
      row: [targetRow, targetRow],
      column: [targetColumn, targetColumn],
      row_focus: targetRow,
      column_focus: targetColumn,
    };
  }

  const rowAnchor = selectionAnchor(
    current.startRow,
    current.endRow,
    current.rowFocus,
  );
  const columnAnchor = selectionAnchor(
    current.startColumn,
    current.endColumn,
    current.columnFocus,
  );
  return {
    row: [Math.min(rowAnchor, targetRow), Math.max(rowAnchor, targetRow)],
    column: [
      Math.min(columnAnchor, targetColumn),
      Math.max(columnAnchor, targetColumn),
    ],
    row_focus: targetRow,
    column_focus: targetColumn,
  };
}

export function scopeSpreadsheetKeyboardSelection(
  sheet: WorkSpreadsheetSheet,
  selection: SpreadsheetKeyboardSelection,
  scope: SpreadsheetSelectionScope,
): SpreadsheetKeyboardSelection {
  const bounds = spreadsheetSheetBounds(sheet);
  const current = normalizeSpreadsheetKeyboardSelection(selection, bounds);
  if (scope === 'column') {
    return {
      row: [0, bounds.lastRow],
      column: [current.columnFocus, current.columnFocus],
      row_focus: current.rowFocus,
      column_focus: current.columnFocus,
    };
  }
  if (scope === 'row') {
    return {
      row: [current.rowFocus, current.rowFocus],
      column: [0, bounds.lastColumn],
      row_focus: current.rowFocus,
      column_focus: current.columnFocus,
    };
  }
  return {
    row: [0, bounds.lastRow],
    column: [0, bounds.lastColumn],
    row_focus: current.rowFocus,
    column_focus: current.columnFocus,
  };
}

export function isSpreadsheetGridKeyboardTarget(
  target: EventTarget | null,
): boolean {
  if (target === null) return true;
  return (
    target instanceof Element && Boolean(target.closest('.fortune-container'))
  );
}

export function runSpreadsheetSelectionMoveShortcut(
  event: KeyboardEvent,
  canExecute: (move: SpreadsheetSelectionMove, extend: boolean) => boolean,
  execute: (move: SpreadsheetSelectionMove, extend: boolean) => boolean,
  move: SpreadsheetSelectionMove,
  extend: boolean,
): boolean {
  if (
    event.isComposing ||
    isOfficeShortcutBlocked(event.target) ||
    isSpreadsheetNativeTextUndoTarget(event.target) ||
    !isSpreadsheetGridKeyboardTarget(event.target) ||
    !canExecute(move, extend)
  ) {
    return false;
  }
  return execute(move, extend);
}

export function runSpreadsheetSelectionScopeShortcut(
  event: KeyboardEvent,
  canExecute: (scope: SpreadsheetSelectionScope) => boolean,
  execute: (scope: SpreadsheetSelectionScope) => boolean,
  scope: SpreadsheetSelectionScope,
): boolean {
  if (
    event.repeat ||
    event.isComposing ||
    isOfficeShortcutBlocked(event.target) ||
    isSpreadsheetNativeTextUndoTarget(event.target) ||
    !isSpreadsheetGridKeyboardTarget(event.target) ||
    !canExecute(scope)
  ) {
    return false;
  }
  return execute(scope);
}

export function spreadsheetSelectionContainsFocus(
  selection: SpreadsheetKeyboardSelection,
  focus: { column: number; row: number },
): boolean {
  const firstRow = selection.row[0] ?? 0;
  const secondRow = selection.row[1] ?? firstRow;
  const firstColumn = selection.column[0] ?? 0;
  const secondColumn = selection.column[1] ?? firstColumn;
  return (
    focus.row >= Math.min(firstRow, secondRow) &&
    focus.row <= Math.max(firstRow, secondRow) &&
    focus.column >= Math.min(firstColumn, secondColumn) &&
    focus.column <= Math.max(firstColumn, secondColumn)
  );
}

export function spreadsheetSheetBounds(sheet: WorkSpreadsheetSheet): {
  lastColumn: number;
  lastRow: number;
} {
  const dataRows = sheet.data?.length ?? 0;
  const dataColumns = sparseMatrixColumnCount(sheet.data);
  return {
    lastRow:
      positiveSpreadsheetDimension(
        sheet.row,
        dataRows,
        spreadsheetDefaultRows,
      ) - 1,
    lastColumn:
      positiveSpreadsheetDimension(
        sheet.column,
        dataColumns,
        spreadsheetDefaultColumns,
      ) - 1,
  };
}

function positiveSpreadsheetDimension(
  configured: unknown,
  dataDimension: number,
  fallback: number,
): number {
  const value = Number(configured);
  if (Number.isFinite(value) && value > 0)
    return Math.max(1, Math.trunc(value));
  if (dataDimension > 0) return dataDimension;
  return fallback;
}

function spreadsheetUsedRangeEnd(sheet: WorkSpreadsheetSheet): {
  column: number;
  row: number;
} {
  const bounds = spreadsheetSheetBounds(sheet);
  let row = 0;
  let column = 0;
  sheet.data?.forEach((cells, rowIndex) => {
    cells?.forEach((cell, columnIndex) => {
      if (!spreadsheetCellHasValue(cell)) return;
      row = Math.max(row, rowIndex);
      column = Math.max(column, columnIndex);
    });
  });
  sheet.celldata?.forEach((cell) => {
    if (!spreadsheetCellHasValue(cell.v)) return;
    row = Math.max(row, Number(cell.r) || 0);
    column = Math.max(column, Number(cell.c) || 0);
  });
  return {
    row: clampSpreadsheetIndex(row, bounds.lastRow),
    column: clampSpreadsheetIndex(column, bounds.lastColumn),
  };
}

function spreadsheetCellHasValue(cell: Cell | null | undefined): boolean {
  if (!cell) return false;
  if (typeof cell.f === 'string' && cell.f.trim()) return true;
  const value = cell.v;
  if (value !== null && value !== undefined) {
    return typeof value !== 'string' || value.trim().length > 0;
  }
  return typeof cell.m === 'string' && cell.m.trim().length > 0;
}

function normalizeSpreadsheetKeyboardSelection(
  selection: SpreadsheetKeyboardSelection,
  bounds: { lastColumn: number; lastRow: number },
) {
  const firstRow = clampSpreadsheetIndex(selection.row[0], bounds.lastRow);
  const secondRow = clampSpreadsheetIndex(
    selection.row[1] ?? firstRow,
    bounds.lastRow,
  );
  const firstColumn = clampSpreadsheetIndex(
    selection.column[0],
    bounds.lastColumn,
  );
  const secondColumn = clampSpreadsheetIndex(
    selection.column[1] ?? firstColumn,
    bounds.lastColumn,
  );
  const startRow = Math.min(firstRow, secondRow);
  const endRow = Math.max(firstRow, secondRow);
  const startColumn = Math.min(firstColumn, secondColumn);
  const endColumn = Math.max(firstColumn, secondColumn);
  return {
    startRow,
    endRow,
    startColumn,
    endColumn,
    rowFocus: clampSpreadsheetIndex(
      selection.row_focus ?? secondRow,
      bounds.lastRow,
    ),
    columnFocus: clampSpreadsheetIndex(
      selection.column_focus ?? secondColumn,
      bounds.lastColumn,
    ),
  };
}

function clampSpreadsheetIndex(value: unknown, maximum: number): number {
  const numeric = Number(value);
  const finite = Number.isFinite(numeric) ? Math.trunc(numeric) : 0;
  return Math.min(maximum, Math.max(0, finite));
}

function selectionAnchor(start: number, end: number, focus: number): number {
  return focus === start ? end : start;
}
