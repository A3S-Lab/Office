import { indexToColumnChar, type Selection } from '@fortune-sheet/core';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetSheet,
} from '../work-types';
import { finiteSpreadsheetSelection } from './spreadsheet-editor-support';

export type SpreadsheetFreezePanePreset =
  | 'selection'
  | 'topRow'
  | 'firstColumn'
  | 'none';

type SpreadsheetFrozenPane = NonNullable<WorkSpreadsheetSheet['frozen']>;

export interface SpreadsheetFreezePaneBoundary {
  columns: number;
  rows: number;
}

export function spreadsheetFreezePaneForSelection(
  selection: Selection,
): SpreadsheetFrozenPane | null {
  const finite = finiteSpreadsheetSelection(selection);
  return spreadsheetFrozenPaneFromCounts(
    finite.row_focus ?? finite.row[0] ?? 0,
    finite.column_focus ?? finite.column[0] ?? 0,
  );
}

export function spreadsheetFreezePanesSelectionLabel(
  selection: Selection,
): string {
  const boundary = spreadsheetFreezePaneForSelection(selection);
  if (!boundary) return '冻结至当前单元格';
  const { rows, columns } = spreadsheetFreezePaneBoundary(boundary);
  if (rows && columns) {
    return `冻结至第 ${rows} 行、${indexToColumnChar(columns - 1)} 列`;
  }
  if (rows) return `冻结至第 ${rows} 行`;
  return `冻结至 ${indexToColumnChar(columns - 1)} 列`;
}

export function spreadsheetFreezePaneBoundary(
  frozen: WorkSpreadsheetSheet['frozen'],
): SpreadsheetFreezePaneBoundary {
  if (!frozen) return { rows: 0, columns: 0 };
  const freezesRows = ['row', 'both', 'rangeRow', 'rangeBoth'].includes(
    frozen.type,
  );
  const freezesColumns = [
    'column',
    'both',
    'rangeColumn',
    'rangeBoth',
  ].includes(frozen.type);
  return {
    rows: freezesRows
      ? Math.max(1, finiteFreezeIndex(frozen.range?.row_focus) + 1)
      : 0,
    columns: freezesColumns
      ? Math.max(1, finiteFreezeIndex(frozen.range?.column_focus) + 1)
      : 0,
  };
}

export function spreadsheetFreezePanesStatus(
  frozen: WorkSpreadsheetSheet['frozen'],
): string {
  const { rows, columns } = spreadsheetFreezePaneBoundary(frozen);
  if (rows && columns) return `已冻结前 ${rows} 行和前 ${columns} 列。`;
  if (rows) return `已冻结前 ${rows} 行。`;
  if (columns) return `已冻结前 ${columns} 列。`;
  return '';
}

export function updateSpreadsheetFreezePanes(
  content: WorkSpreadsheetContent,
  sheetId: string,
  preset: SpreadsheetFreezePanePreset,
  selection: Selection,
): WorkSpreadsheetContent | null {
  const sheetIndex = content.sheets.findIndex((sheet) => sheet.id === sheetId);
  const sheet = content.sheets[sheetIndex];
  if (!sheet) return null;
  const nextFrozen = spreadsheetFrozenPaneForPreset(preset, selection);
  if (preset !== 'none' && !nextFrozen) return null;
  if (!nextFrozen && !sheet.frozen) return null;
  if (nextFrozen && sameSpreadsheetFreezePane(sheet.frozen, nextFrozen))
    return null;

  const savedSelection = finiteSpreadsheetSelection(selection);
  const nextSheet = nextFrozen
    ? {
        ...sheet,
        frozen: nextFrozen,
        luckysheet_select_save: [savedSelection],
      }
    : spreadsheetSheetWithoutFreezePanes(sheet, savedSelection);
  const sheets = [...content.sheets];
  sheets[sheetIndex] = nextSheet;
  return { ...content, sheets };
}

function spreadsheetFrozenPaneForPreset(
  preset: SpreadsheetFreezePanePreset,
  selection: Selection,
): SpreadsheetFrozenPane | null {
  if (preset === 'selection') {
    return spreadsheetFreezePaneForSelection(selection);
  }
  if (preset === 'topRow') return spreadsheetFrozenPaneFromCounts(1, 0);
  if (preset === 'firstColumn') return spreadsheetFrozenPaneFromCounts(0, 1);
  return null;
}

function spreadsheetFrozenPaneFromCounts(
  rows: number,
  columns: number,
): SpreadsheetFrozenPane | null {
  const safeRows = finiteFreezeIndex(rows);
  const safeColumns = finiteFreezeIndex(columns);
  if (!safeRows && !safeColumns) return null;
  return {
    type:
      safeRows && safeColumns
        ? 'rangeBoth'
        : safeRows
          ? 'rangeRow'
          : 'rangeColumn',
    range: {
      row_focus: Math.max(0, safeRows - 1),
      column_focus: Math.max(0, safeColumns - 1),
    },
  };
}

function spreadsheetSheetWithoutFreezePanes(
  sheet: WorkSpreadsheetSheet,
  selection: Selection,
): WorkSpreadsheetSheet {
  const { frozen: _frozen, ...remaining } = sheet;
  return { ...remaining, luckysheet_select_save: [selection] };
}

function sameSpreadsheetFreezePane(
  current: WorkSpreadsheetSheet['frozen'],
  next: WorkSpreadsheetSheet['frozen'],
): boolean {
  const currentBoundary = spreadsheetFreezePaneBoundary(current);
  const nextBoundary = spreadsheetFreezePaneBoundary(next);
  return (
    currentBoundary.rows === nextBoundary.rows &&
    currentBoundary.columns === nextBoundary.columns
  );
}

function finiteFreezeIndex(value: unknown): number {
  return typeof value === 'number' && Number.isFinite(value)
    ? Math.max(0, Math.trunc(value))
    : 0;
}
