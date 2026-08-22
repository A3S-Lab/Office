import type { Cell } from '@fortune-sheet/core';
import { canMutateSpreadsheetCellRange } from './spreadsheet-cell-mutation-guard';
import type { SpreadsheetCellRange } from './spreadsheet-cell-range';
import { spreadsheetCommandCatalog } from './spreadsheet-command-catalog';
import type {
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import {
  finiteSpreadsheetSelection,
  isSpreadsheetNativeTextUndoTarget,
} from './spreadsheet-editor-support';
import { isSpreadsheetGridKeyboardTarget } from './spreadsheet-keyboard-navigation';
import {
  createOfficeEditorExtension,
  type OfficeEditorExtension,
} from './office-editor-extension';
import { isOfficeShortcutBlocked } from './office-shortcuts';
import {
  MAX_SPREADSHEET_COLUMNS,
  MAX_SPREADSHEET_ROWS,
} from './spreadsheet-paste-special-types';

const MILLISECONDS_PER_DAY = 86_400_000;
const MINUTES_PER_DAY = 1_440;
const EXCEL_1900_EPOCH_UTC = Date.UTC(1899, 11, 30);

export type SpreadsheetDateTimeKind = 'date' | 'time';

export interface SpreadsheetDateTimeEntry {
  format: NonNullable<Cell['ct']> & { fa: string; t: string };
  formulaBarValue: string;
  value: number;
}

interface SpreadsheetDateTimeTarget {
  range: SpreadsheetCellRange;
  sheetId: string;
}

export function spreadsheetDateTimeEntry(
  kind: SpreadsheetDateTimeKind,
  now: Date,
): SpreadsheetDateTimeEntry | null {
  if (!Number.isFinite(now.getTime())) return null;

  if (kind === 'date') {
    const year = now.getFullYear();
    const month = now.getMonth();
    const day = now.getDate();
    const value =
      (Date.UTC(year, month, day) - EXCEL_1900_EPOCH_UTC) /
      MILLISECONDS_PER_DAY;
    if (!Number.isFinite(value)) return null;
    return {
      format: { fa: 'yyyy-MM-dd', t: 'd' },
      formulaBarValue: `${padDatePart(year, 4)}-${padDatePart(month + 1)}-${padDatePart(day)}`,
      value,
    };
  }

  if (kind !== 'time') return null;
  const hours = now.getHours();
  const minutes = now.getMinutes();
  return {
    format: { fa: 'hh:mm', t: 'd' },
    formulaBarValue: `${padDatePart(hours)}:${padDatePart(minutes)}`,
    value: (hours * 60 + minutes) / MINUTES_PER_DAY,
  };
}

export function createSpreadsheetDateTimeExtension(): OfficeEditorExtension<
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands
> {
  return createOfficeEditorExtension<
    SpreadsheetCommandContext,
    SpreadsheetEditorCommands
  >({
    name: 'spreadsheetDateTime',
    addCommands: () => ({
      insertCurrentDateTime: {
        canExecute: (context, kind) =>
          isSpreadsheetDateTimeKind(kind) &&
          spreadsheetDateTimeTarget(context) !== null,
        execute: insertCurrentSpreadsheetDateTime,
      },
    }),
    addKeyboardShortcuts: () => ({
      [spreadsheetCommandCatalog.insertCurrentDate.shortcut.editor[0]]: (
        { can, commands },
        event,
      ) =>
        runSpreadsheetDateTimeShortcut(
          event,
          'date',
          can.insertCurrentDateTime,
          commands.insertCurrentDateTime,
        ),
      [spreadsheetCommandCatalog.insertCurrentTime.shortcut.editor[0]]: (
        { can, commands },
        event,
      ) =>
        runSpreadsheetDateTimeShortcut(
          event,
          'time',
          can.insertCurrentDateTime,
          commands.insertCurrentDateTime,
        ),
    }),
  });
}

function insertCurrentSpreadsheetDateTime(
  context: SpreadsheetCommandContext,
  kind: SpreadsheetDateTimeKind,
): boolean {
  if (!isSpreadsheetDateTimeKind(kind)) return false;
  const target = spreadsheetDateTimeTarget(context);
  const entry = spreadsheetDateTimeEntry(kind, new Date());
  if (!target || !entry || !context.workbook) return false;

  try {
    context.workbook.batchCallApis([
      {
        name: 'setCellValuesByRange',
        args: [[[entry.value]], target.range, { id: target.sheetId }],
      },
      {
        name: 'setCellFormatByRange',
        args: ['ct', entry.format, target.range, { id: target.sheetId }],
      },
    ]);
  } catch {
    return false;
  }

  try {
    context.formulaBar?.setValue(entry.formulaBarValue);
  } catch {
    // The cell mutation succeeded; a detached formula bar is best effort.
  }
  return true;
}

function spreadsheetDateTimeTarget(
  context: SpreadsheetCommandContext,
): SpreadsheetDateTimeTarget | null {
  const workbook = context.workbook;
  const sheetId = context.targetSheetId;
  if (
    !context.editable ||
    !workbook ||
    !sheetId ||
    sheetId !== context.activeSheetId
  ) {
    return null;
  }

  try {
    const selection = finiteSpreadsheetSelection(
      workbook.getSelection()?.at(-1) ?? context.fallbackRange,
    );
    const row = selection.row_focus ?? selection.row[0] ?? 0;
    const column = selection.column_focus ?? selection.column[0] ?? 0;
    if (
      !Number.isSafeInteger(row) ||
      row < 0 ||
      row >= MAX_SPREADSHEET_ROWS ||
      !Number.isSafeInteger(column) ||
      column < 0 ||
      column >= MAX_SPREADSHEET_COLUMNS
    ) {
      return null;
    }
    const range: SpreadsheetCellRange = {
      row: [row, row],
      column: [column, column],
    };
    const sheet = context.content.sheets.find(
      (candidate) => candidate.id === sheetId,
    );
    return canMutateSpreadsheetCellRange(sheet, range)
      ? { range, sheetId }
      : null;
  } catch {
    return null;
  }
}

function runSpreadsheetDateTimeShortcut(
  event: KeyboardEvent,
  kind: SpreadsheetDateTimeKind,
  canExecute: SpreadsheetEditorCommands['insertCurrentDateTime'],
  execute: SpreadsheetEditorCommands['insertCurrentDateTime'],
): boolean {
  if (
    event.repeat ||
    event.isComposing ||
    isOfficeShortcutBlocked(event.target) ||
    isSpreadsheetNativeTextUndoTarget(event.target) ||
    !isSpreadsheetGridKeyboardTarget(event.target) ||
    !canExecute(kind)
  ) {
    return false;
  }
  return execute(kind);
}

function isSpreadsheetDateTimeKind(
  value: unknown,
): value is SpreadsheetDateTimeKind {
  return value === 'date' || value === 'time';
}

function padDatePart(value: number, length = 2): string {
  return String(value).padStart(length, '0');
}
