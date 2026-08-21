import type { Cell } from '@fortune-sheet/core';
import type { WorkSpreadsheetContent } from '../work-types';
import { isOfficeShortcutBlocked } from './office-shortcuts';
import type {
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import { spreadsheetCommandCatalog } from './spreadsheet-command-catalog';
import {
  normalizeSpreadsheetCellRange,
  type SpreadsheetCellRange,
  type SpreadsheetCellRangeInput,
  spreadsheetCellRangeArea,
} from './spreadsheet-cell-range';
import {
  isSpreadsheetNativeTextUndoTarget,
  spreadsheetSingleRange,
} from './spreadsheet-editor-support';
import {
  DEFAULT_SPREADSHEET_FONT_SIZE,
  nextSpreadsheetFontSize,
  type SpreadsheetFontSizeDirection,
} from './spreadsheet-font-size';
import {
  createOfficeEditorExtension,
  type OfficeEditorExtension,
} from './office-editor-extension';

export { nextSpreadsheetFontSize } from './spreadsheet-font-size';
export type { SpreadsheetFontSizeDirection } from './spreadsheet-font-size';

export const MAX_SPREADSHEET_FONT_SIZE_CELLS = 10_000;

export interface SpreadsheetFontSizeApiCall {
  name: 'setCellFormatByRange';
  args: ['fs', number, SpreadsheetCellRange, { id: string }];
}

interface SpreadsheetFontSizeRun {
  key: string;
  range: SpreadsheetCellRange;
  size: number;
}

export function createSpreadsheetFontSizeExtension(): OfficeEditorExtension<
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands
> {
  const growShortcuts = spreadsheetCommandCatalog.growFont.shortcut.editor;
  const shrinkShortcuts = spreadsheetCommandCatalog.shrinkFont.shortcut.editor;
  return createOfficeEditorExtension<
    SpreadsheetCommandContext,
    SpreadsheetEditorCommands
  >({
    name: 'spreadsheetFontSize',
    addCommands: () => ({
      adjustFontSize: {
        canExecute: canAdjustSelectedFontSize,
        execute: adjustSelectedFontSize,
      },
    }),
    addKeyboardShortcuts: () => ({
      [growShortcuts[0]]: ({ can, commands }, event) =>
        runSpreadsheetFontSizeShortcut(
          event,
          can.adjustFontSize,
          commands.adjustFontSize,
          'grow',
        ),
      [growShortcuts[1]]: ({ can, commands }, event) =>
        runSpreadsheetFontSizeShortcut(
          event,
          can.adjustFontSize,
          commands.adjustFontSize,
          'grow',
        ),
      [shrinkShortcuts[0]]: ({ can, commands }, event) =>
        runSpreadsheetFontSizeShortcut(
          event,
          can.adjustFontSize,
          commands.adjustFontSize,
          'shrink',
        ),
      [shrinkShortcuts[1]]: ({ can, commands }, event) =>
        runSpreadsheetFontSizeShortcut(
          event,
          can.adjustFontSize,
          commands.adjustFontSize,
          'shrink',
        ),
    }),
  });
}

export function canAdjustSpreadsheetFontSize(
  content: WorkSpreadsheetContent,
  sheetId: string,
  range: SpreadsheetCellRangeInput,
  direction: SpreadsheetFontSizeDirection,
): boolean {
  const normalizedRange = normalizeSpreadsheetCellRange(range);
  const sheet = content.sheets.find((candidate) => candidate.id === sheetId);
  if (
    !sheet ||
    !normalizedRange ||
    !isSpreadsheetFontSizeDirection(direction) ||
    spreadsheetCellRangeArea(normalizedRange) > MAX_SPREADSHEET_FONT_SIZE_CELLS
  ) {
    return false;
  }

  const sparseCells = new Map<string, Cell>();
  for (const entry of sheet.celldata ?? []) {
    if (
      entry.r >= normalizedRange.row[0] &&
      entry.r <= normalizedRange.row[1] &&
      entry.c >= normalizedRange.column[0] &&
      entry.c <= normalizedRange.column[1] &&
      entry.v
    ) {
      sparseCells.set(`${entry.r}:${entry.c}`, entry.v);
    }
  }
  return canAdjustSpreadsheetFontSizeCells(
    normalizedRange.row[1] - normalizedRange.row[0] + 1,
    normalizedRange.column[1] - normalizedRange.column[0] + 1,
    (rowOffset, columnOffset) => {
      const row = normalizedRange.row[0] + rowOffset;
      const column = normalizedRange.column[0] + columnOffset;
      return (
        sheet.data?.[row]?.[column] ??
        sparseCells.get(`${row}:${column}`) ??
        null
      );
    },
    direction,
  );
}

export function spreadsheetFontSizeApiCalls(
  cells: readonly (readonly (Cell | null)[])[],
  range: SpreadsheetCellRangeInput,
  sheetId: string,
  direction: SpreadsheetFontSizeDirection,
): SpreadsheetFontSizeApiCall[] {
  const normalizedRange = normalizeSpreadsheetCellRange(range);
  if (
    !normalizedRange ||
    !sheetId ||
    !isSpreadsheetFontSizeDirection(direction) ||
    spreadsheetCellRangeArea(normalizedRange) > MAX_SPREADSHEET_FONT_SIZE_CELLS
  ) {
    return [];
  }

  const rectangles: SpreadsheetFontSizeRun[] = [];
  let active = new Map<string, SpreadsheetFontSizeRun>();
  const rowCount = normalizedRange.row[1] - normalizedRange.row[0] + 1;
  const columnCount = normalizedRange.column[1] - normalizedRange.column[0] + 1;

  for (let rowOffset = 0; rowOffset < rowCount; rowOffset += 1) {
    const row = normalizedRange.row[0] + rowOffset;
    const runs = spreadsheetFontSizeRowRuns(
      cells[rowOffset],
      row,
      normalizedRange.column[0],
      columnCount,
      direction,
    );
    const nextActive = new Map<string, SpreadsheetFontSizeRun>();
    for (const run of runs) {
      const signature = `${run.key}:${run.range.column[0]}:${run.range.column[1]}`;
      const previous = active.get(signature);
      if (previous && previous.range.row[1] === row - 1) {
        previous.range.row[1] = row;
        nextActive.set(signature, previous);
      } else {
        nextActive.set(signature, run);
      }
    }
    for (const [signature, run] of active) {
      if (!nextActive.has(signature)) rectangles.push(run);
    }
    active = nextActive;
  }
  rectangles.push(...active.values());
  rectangles.sort(
    (left, right) =>
      left.range.row[0] - right.range.row[0] ||
      left.range.column[0] - right.range.column[0] ||
      left.range.row[1] - right.range.row[1] ||
      left.range.column[1] - right.range.column[1],
  );

  return rectangles.map(({ range: targetRange, size }) => ({
    name: 'setCellFormatByRange',
    args: ['fs', size, targetRange, { id: sheetId }],
  }));
}

function spreadsheetFontSizeRowRuns(
  cells: readonly (Cell | null)[] | undefined,
  row: number,
  firstColumn: number,
  columnCount: number,
  direction: SpreadsheetFontSizeDirection,
): SpreadsheetFontSizeRun[] {
  const runs: SpreadsheetFontSizeRun[] = [];
  let current: SpreadsheetFontSizeRun | null = null;
  for (let columnOffset = 0; columnOffset < columnCount; columnOffset += 1) {
    const column = firstColumn + columnOffset;
    const size = adjustedSpreadsheetFontSize(cells?.[columnOffset], direction);
    if (size === null) {
      if (current) runs.push(current);
      current = null;
      continue;
    }
    const key = String(size);
    if (current && current.key === key) {
      current.range.column[1] = column;
      continue;
    }
    if (current) runs.push(current);
    current = {
      key,
      range: { row: [row, row], column: [column, column] },
      size,
    };
  }
  if (current) runs.push(current);
  return runs;
}

function adjustedSpreadsheetFontSize(
  cell: Cell | null | undefined,
  direction: SpreadsheetFontSizeDirection,
): number | null {
  const size =
    cell?.fs === undefined ? DEFAULT_SPREADSHEET_FONT_SIZE : Number(cell.fs);
  return nextSpreadsheetFontSize(size, direction);
}

function canAdjustSelectedFontSize(
  context: SpreadsheetCommandContext,
  direction: SpreadsheetFontSizeDirection,
): boolean {
  const workbook = context.workbook;
  const sheetId = context.targetSheetId;
  if (
    !context.editable ||
    !workbook ||
    !sheetId ||
    sheetId !== context.activeSheetId ||
    !isSpreadsheetFontSizeDirection(direction)
  ) {
    return false;
  }
  const range = liveSpreadsheetFontSizeRange(context);
  if (spreadsheetCellRangeArea(range) > MAX_SPREADSHEET_FONT_SIZE_CELLS) {
    return false;
  }
  try {
    const cells = workbook.getCellsByRange(range, { id: sheetId });
    return canAdjustSpreadsheetFontSizeCells(
      range.row[1] - range.row[0] + 1,
      range.column[1] - range.column[0] + 1,
      (rowOffset, columnOffset) => cells[rowOffset]?.[columnOffset] ?? null,
      direction,
    );
  } catch {
    return false;
  }
}

function adjustSelectedFontSize(
  context: SpreadsheetCommandContext,
  direction: SpreadsheetFontSizeDirection,
): boolean {
  const workbook = context.workbook;
  const sheetId = context.targetSheetId;
  if (
    !context.editable ||
    !workbook ||
    !sheetId ||
    sheetId !== context.activeSheetId ||
    !isSpreadsheetFontSizeDirection(direction)
  ) {
    return false;
  }
  const range = liveSpreadsheetFontSizeRange(context);
  if (spreadsheetCellRangeArea(range) > MAX_SPREADSHEET_FONT_SIZE_CELLS) {
    return false;
  }
  try {
    const cells = workbook.getCellsByRange(range, { id: sheetId });
    const calls = spreadsheetFontSizeApiCalls(cells, range, sheetId, direction);
    if (!calls.length) return false;
    workbook.batchCallApis(calls);
    return true;
  } catch {
    return false;
  }
}

function canAdjustSpreadsheetFontSizeCells(
  rowCount: number,
  columnCount: number,
  cellAt: (rowOffset: number, columnOffset: number) => Cell | null | undefined,
  direction: SpreadsheetFontSizeDirection,
): boolean {
  for (let rowOffset = 0; rowOffset < rowCount; rowOffset += 1) {
    for (let columnOffset = 0; columnOffset < columnCount; columnOffset += 1) {
      if (
        adjustedSpreadsheetFontSize(
          cellAt(rowOffset, columnOffset),
          direction,
        ) !== null
      ) {
        return true;
      }
    }
  }
  return false;
}

function liveSpreadsheetFontSizeRange(
  context: SpreadsheetCommandContext,
): SpreadsheetCellRange {
  return (
    normalizeSpreadsheetCellRange(
      spreadsheetSingleRange(
        context.workbook?.getSelection()?.at(-1) ?? context.fallbackRange,
      ),
    ) ?? { row: [0, 0], column: [0, 0] }
  );
}

function isSpreadsheetFontSizeDirection(
  value: unknown,
): value is SpreadsheetFontSizeDirection {
  return value === 'grow' || value === 'shrink';
}

function runSpreadsheetFontSizeShortcut(
  event: KeyboardEvent,
  canExecute: SpreadsheetEditorCommands['adjustFontSize'],
  execute: SpreadsheetEditorCommands['adjustFontSize'],
  direction: SpreadsheetFontSizeDirection,
): boolean {
  if (
    event.repeat ||
    isOfficeShortcutBlocked(event.target) ||
    isSpreadsheetNativeTextUndoTarget(event.target) ||
    !canExecute(direction)
  ) {
    return false;
  }
  return execute(direction);
}
