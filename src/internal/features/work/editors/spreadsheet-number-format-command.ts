import type { Cell } from '@fortune-sheet/core';
import type { WorkSpreadsheetContent } from '../work-types';
import {
  normalizeSpreadsheetCellRange,
  type SpreadsheetCellRange,
  type SpreadsheetCellRangeInput,
  spreadsheetCellRangeArea,
} from './spreadsheet-cell-range';
import type {
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import { spreadsheetSingleRange } from './spreadsheet-editor-support';
import {
  adjustSpreadsheetNumberFormat,
  spreadsheetNumberFormatValue,
} from './spreadsheet-number-format';
import {
  createOfficeEditorExtension,
  type OfficeEditorExtension,
} from './office-editor-extension';

export type SpreadsheetDecimalPlacesDirection = 'decrease' | 'increase';

export const MAX_SPREADSHEET_DECIMAL_FORMAT_CELLS = 10_000;

export interface SpreadsheetDecimalFormatApiCall {
  name: 'setCellFormatByRange';
  args: [
    'ct',
    NonNullable<Cell['ct']> & { fa: string; t: string },
    SpreadsheetCellRange,
    { id: string },
  ];
}

interface SpreadsheetDecimalFormatRun {
  format: NonNullable<Cell['ct']> & { fa: string; t: string };
  key: string;
  range: SpreadsheetCellRange;
}

export function createSpreadsheetNumberFormatExtension(): OfficeEditorExtension<
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands
> {
  return createOfficeEditorExtension<
    SpreadsheetCommandContext,
    SpreadsheetEditorCommands
  >({
    name: 'spreadsheetNumberFormats',
    addCommands: () => ({
      adjustDecimalPlaces: {
        canExecute: canAdjustSelectedDecimalPlaces,
        execute: adjustSelectedDecimalPlaces,
      },
    }),
  });
}

export function canAdjustSpreadsheetDecimalPlaces(
  content: WorkSpreadsheetContent,
  sheetId: string,
  range: SpreadsheetCellRangeInput,
  direction: SpreadsheetDecimalPlacesDirection,
): boolean {
  const normalizedRange = normalizeSpreadsheetCellRange(range);
  const sheet = content.sheets.find((candidate) => candidate.id === sheetId);
  if (
    !sheet ||
    !normalizedRange ||
    !isSpreadsheetDecimalPlacesDirection(direction) ||
    spreadsheetCellRangeArea(normalizedRange) >
      MAX_SPREADSHEET_DECIMAL_FORMAT_CELLS
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
  return canAdjustSpreadsheetDecimalCells(
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

export function spreadsheetDecimalFormatApiCalls(
  cells: readonly (readonly (Cell | null)[])[],
  range: SpreadsheetCellRangeInput,
  sheetId: string,
  direction: SpreadsheetDecimalPlacesDirection,
): SpreadsheetDecimalFormatApiCall[] {
  const normalizedRange = normalizeSpreadsheetCellRange(range);
  if (
    !normalizedRange ||
    !sheetId ||
    !isSpreadsheetDecimalPlacesDirection(direction) ||
    spreadsheetCellRangeArea(normalizedRange) >
      MAX_SPREADSHEET_DECIMAL_FORMAT_CELLS
  ) {
    return [];
  }

  const rectangles: SpreadsheetDecimalFormatRun[] = [];
  let active = new Map<string, SpreadsheetDecimalFormatRun>();
  const rowCount = normalizedRange.row[1] - normalizedRange.row[0] + 1;
  const columnCount = normalizedRange.column[1] - normalizedRange.column[0] + 1;

  for (let rowOffset = 0; rowOffset < rowCount; rowOffset += 1) {
    const row = normalizedRange.row[0] + rowOffset;
    const runs = spreadsheetDecimalFormatRowRuns(
      cells[rowOffset],
      row,
      normalizedRange.column[0],
      columnCount,
      direction,
    );
    const nextActive = new Map<string, SpreadsheetDecimalFormatRun>();
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

  return rectangles.map(({ format, range: targetRange }) => ({
    name: 'setCellFormatByRange',
    args: ['ct', format, targetRange, { id: sheetId }],
  }));
}

function spreadsheetDecimalFormatRowRuns(
  cells: readonly (Cell | null)[] | undefined,
  row: number,
  firstColumn: number,
  columnCount: number,
  direction: SpreadsheetDecimalPlacesDirection,
): SpreadsheetDecimalFormatRun[] {
  const runs: SpreadsheetDecimalFormatRun[] = [];
  let current: SpreadsheetDecimalFormatRun | null = null;
  for (let columnOffset = 0; columnOffset < columnCount; columnOffset += 1) {
    const column = firstColumn + columnOffset;
    const format = adjustedSpreadsheetDecimalFormat(
      cells?.[columnOffset],
      direction,
    );
    if (!format) {
      if (current) runs.push(current);
      current = null;
      continue;
    }
    const key = `${format.t}\u0000${format.fa}`;
    if (current && key === current.key) {
      current.range.column[1] = column;
      continue;
    }
    if (current) runs.push(current);
    current = {
      format,
      key,
      range: { row: [row, row], column: [column, column] },
    };
  }
  if (current) runs.push(current);
  return runs;
}

function adjustedSpreadsheetDecimalFormat(
  cell: Cell | null | undefined,
  direction: SpreadsheetDecimalPlacesDirection,
): (NonNullable<Cell['ct']> & { fa: string; t: string }) | null {
  const current = cell?.ct?.fa?.trim() || 'General';
  const adjusted = adjustSpreadsheetNumberFormat(
    current,
    direction === 'increase' ? 1 : -1,
  );
  return adjusted === current
    ? null
    : spreadsheetNumberFormatValue(adjusted, cell);
}

function isSpreadsheetDecimalPlacesDirection(
  value: unknown,
): value is SpreadsheetDecimalPlacesDirection {
  return value === 'decrease' || value === 'increase';
}

function canAdjustSelectedDecimalPlaces(
  context: SpreadsheetCommandContext,
  direction: SpreadsheetDecimalPlacesDirection,
): boolean {
  const workbook = context.workbook;
  const sheetId = context.targetSheetId;
  if (
    !context.editable ||
    !workbook ||
    !sheetId ||
    sheetId !== context.activeSheetId ||
    !isSpreadsheetDecimalPlacesDirection(direction)
  ) {
    return false;
  }
  const range = liveSpreadsheetDecimalFormatRange(context);
  if (spreadsheetCellRangeArea(range) > MAX_SPREADSHEET_DECIMAL_FORMAT_CELLS) {
    return false;
  }
  try {
    const cells = workbook.getCellsByRange(range, { id: sheetId });
    return canAdjustSpreadsheetDecimalCells(
      range.row[1] - range.row[0] + 1,
      range.column[1] - range.column[0] + 1,
      (rowOffset, columnOffset) => cells[rowOffset]?.[columnOffset] ?? null,
      direction,
    );
  } catch {
    return false;
  }
}

function adjustSelectedDecimalPlaces(
  context: SpreadsheetCommandContext,
  direction: SpreadsheetDecimalPlacesDirection,
): boolean {
  const workbook = context.workbook;
  const sheetId = context.targetSheetId;
  if (
    !context.editable ||
    !workbook ||
    !sheetId ||
    sheetId !== context.activeSheetId ||
    !isSpreadsheetDecimalPlacesDirection(direction)
  ) {
    return false;
  }
  const range = liveSpreadsheetDecimalFormatRange(context);
  if (spreadsheetCellRangeArea(range) > MAX_SPREADSHEET_DECIMAL_FORMAT_CELLS) {
    return false;
  }
  try {
    const cells = workbook.getCellsByRange(range, { id: sheetId });
    const calls = spreadsheetDecimalFormatApiCalls(
      cells,
      range,
      sheetId,
      direction,
    );
    if (!calls.length) return false;
    workbook.batchCallApis(calls);
    return true;
  } catch {
    return false;
  }
}

function canAdjustSpreadsheetDecimalCells(
  rowCount: number,
  columnCount: number,
  cellAt: (rowOffset: number, columnOffset: number) => Cell | null | undefined,
  direction: SpreadsheetDecimalPlacesDirection,
): boolean {
  for (let rowOffset = 0; rowOffset < rowCount; rowOffset += 1) {
    for (let columnOffset = 0; columnOffset < columnCount; columnOffset += 1) {
      if (
        adjustedSpreadsheetDecimalFormat(
          cellAt(rowOffset, columnOffset),
          direction,
        )
      ) {
        return true;
      }
    }
  }
  return false;
}

function liveSpreadsheetDecimalFormatRange(
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
