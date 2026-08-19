import type { Cell } from '@fortune-sheet/core';
import type { WorkSpreadsheetSheet } from '../work-types';
import {
  normalizeSpreadsheetCellRange,
  parseSpreadsheetCellRange,
  type SpreadsheetCellRange,
  type SpreadsheetCellRangeInput,
  spreadsheetCellRangeArea,
  spreadsheetCellRangesIntersect,
} from './spreadsheet-cell-range';

export const spreadsheetCellFillMaximumCells = 50_000;

export type SpreadsheetCellFillDirection = 'down' | 'left' | 'right' | 'up';

export interface SpreadsheetCellFillPlan {
  applyRange: SpreadsheetCellRange;
  copyRange: SpreadsheetCellRange;
  direction: SpreadsheetCellFillDirection;
  targetCellCount: number;
}

export function planSpreadsheetCellFill(
  selection: SpreadsheetCellRangeInput,
  direction: SpreadsheetCellFillDirection,
  maximumCells = spreadsheetCellFillMaximumCells,
): SpreadsheetCellFillPlan | null {
  const range = normalizeSpreadsheetCellRange(selection);
  if (
    !range ||
    !spreadsheetCellFillRangeIsSafe(range) ||
    !Number.isSafeInteger(maximumCells) ||
    maximumCells < 1
  ) {
    return null;
  }

  const [top, bottom] = range.row;
  const [left, right] = range.column;
  let copyRange: SpreadsheetCellRange;
  let applyRange: SpreadsheetCellRange;

  switch (direction) {
    case 'down':
      if (top === bottom) return null;
      copyRange = { row: [top, top], column: [left, right] };
      applyRange = { row: [top + 1, bottom], column: [left, right] };
      break;
    case 'up':
      if (top === bottom) return null;
      copyRange = { row: [bottom, bottom], column: [left, right] };
      applyRange = { row: [top, bottom - 1], column: [left, right] };
      break;
    case 'right':
      if (left === right) return null;
      copyRange = { row: [top, bottom], column: [left, left] };
      applyRange = { row: [top, bottom], column: [left + 1, right] };
      break;
    case 'left':
      if (left === right) return null;
      copyRange = { row: [top, bottom], column: [right, right] };
      applyRange = { row: [top, bottom], column: [left, right - 1] };
      break;
  }

  const targetCellCount = spreadsheetCellRangeArea(applyRange);
  if (
    !Number.isSafeInteger(targetCellCount) ||
    targetCellCount > maximumCells
  ) {
    return null;
  }
  return { applyRange, copyRange, direction, targetCellCount };
}

export function canApplySpreadsheetCellFill(
  sheet: WorkSpreadsheetSheet | undefined,
  plan: SpreadsheetCellFillPlan,
): boolean {
  if (!sheet || !spreadsheetCellFillPlanIsValid(plan)) return false;
  if (sheet.isPivotTable || sheet.pivotTable || sheet.pivotTables?.length) {
    return false;
  }

  const selection = spreadsheetCellFillSelection(plan);
  const merges = Object.values(sheet.config?.merge ?? {});
  if (
    merges.some((merge) => {
      const range = spreadsheetNativeMergeRange(merge);
      return range && spreadsheetCellRangesIntersect(range, selection);
    })
  ) {
    return false;
  }

  for (let row = selection.row[0]; row <= selection.row[1]; row += 1) {
    if (sheet.config?.rowReadOnly?.[row]) return false;
  }
  for (
    let column = selection.column[0];
    column <= selection.column[1];
    column += 1
  ) {
    if (sheet.config?.colReadOnly?.[column]) return false;
  }
  for (let row = selection.row[0]; row <= selection.row[1]; row += 1) {
    for (
      let column = selection.column[0];
      column <= selection.column[1];
      column += 1
    ) {
      if (spreadsheetCellIsLocked(sheet, row, column)) return false;
    }
  }
  return true;
}

function spreadsheetCellFillPlanIsValid(
  plan: SpreadsheetCellFillPlan,
): boolean {
  const selection = spreadsheetCellFillSelection(plan);
  const expected = planSpreadsheetCellFill(selection, plan.direction);
  return Boolean(
    expected &&
      sameSpreadsheetCellRange(expected.copyRange, plan.copyRange) &&
      sameSpreadsheetCellRange(expected.applyRange, plan.applyRange) &&
      expected.targetCellCount === plan.targetCellCount,
  );
}

function spreadsheetCellFillSelection(
  plan: SpreadsheetCellFillPlan,
): SpreadsheetCellRange {
  return {
    row: [
      Math.min(plan.copyRange.row[0], plan.applyRange.row[0]),
      Math.max(plan.copyRange.row[1], plan.applyRange.row[1]),
    ],
    column: [
      Math.min(plan.copyRange.column[0], plan.applyRange.column[0]),
      Math.max(plan.copyRange.column[1], plan.applyRange.column[1]),
    ],
  };
}

function spreadsheetCellFillRangeIsSafe(range: SpreadsheetCellRange): boolean {
  return [...range.row, ...range.column].every(Number.isSafeInteger);
}

function sameSpreadsheetCellRange(
  left: SpreadsheetCellRange,
  right: SpreadsheetCellRange,
): boolean {
  return (
    left.row[0] === right.row[0] &&
    left.row[1] === right.row[1] &&
    left.column[0] === right.column[0] &&
    left.column[1] === right.column[1]
  );
}

function spreadsheetNativeMergeRange(
  value: unknown,
): SpreadsheetCellRange | null {
  if (!isRecord(value)) return null;
  const row = finiteIndex(value.r);
  const column = finiteIndex(value.c);
  const rowSpan = positiveSpan(value.rs);
  const columnSpan = positiveSpan(value.cs);
  if (
    row === null ||
    column === null ||
    rowSpan === null ||
    columnSpan === null ||
    !Number.isSafeInteger(row + rowSpan - 1) ||
    !Number.isSafeInteger(column + columnSpan - 1)
  ) {
    return null;
  }
  return {
    row: [row, row + rowSpan - 1],
    column: [column, column + columnSpan - 1],
  };
}

function spreadsheetCellIsLocked(
  sheet: WorkSpreadsheetSheet,
  row: number,
  column: number,
): boolean {
  const cell = sheet.data?.[row]?.[column] as Cell | null | undefined;
  if (cell?.lo !== undefined && cell.lo !== null) return Boolean(cell.lo);

  const authority = sheet.config?.authority;
  if (!isRecord(authority)) return false;
  const ranges = authority.cellProtectionRanges;
  if (Array.isArray(ranges)) {
    for (let index = ranges.length - 1; index >= 0; index -= 1) {
      const candidate = ranges[index];
      if (!isRecord(candidate)) continue;
      const range = parseSpreadsheetCellRange(candidate.range);
      if (
        range &&
        row >= range.row[0] &&
        row <= range.row[1] &&
        column >= range.column[0] &&
        column <= range.column[1]
      ) {
        return candidate.locked !== false;
      }
    }
  }
  return authority.sheet !== undefined && authority.sheet !== 0;
}

function finiteIndex(value: unknown): number | null {
  return Number.isSafeInteger(value) && Number(value) >= 0
    ? Number(value)
    : null;
}

function positiveSpan(value: unknown): number | null {
  return Number.isSafeInteger(value) && Number(value) > 0
    ? Number(value)
    : null;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value && typeof value === 'object' && !Array.isArray(value));
}
