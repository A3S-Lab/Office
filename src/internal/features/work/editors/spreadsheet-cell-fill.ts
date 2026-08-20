import type { WorkSpreadsheetSheet } from '../work-types';
import {
  normalizeSpreadsheetCellRange,
  type SpreadsheetCellRange,
  type SpreadsheetCellRangeInput,
  spreadsheetCellRangeArea,
} from './spreadsheet-cell-range';
import { canMutateSpreadsheetCellRange } from './spreadsheet-cell-mutation-guard';

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
  const selection = spreadsheetCellFillSelection(plan);
  return canMutateSpreadsheetCellRange(sheet, selection);
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
