import type { Cell } from '@fortune-sheet/core';
import type { WorkSpreadsheetSheet } from '../work-types';
import {
  parseSpreadsheetCellRange,
  type SpreadsheetCellRange,
  spreadsheetCellRangesIntersect,
} from './spreadsheet-cell-range';

export function canMutateSpreadsheetCellRange(
  sheet: WorkSpreadsheetSheet | undefined,
  range: SpreadsheetCellRange,
): boolean {
  return canMutateSpreadsheetCellRanges(sheet, [range]);
}

export function canMutateSpreadsheetCellRanges(
  sheet: WorkSpreadsheetSheet | undefined,
  ranges: readonly SpreadsheetCellRange[],
): boolean {
  if (
    !sheet ||
    ranges.length === 0 ||
    ranges.some((range) => !spreadsheetCellMutationRangeIsValid(range))
  ) {
    return false;
  }
  if (sheet.isPivotTable || sheet.pivotTable || sheet.pivotTables?.length) {
    return false;
  }

  const merges = Object.values(sheet.config?.merge ?? {});
  if (
    merges.some((merge) => {
      const mergeRange = spreadsheetNativeMergeRange(merge);
      return (
        mergeRange &&
        ranges.some((range) =>
          spreadsheetCellRangesIntersect(mergeRange, range),
        )
      );
    })
  ) {
    return false;
  }

  const cellAt = createSpreadsheetMutationCellReader(sheet);
  for (const range of ranges) {
    for (let row = range.row[0]; row <= range.row[1]; row += 1) {
      if (sheet.config?.rowReadOnly?.[row]) return false;
    }
    for (let column = range.column[0]; column <= range.column[1]; column += 1) {
      if (sheet.config?.colReadOnly?.[column]) return false;
    }
    for (let row = range.row[0]; row <= range.row[1]; row += 1) {
      for (
        let column = range.column[0];
        column <= range.column[1];
        column += 1
      ) {
        if (spreadsheetCellIsLocked(sheet, cellAt, row, column)) return false;
      }
    }
  }
  return true;
}

function spreadsheetCellMutationRangeIsValid(
  range: SpreadsheetCellRange,
): boolean {
  return (
    Number.isSafeInteger(range.row[0]) &&
    Number.isSafeInteger(range.row[1]) &&
    Number.isSafeInteger(range.column[0]) &&
    Number.isSafeInteger(range.column[1]) &&
    range.row[0] >= 0 &&
    range.row[0] <= range.row[1] &&
    range.column[0] >= 0 &&
    range.column[0] <= range.column[1]
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
  cellAt: (row: number, column: number) => Cell | null | undefined,
  row: number,
  column: number,
): boolean {
  const cell = cellAt(row, column);
  if (cell?.lo !== undefined && cell.lo !== null) return Boolean(cell.lo);

  const authority = sheet.config?.authority;
  if (!isRecord(authority)) return false;
  const ranges = authority.cellProtectionRanges;
  if (Array.isArray(ranges)) {
    for (let index = ranges.length - 1; index >= 0; index -= 1) {
      const candidate = ranges[index];
      if (!isRecord(candidate)) continue;
      const protectedRange = parseSpreadsheetCellRange(candidate.range);
      if (
        protectedRange &&
        row >= protectedRange.row[0] &&
        row <= protectedRange.row[1] &&
        column >= protectedRange.column[0] &&
        column <= protectedRange.column[1]
      ) {
        return candidate.locked !== false;
      }
    }
  }
  return authority.sheet !== undefined && authority.sheet !== 0;
}

function createSpreadsheetMutationCellReader(
  sheet: WorkSpreadsheetSheet,
): (row: number, column: number) => Cell | null | undefined {
  let sparseRows: Map<number, Map<number, Cell | null>> | null = null;
  const sparseCell = (row: number, column: number) => {
    if (!sparseRows) {
      sparseRows = new Map();
      for (const entry of sheet.celldata ?? []) {
        let sparseRow = sparseRows.get(entry.r);
        if (!sparseRow) {
          sparseRow = new Map();
          sparseRows.set(entry.r, sparseRow);
        }
        sparseRow.set(entry.c, entry.v);
      }
    }
    return sparseRows.get(row)?.get(column);
  };
  return (row, column) =>
    sheet.data?.[row]?.[column] ?? sparseCell(row, column);
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
