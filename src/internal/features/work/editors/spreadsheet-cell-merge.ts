import type { Cell } from '@fortune-sheet/core';
import type { WorkSpreadsheetSheet } from '../work-types';

export type SpreadsheetCellMergeCommand =
  | 'merge-and-center'
  | 'merge-cells'
  | 'merge-across'
  | 'unmerge-cells'
  | 'unmerge-and-fill';

export interface SpreadsheetCellMergeRange {
  row: number[];
  column: number[];
}

export interface SpreadsheetCellMergeApiCall {
  name: string;
  args: unknown[];
}

interface SpreadsheetNativeMerge {
  r: number;
  c: number;
  rs: number;
  cs: number;
}

export function canApplySpreadsheetCellMerge(
  sheet: WorkSpreadsheetSheet | undefined,
  range: SpreadsheetCellMergeRange,
  command: SpreadsheetCellMergeCommand,
): boolean {
  if (!sheet) return false;
  const normalized = normalizeSpreadsheetCellMergeRange(range);
  if (!normalized) return false;
  const merged = spreadsheetMergesIntersectingRange(sheet, normalized);
  if (command === 'unmerge-cells' || command === 'unmerge-and-fill') {
    return merged.length > 0;
  }
  if (merged.length > 0) return false;
  const rowCount = normalized.row[1] - normalized.row[0] + 1;
  const columnCount = normalized.column[1] - normalized.column[0] + 1;
  if (command === 'merge-across') return columnCount > 1;
  return rowCount * columnCount > 1;
}

export function spreadsheetCellMergeApiCalls(
  sheet: WorkSpreadsheetSheet | undefined,
  sheetId: string,
  range: SpreadsheetCellMergeRange,
  command: SpreadsheetCellMergeCommand,
): SpreadsheetCellMergeApiCall[] {
  if (!canApplySpreadsheetCellMerge(sheet, range, command) || !sheet) {
    return [];
  }
  const normalized = normalizeSpreadsheetCellMergeRange(range);
  if (!normalized) return [];
  const options = { id: sheetId };

  if (command === 'merge-and-center') {
    return [
      {
        name: 'mergeCells',
        args: [[normalized], 'merge-all', options],
      },
      {
        name: 'setCellFormatByRange',
        args: ['ht', '0', normalized, options],
      },
    ];
  }
  if (command === 'merge-cells' || command === 'merge-across') {
    return [
      {
        name: 'mergeCells',
        args: [
          [normalized],
          command === 'merge-across' ? 'merge-horizontal' : 'merge-all',
          options,
        ],
      },
    ];
  }

  const merges = spreadsheetMergesIntersectingRange(sheet, normalized);
  const mergeRanges = merges.map(spreadsheetNativeMergeRange);
  const calls: SpreadsheetCellMergeApiCall[] = [
    {
      name: 'cancelMerge',
      args: [mergeRanges, options],
    },
  ];
  if (command === 'unmerge-cells') return calls;

  for (const merge of merges) {
    const anchor = sheet.data?.[merge.r]?.[merge.c];
    if (!anchor) continue;
    const targetRange = spreadsheetNativeMergeRange(merge);
    calls.push({
      name: 'setCellValuesByRange',
      args: [
        Array.from({ length: merge.rs }, () =>
          Array.from({ length: merge.cs }, () =>
            cloneSpreadsheetCellWithoutMerge(anchor),
          ),
        ),
        targetRange,
        null,
        options,
      ],
    });
  }
  return calls;
}

function normalizeSpreadsheetCellMergeRange(
  range: SpreadsheetCellMergeRange,
): { row: [number, number]; column: [number, number] } | null {
  const row = spreadsheetCellMergeBoundary(range.row);
  const column = spreadsheetCellMergeBoundary(range.column);
  return row && column ? { row, column } : null;
}

function spreadsheetCellMergeBoundary(
  values: readonly number[],
): [number, number] | null {
  if (
    !values.length ||
    values.some((value) => !Number.isInteger(value) || value < 0)
  ) {
    return null;
  }
  return [Math.min(...values), Math.max(...values)];
}

function spreadsheetMergesIntersectingRange(
  sheet: WorkSpreadsheetSheet,
  range: { row: [number, number]; column: [number, number] },
): SpreadsheetNativeMerge[] {
  return Object.values(sheet.config?.merge ?? {})
    .filter(isSpreadsheetNativeMerge)
    .filter(
      (merge) =>
        merge.r <= range.row[1] &&
        merge.r + merge.rs - 1 >= range.row[0] &&
        merge.c <= range.column[1] &&
        merge.c + merge.cs - 1 >= range.column[0],
    )
    .sort((left, right) => left.r - right.r || left.c - right.c);
}

function isSpreadsheetNativeMerge(
  value: unknown,
): value is SpreadsheetNativeMerge {
  if (!value || typeof value !== 'object') return false;
  const merge = value as Partial<SpreadsheetNativeMerge>;
  return (
    Number.isInteger(merge.r) &&
    Number.isInteger(merge.c) &&
    Number.isInteger(merge.rs) &&
    Number.isInteger(merge.cs) &&
    Number(merge.r) >= 0 &&
    Number(merge.c) >= 0 &&
    Number(merge.rs) > 0 &&
    Number(merge.cs) > 0
  );
}

function spreadsheetNativeMergeRange(merge: SpreadsheetNativeMerge): {
  row: [number, number];
  column: [number, number];
} {
  return {
    row: [merge.r, merge.r + merge.rs - 1],
    column: [merge.c, merge.c + merge.cs - 1],
  };
}

function cloneSpreadsheetCellWithoutMerge(cell: Cell): Cell {
  const clone = structuredClone(cell);
  delete clone.mc;
  return clone;
}
