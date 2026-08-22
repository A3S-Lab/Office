import type { Cell } from '@fortune-sheet/core';
import { xlsxNativeFillCellKeys } from '../work-xlsx-native-fill';

export const spreadsheetCellFormatAttributes = [
  'bl',
  'it',
  'ff',
  'fs',
  'fc',
  'ht',
  'vt',
  'tb',
  'cl',
  'un',
  'tr',
  'ct',
  'bg',
  ...xlsxNativeFillCellKeys,
] as const;

export type SpreadsheetCellFormatAttribute =
  (typeof spreadsheetCellFormatAttributes)[number];

export type SpreadsheetFormatPainterMode = 'once' | 'locked';

export type SpreadsheetCellFormatSnapshot = Record<
  SpreadsheetCellFormatAttribute,
  unknown
>;

export interface SpreadsheetFormatPattern {
  cells: readonly (readonly SpreadsheetCellFormatSnapshot[])[];
  columnCount: number;
  rowCount: number;
}

export interface SpreadsheetFormatPainterRange {
  column: number[];
  row: number[];
}

export interface SpreadsheetFormatPainterSheetBounds {
  columnCount: number;
  rowCount: number;
}

export interface SpreadsheetFormatPainterBatch {
  attribute: SpreadsheetCellFormatAttribute;
  ranges: SpreadsheetFormatPainterRange[];
  value: unknown;
}

export function captureSpreadsheetFormatPattern(
  cells: readonly (readonly (Cell | null)[])[],
): SpreadsheetFormatPattern | null {
  const rowCount = cells.length;
  const columnCount = cells[0]?.length ?? 0;
  if (
    rowCount < 1 ||
    columnCount < 1 ||
    cells.some((row) => row.length !== columnCount)
  ) {
    return null;
  }

  return {
    cells: cells.map((row) =>
      row.map((cell) => spreadsheetCellFormatSnapshot(cell)),
    ),
    columnCount,
    rowCount,
  };
}

export function spreadsheetFormatPainterTargetRange(
  target: SpreadsheetFormatPainterRange,
  pattern: Pick<SpreadsheetFormatPattern, 'columnCount' | 'rowCount'>,
  bounds: SpreadsheetFormatPainterSheetBounds,
): SpreadsheetFormatPainterRange | null {
  if (
    pattern.rowCount < 1 ||
    pattern.columnCount < 1 ||
    bounds.rowCount < 1 ||
    bounds.columnCount < 1
  ) {
    return null;
  }
  const normalized = normalizeSpreadsheetFormatPainterRange(target);
  if (!normalized) return null;
  const singleCell =
    normalized.row[0] === normalized.row[1] &&
    normalized.column[0] === normalized.column[1];
  const startRow = clampIndex(normalized.row[0], bounds.rowCount);
  const startColumn = clampIndex(normalized.column[0], bounds.columnCount);
  const requestedEndRow = singleCell
    ? startRow + pattern.rowCount - 1
    : normalized.row[1];
  const requestedEndColumn = singleCell
    ? startColumn + pattern.columnCount - 1
    : normalized.column[1];

  return {
    row: [startRow, clampIndex(requestedEndRow, bounds.rowCount)],
    column: [startColumn, clampIndex(requestedEndColumn, bounds.columnCount)],
  };
}

export function spreadsheetFormatPainterCellCount(
  range: SpreadsheetFormatPainterRange,
): number {
  const normalized = normalizeSpreadsheetFormatPainterRange(range);
  if (!normalized) return 0;
  return (
    (normalized.row[1] - normalized.row[0] + 1) *
    (normalized.column[1] - normalized.column[0] + 1)
  );
}

export function spreadsheetFormatPainterBatches(
  pattern: SpreadsheetFormatPattern,
  target: SpreadsheetFormatPainterRange,
): SpreadsheetFormatPainterBatch[] {
  const normalized = normalizeSpreadsheetFormatPainterRange(target);
  if (!normalized || pattern.rowCount < 1 || pattern.columnCount < 1) {
    return [];
  }
  const batches: SpreadsheetFormatPainterBatch[] = [];
  for (const attribute of spreadsheetCellFormatAttributes) {
    const rectangles = spreadsheetFormatRectangles(
      pattern,
      normalized,
      attribute,
    );
    const groups = new Map<
      string,
      { ranges: SpreadsheetFormatPainterRange[]; value: unknown }
    >();
    for (const rectangle of rectangles) {
      const key = spreadsheetFormatValueKey(rectangle.value);
      const group = groups.get(key);
      if (group) {
        group.ranges.push(rectangle.range);
      } else {
        groups.set(key, {
          ranges: [rectangle.range],
          value: cloneSpreadsheetFormatValue(rectangle.value),
        });
      }
    }
    for (const group of groups.values()) {
      batches.push({
        attribute,
        ranges: group.ranges,
        value: group.value,
      });
    }
  }
  return batches;
}

function spreadsheetCellFormatSnapshot(
  cell: Cell | null,
): SpreadsheetCellFormatSnapshot {
  const snapshot = {} as SpreadsheetCellFormatSnapshot;
  const source = cell as (Cell & Record<string, unknown>) | null;
  for (const attribute of spreadsheetCellFormatAttributes) {
    snapshot[attribute] =
      attribute === 'ct'
        ? spreadsheetNumberFormatSnapshot(cell?.ct)
        : cloneSpreadsheetFormatValue(source?.[attribute]);
  }
  return snapshot;
}

function spreadsheetNumberFormatSnapshot(
  value: Cell['ct'] | undefined,
): NonNullable<Cell['ct']> {
  const format = value?.fa?.trim();
  const type = value?.t?.trim();
  return {
    fa: format || 'General',
    t: type || 'g',
  };
}

function spreadsheetFormatRectangles(
  pattern: SpreadsheetFormatPattern,
  target: SpreadsheetFormatPainterRange,
  attribute: SpreadsheetCellFormatAttribute,
): Array<{ range: SpreadsheetFormatPainterRange; value: unknown }> {
  const rectangles: Array<{
    range: SpreadsheetFormatPainterRange;
    value: unknown;
  }> = [];
  let previous = new Map<
    string,
    { range: SpreadsheetFormatPainterRange; value: unknown }
  >();
  for (let row = target.row[0]; row <= target.row[1]; row += 1) {
    const runs: Array<{
      endColumn: number;
      key: string;
      startColumn: number;
      value: unknown;
    }> = [];
    let startColumn = target.column[0];
    let value = spreadsheetPatternValue(
      pattern,
      target,
      row,
      startColumn,
      attribute,
    );
    let key = spreadsheetFormatValueKey(value);
    for (
      let column = target.column[0] + 1;
      column <= target.column[1] + 1;
      column += 1
    ) {
      const nextValue =
        column <= target.column[1]
          ? spreadsheetPatternValue(pattern, target, row, column, attribute)
          : undefined;
      const nextKey =
        column <= target.column[1]
          ? spreadsheetFormatValueKey(nextValue)
          : '__end__';
      if (nextKey === key) continue;
      runs.push({
        endColumn: column - 1,
        key,
        startColumn,
        value,
      });
      startColumn = column;
      value = nextValue;
      key = nextKey;
    }

    const current = new Map<
      string,
      { range: SpreadsheetFormatPainterRange; value: unknown }
    >();
    for (const run of runs) {
      const signature = `${run.startColumn}:${run.endColumn}:${run.key}`;
      const continuing = previous.get(signature);
      if (continuing) {
        continuing.range.row[1] = row;
        current.set(signature, continuing);
        continue;
      }
      const rectangle = {
        range: {
          row: [row, row],
          column: [run.startColumn, run.endColumn],
        },
        value: run.value,
      };
      rectangles.push(rectangle);
      current.set(signature, rectangle);
    }
    previous = current;
  }
  return rectangles;
}

function spreadsheetPatternValue(
  pattern: SpreadsheetFormatPattern,
  target: SpreadsheetFormatPainterRange,
  row: number,
  column: number,
  attribute: SpreadsheetCellFormatAttribute,
): unknown {
  const sourceRow = (row - target.row[0]) % pattern.rowCount;
  const sourceColumn = (column - target.column[0]) % pattern.columnCount;
  return pattern.cells[sourceRow]?.[sourceColumn]?.[attribute];
}

function normalizeSpreadsheetFormatPainterRange(
  range: SpreadsheetFormatPainterRange,
): SpreadsheetFormatPainterRange | null {
  const values = [...range.row, ...range.column];
  if (values.length !== 4 || values.some((value) => !Number.isFinite(value))) {
    return null;
  }
  return {
    row: [
      Math.max(0, Math.floor(Math.min(range.row[0], range.row[1]))),
      Math.max(0, Math.floor(Math.max(range.row[0], range.row[1]))),
    ],
    column: [
      Math.max(0, Math.floor(Math.min(range.column[0], range.column[1]))),
      Math.max(0, Math.floor(Math.max(range.column[0], range.column[1]))),
    ],
  };
}

function clampIndex(value: number, length: number): number {
  return Math.max(0, Math.min(Math.floor(value), Math.floor(length) - 1));
}

function spreadsheetFormatValueKey(value: unknown): string {
  if (value === undefined) return 'undefined';
  if (typeof value === 'number' && Number.isNaN(value)) return 'number:NaN';
  return `${typeof value}:${JSON.stringify(value)}`;
}

function cloneSpreadsheetFormatValue(value: unknown): unknown {
  if (!value || typeof value !== 'object') return value;
  if (Array.isArray(value)) {
    return value.map((item) => cloneSpreadsheetFormatValue(item));
  }
  return Object.fromEntries(
    Object.entries(value).map(([key, item]) => [
      key,
      cloneSpreadsheetFormatValue(item),
    ]),
  );
}
