export interface SpreadsheetCellRange {
  row: [number, number];
  column: [number, number];
}

export interface SpreadsheetCellRangeInput {
  row: readonly number[];
  column: readonly number[];
}

type UnknownRecord = Record<string, unknown>;

export function normalizeSpreadsheetCellRange(
  value: SpreadsheetCellRangeInput,
): SpreadsheetCellRange | null {
  const row = normalizeSpreadsheetCellRangeAxis(value.row);
  const column = normalizeSpreadsheetCellRangeAxis(value.column);
  return row && column ? { row, column } : null;
}

export function parseSpreadsheetCellRange(
  value: unknown,
): SpreadsheetCellRange | null {
  if (!isRecord(value)) return null;
  const row = normalizeSpreadsheetCellRangeAxis(value.row);
  const column = normalizeSpreadsheetCellRangeAxis(value.column);
  return row && column ? { row, column } : null;
}

export function subtractSpreadsheetCellRange(
  source: SpreadsheetCellRange,
  removed: SpreadsheetCellRange,
): SpreadsheetCellRange[] {
  const top = Math.max(source.row[0], removed.row[0]);
  const bottom = Math.min(source.row[1], removed.row[1]);
  const left = Math.max(source.column[0], removed.column[0]);
  const right = Math.min(source.column[1], removed.column[1]);
  if (top > bottom || left > right) return [source];

  const remaining: SpreadsheetCellRange[] = [];
  if (source.row[0] < top) {
    remaining.push({
      row: [source.row[0], top - 1],
      column: [...source.column],
    });
  }
  if (bottom < source.row[1]) {
    remaining.push({
      row: [bottom + 1, source.row[1]],
      column: [...source.column],
    });
  }
  if (source.column[0] < left) {
    remaining.push({
      row: [top, bottom],
      column: [source.column[0], left - 1],
    });
  }
  if (right < source.column[1]) {
    remaining.push({
      row: [top, bottom],
      column: [right + 1, source.column[1]],
    });
  }
  return remaining;
}

export function spreadsheetCellRangesIntersect(
  left: SpreadsheetCellRange,
  right: SpreadsheetCellRange,
): boolean {
  return !(
    left.row[1] < right.row[0] ||
    right.row[1] < left.row[0] ||
    left.column[1] < right.column[0] ||
    right.column[1] < left.column[0]
  );
}

export function spreadsheetCellRangeContains(
  range: SpreadsheetCellRange,
  row: number,
  column: number,
): boolean {
  return (
    row >= range.row[0] &&
    row <= range.row[1] &&
    column >= range.column[0] &&
    column <= range.column[1]
  );
}

export function spreadsheetCellRangeArea(range: SpreadsheetCellRange): number {
  return (
    (range.row[1] - range.row[0] + 1) * (range.column[1] - range.column[0] + 1)
  );
}

function normalizeSpreadsheetCellRangeAxis(
  value: unknown,
): [number, number] | null {
  if (!Array.isArray(value)) return null;
  const first = finiteSpreadsheetCellIndex(value[0]);
  const second = finiteSpreadsheetCellIndex(value[1]);
  if (first === null || second === null) return null;
  return [Math.min(first, second), Math.max(first, second)];
}

function finiteSpreadsheetCellIndex(value: unknown): number | null {
  return typeof value === 'number' && Number.isFinite(value) && value >= 0
    ? Math.trunc(value)
    : null;
}

function isRecord(value: unknown): value is UnknownRecord {
  return Boolean(value && typeof value === 'object' && !Array.isArray(value));
}
