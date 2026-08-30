import type { Cell } from '@fortune-sheet/core';
import { formatSpreadsheetCellRanges } from '../work-spreadsheet-ranges';
import {
  normalizeSpreadsheetCellRange,
  spreadsheetCellRangeArea,
  type SpreadsheetCellRange,
} from './spreadsheet-cell-range';
import { translateSpreadsheetFormula } from './spreadsheet-paste-special-cell';

export const MAX_SPREADSHEET_SORT_KEYS = 64;
export const MAX_SPREADSHEET_SORT_CELLS = 1_000_000;

const MAX_SPREADSHEET_SORT_HEADER_LENGTH = 48;

export type SpreadsheetSortDirection = 'ascending' | 'descending';

export interface SpreadsheetSortKey {
  column: number;
  direction: SpreadsheetSortDirection;
}

export interface SpreadsheetSortDialogValue {
  hasHeader: boolean;
  keys: SpreadsheetSortKey[];
}

export interface SpreadsheetSortTarget {
  activeColumn: number;
  range: SpreadsheetCellRange;
}

export interface SpreadsheetSortOpenRequest extends SpreadsheetSortTarget {
  sheetId: string;
}

export interface SpreadsheetSortRequest extends SpreadsheetSortDialogValue {
  range: SpreadsheetCellRange;
  sheetId: string;
}

export interface SpreadsheetSortDialogSource {
  columns: Array<{ column: number; label: string }>;
  range: SpreadsheetCellRange;
  rangeReference: string;
  sheetId: string;
  sheetName: string;
  value: SpreadsheetSortDialogValue;
}

export type SpreadsheetSortErrorCode =
  | 'column-out-of-range'
  | 'duplicate-key'
  | 'formula-reference-out-of-range'
  | 'invalid-direction'
  | 'invalid-matrix'
  | 'invalid-range'
  | 'missing-key'
  | 'not-enough-rows'
  | 'range-too-large'
  | 'too-many-keys'
  | 'unsupported-linked-cell';

export type SpreadsheetSortValidationResult =
  | { ok: true; request: SpreadsheetSortRequest }
  | { code: SpreadsheetSortErrorCode; message: string; ok: false };

export type SpreadsheetSortResult =
  | { ok: true; rows: (Cell | null)[][] }
  | { code: SpreadsheetSortErrorCode; message: string; ok: false };

export function createSpreadsheetSortDialogSource(
  sheetId: string,
  sheetName: string,
  target: SpreadsheetSortTarget,
  rows: readonly (readonly (Cell | null)[])[],
): SpreadsheetSortDialogSource | null {
  const range = normalizeSpreadsheetCellRange(target.range);
  if (!range || spreadsheetSortRangeError(range)) return null;
  const width = range.column[1] - range.column[0] + 1;
  const height = range.row[1] - range.row[0] + 1;
  if (rows.length !== height || rows.some((row) => row.length !== width)) {
    return null;
  }
  const hasHeader = height > 2 && spreadsheetSortRowsHaveHeader(rows);
  const firstColumn =
    target.activeColumn >= range.column[0] &&
    target.activeColumn <= range.column[1]
      ? target.activeColumn
      : range.column[0];
  const header = rows[0] ?? [];
  return {
    sheetId,
    sheetName,
    range,
    rangeReference: formatSpreadsheetCellRanges([range]),
    columns: Array.from({ length: width }, (_, offset) => {
      const column = range.column[0] + offset;
      const name = spreadsheetSortHeaderText(header[offset]);
      const coordinate = spreadsheetSortColumnLabel(column);
      return {
        column,
        label: name ? `${coordinate}（${name}）` : coordinate,
      };
    }),
    value: {
      hasHeader,
      keys: [{ column: firstColumn, direction: 'ascending' }],
    },
  };
}

export function validateSpreadsheetSortRequest(
  request: SpreadsheetSortRequest,
): SpreadsheetSortValidationResult {
  const range = normalizeSpreadsheetCellRange(request.range);
  if (!range || !request.sheetId.trim()) {
    return spreadsheetSortError('invalid-range');
  }
  const rangeError = spreadsheetSortRangeError(range);
  if (rangeError) return rangeError;
  if (!request.keys.length) return spreadsheetSortError('missing-key');
  if (request.keys.length > MAX_SPREADSHEET_SORT_KEYS) {
    return spreadsheetSortError('too-many-keys');
  }
  const seen = new Set<number>();
  const keys: SpreadsheetSortKey[] = [];
  for (const key of request.keys) {
    if (!Number.isSafeInteger(key.column)) {
      return spreadsheetSortError('column-out-of-range');
    }
    if (key.column < range.column[0] || key.column > range.column[1]) {
      return spreadsheetSortError('column-out-of-range');
    }
    if (seen.has(key.column)) return spreadsheetSortError('duplicate-key');
    if (key.direction !== 'ascending' && key.direction !== 'descending') {
      return spreadsheetSortError('invalid-direction');
    }
    seen.add(key.column);
    keys.push({ column: key.column, direction: key.direction });
  }
  const rowCount = range.row[1] - range.row[0] + 1;
  if (rowCount - (request.hasHeader ? 1 : 0) < 2) {
    return spreadsheetSortError('not-enough-rows');
  }
  return {
    ok: true,
    request: {
      sheetId: request.sheetId,
      range,
      hasHeader: Boolean(request.hasHeader),
      keys,
    },
  };
}

export function sortSpreadsheetRows(
  source: readonly (readonly (Cell | null)[])[],
  request: SpreadsheetSortRequest,
): SpreadsheetSortResult {
  const validation = validateSpreadsheetSortRequest(request);
  if (!validation.ok) return validation;
  const { range, hasHeader, keys } = validation.request;
  const width = range.column[1] - range.column[0] + 1;
  const height = range.row[1] - range.row[0] + 1;
  if (source.length !== height || source.some((row) => row.length !== width)) {
    return spreadsheetSortError('invalid-matrix');
  }
  if (source.some((row) => row.some((cell) => Boolean(cell?.hl)))) {
    return spreadsheetSortError('unsupported-linked-cell');
  }

  const header = hasHeader ? source[0] : undefined;
  const rows = source.slice(hasHeader ? 1 : 0);
  const sorted = rows
    .map((cells, index) => ({ cells, index }))
    .sort((left, right) => {
      for (const key of keys) {
        const offset = key.column - range.column[0];
        const order = compareSpreadsheetSortCells(
          left.cells[offset] ?? null,
          right.cells[offset] ?? null,
          key.direction,
        );
        if (order) return order;
      }
      return left.index - right.index;
    });
  const translatedRows: (Cell | null)[][] = [];
  for (const [targetIndex, row] of sorted.entries()) {
    const translated = translateSpreadsheetSortRow(
      row.cells,
      targetIndex - row.index,
    );
    if (!translated) {
      return spreadsheetSortError('formula-reference-out-of-range');
    }
    translatedRows.push(translated);
  }
  return {
    ok: true,
    rows: header
      ? [header as (Cell | null)[], ...translatedRows]
      : translatedRows,
  };
}

export function spreadsheetSortFailureMessage(
  result: SpreadsheetSortValidationResult | SpreadsheetSortResult,
): string | null {
  return result.ok ? null : result.message;
}

function spreadsheetSortRangeError(
  range: SpreadsheetCellRange,
): Extract<SpreadsheetSortValidationResult, { ok: false }> | null {
  const area = spreadsheetCellRangeArea(range);
  if (!Number.isSafeInteger(area) || area > MAX_SPREADSHEET_SORT_CELLS) {
    return spreadsheetSortError('range-too-large');
  }
  return null;
}

function compareSpreadsheetSortCells(
  left: Cell | null,
  right: Cell | null,
  direction: SpreadsheetSortDirection,
): number {
  const leftValue = spreadsheetSortValue(left);
  const rightValue = spreadsheetSortValue(right);
  if (leftValue === null) return rightValue === null ? 0 : 1;
  if (rightValue === null) return -1;
  const order =
    typeof leftValue === 'number' && typeof rightValue === 'number'
      ? leftValue - rightValue
      : spreadsheetSortCollator.compare(String(leftValue), String(rightValue));
  return direction === 'ascending' ? order : -order;
}

function spreadsheetSortValue(cell: Cell | null): number | string | null {
  const value = cell?.v ?? cell?.m;
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (typeof value === 'string' && value.trim()) return value.trim();
  if (typeof value === 'boolean') return value ? 1 : 0;
  return null;
}

function spreadsheetSortRowsHaveHeader(
  rows: readonly (readonly (Cell | null)[])[],
): boolean {
  const first = rows[0];
  if (!first?.length) return false;
  const labels = first.map(spreadsheetSortHeaderText);
  return (
    labels.every(Boolean) &&
    new Set(labels.map((label) => label.toLocaleLowerCase())).size ===
      labels.length
  );
}

function translateSpreadsheetSortRow(
  row: readonly (Cell | null)[],
  rowOffset: number,
): (Cell | null)[] | null {
  if (rowOffset === 0) return row as (Cell | null)[];
  const translated: (Cell | null)[] = [];
  for (const cell of row) {
    if (!cell?.f) {
      translated.push(cell);
      continue;
    }
    const formula = translateSpreadsheetFormula(cell.f, rowOffset, 0);
    if (formula === null) return null;
    translated.push(formula === cell.f ? cell : { ...cell, f: formula });
  }
  return translated;
}

function spreadsheetSortHeaderText(cell: Cell | null | undefined): string {
  const value = cell?.m ?? cell?.v;
  if (typeof value !== 'string') return '';
  return Array.from(value.trim())
    .slice(0, MAX_SPREADSHEET_SORT_HEADER_LENGTH)
    .join('');
}

function spreadsheetSortColumnLabel(column: number): string {
  return formatSpreadsheetCellRanges([
    { row: [0, 0], column: [column, column] },
  ]).replace(/\d+$/, '');
}

function spreadsheetSortError(
  code: SpreadsheetSortErrorCode,
): Extract<SpreadsheetSortValidationResult, { ok: false }> {
  const messages: Record<SpreadsheetSortErrorCode, string> = {
    'column-out-of-range': '排序列必须位于当前选定区域内。',
    'duplicate-key': '每个排序条件必须使用不同的列。',
    'formula-reference-out-of-range':
      '排序会使相对公式引用超出工作表范围，因此未应用任何更改。',
    'invalid-direction': '请选择有效的升序或降序次序。',
    'invalid-matrix': '排序只能应用到已完整读取的矩形区域。',
    'invalid-range': '请选择一个有效的连续单元格区域。',
    'missing-key': '请至少添加一个排序条件。',
    'not-enough-rows': '当前区域没有足够的数据行可供排序。',
    'range-too-large': `一次最多可排序 ${MAX_SPREADSHEET_SORT_CELLS.toLocaleString('en-US')} 个单元格。`,
    'too-many-keys': `一次最多可设置 ${MAX_SPREADSHEET_SORT_KEYS} 个排序条件。`,
    'unsupported-linked-cell':
      '当前区域包含坐标关联的超链接，尚不能安全地随排序移动。',
  };
  return { ok: false, code, message: messages[code] };
}

const spreadsheetSortCollator = new Intl.Collator('zh-CN', {
  numeric: true,
  sensitivity: 'base',
});
