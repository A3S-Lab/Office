import type { Cell } from '@fortune-sheet/core';
import { formatSpreadsheetCellRanges } from '../work-spreadsheet-ranges';
import type { WorkSpreadsheetSheet } from '../work-types';
import {
  normalizeSpreadsheetCellRange,
  type SpreadsheetCellRange,
  spreadsheetCellRangeArea,
  spreadsheetCellRangesEqual,
} from './spreadsheet-cell-range';
import { spreadsheetCurrentRegion } from './spreadsheet-current-region';
import { translateSpreadsheetFormula } from './spreadsheet-paste-special-cell';
import {
  createSpreadsheetSortDirectAppearanceRows,
  normalizeSpreadsheetSortColor,
  normalizeSpreadsheetSortIcon,
  spreadsheetSortAppearanceTargetValue,
  spreadsheetSortCellMatchesAppearance,
  type SpreadsheetSortAppearanceKind,
  type SpreadsheetSortAppearancePosition,
  type SpreadsheetSortAppearanceRows,
  type SpreadsheetSortAppearanceTarget,
  type SpreadsheetSortIconTarget,
} from './spreadsheet-sort-appearance';
import {
  mergeSpreadsheetSortCustomLists,
  spreadsheetSortCustomListMatchKey,
  type SpreadsheetSortCustomList,
  validateSpreadsheetSortCustomList,
} from './spreadsheet-sort-custom-list';

export type { SpreadsheetSortCustomList } from './spreadsheet-sort-custom-list';
export type {
  SpreadsheetSortAppearanceRows,
  SpreadsheetSortCellAppearance,
  SpreadsheetSortIconTarget,
} from './spreadsheet-sort-appearance';

export const MAX_SPREADSHEET_SORT_KEYS = 64;
export const MAX_SPREADSHEET_SORT_CELLS = 1_000_000;

const MAX_SPREADSHEET_SORT_HEADER_LENGTH = 48;

export type SpreadsheetSortDirection = 'ascending' | 'descending';

export type SpreadsheetSortKey =
  | {
      color?: never;
      column: number;
      customList?: never;
      direction: SpreadsheetSortDirection;
      icon?: never;
      position?: never;
      sortOn?: never;
    }
  | {
      color?: never;
      column: number;
      customList: readonly string[];
      direction?: never;
      icon?: never;
      position?: never;
      sortOn?: never;
    }
  | {
      color: string | null;
      column: number;
      customList?: never;
      direction?: never;
      icon?: never;
      position: SpreadsheetSortAppearancePosition;
      sortOn: Exclude<SpreadsheetSortAppearanceKind, 'icon'>;
    }
  | {
      color?: never;
      column: number;
      customList?: never;
      direction?: never;
      icon: SpreadsheetSortIconTarget;
      position: SpreadsheetSortAppearancePosition;
      sortOn: 'icon';
    };

export interface SpreadsheetSortDialogValue {
  hasHeader: boolean;
  keys: SpreadsheetSortKey[];
}

export interface SpreadsheetSortTarget {
  activeColumn: number;
  range: SpreadsheetCellRange;
}

export interface SpreadsheetSortRangeCandidate {
  available: boolean;
  range: SpreadsheetCellRange;
}

export type SpreadsheetSortIntent =
  | { type: 'custom' }
  | { type: 'quick'; direction: SpreadsheetSortDirection };

export interface SpreadsheetSortOpenRequest {
  activeColumn: number;
  expanded?: SpreadsheetSortRangeCandidate;
  intent: SpreadsheetSortIntent;
  selected: SpreadsheetSortRangeCandidate;
  sheetId: string;
}

export interface SpreadsheetSortRangePlan {
  expandedRange: SpreadsheetCellRange | null;
  selectedRange: SpreadsheetCellRange;
}

export type SpreadsheetSortRangeChoice = 'expand' | 'selection';

export interface SpreadsheetSortRangeDialogSource {
  canSortExpandedRange: boolean;
  canSortSelection: boolean;
  expandedRangeReference: string;
  selectedRangeReference: string;
  sheetName: string;
}

export interface SpreadsheetSortRequest extends SpreadsheetSortDialogValue {
  range: SpreadsheetCellRange;
  sheetId: string;
}

export interface SpreadsheetSortDialogSource {
  appearanceRows: SpreadsheetSortAppearanceRows;
  columns: Array<{ column: number; label: string }>;
  customLists: readonly SpreadsheetSortCustomList[];
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
  | 'invalid-appearance'
  | 'invalid-custom-list'
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

export function createSpreadsheetSortRangePlan(
  sheet: WorkSpreadsheetSheet,
  selectedRange: SpreadsheetCellRange,
): SpreadsheetSortRangePlan | null {
  const selected = normalizeSpreadsheetCellRange(selectedRange);
  if (!selected) return null;
  const currentRegion = spreadsheetCurrentRegion(sheet, selected);
  return {
    selectedRange: selected,
    expandedRange:
      currentRegion && !spreadsheetCellRangesEqual(currentRegion, selected)
        ? currentRegion
        : null,
  };
}

export function createSpreadsheetSortRangeDialogSource(
  sheetName: string,
  request: SpreadsheetSortOpenRequest,
): SpreadsheetSortRangeDialogSource | null {
  if (!request.expanded) return null;
  return {
    sheetName,
    selectedRangeReference: formatSpreadsheetCellRanges([
      request.selected.range,
    ]),
    expandedRangeReference: formatSpreadsheetCellRanges([
      request.expanded.range,
    ]),
    canSortSelection: request.selected.available,
    canSortExpandedRange: request.expanded.available,
  };
}

export function createSpreadsheetSortDialogSource(
  sheetId: string,
  sheetName: string,
  target: SpreadsheetSortTarget,
  rows: readonly (readonly (Cell | null)[])[],
  customLists: readonly SpreadsheetSortCustomList[] = [],
  appearanceRows?: SpreadsheetSortAppearanceRows,
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
  const resolvedAppearanceRows =
    appearanceRows && spreadsheetSortAppearanceRowsMatch(rows, appearanceRows)
      ? appearanceRows
      : createSpreadsheetSortDirectAppearanceRows(rows);
  return {
    sheetId,
    sheetName,
    range,
    rangeReference: formatSpreadsheetCellRanges([range]),
    appearanceRows: resolvedAppearanceRows,
    customLists: mergeSpreadsheetSortCustomLists(customLists),
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
  const keys: SpreadsheetSortKey[] = [];
  for (const key of request.keys) {
    if (!Number.isSafeInteger(key.column)) {
      return spreadsheetSortError('column-out-of-range');
    }
    if (key.column < range.column[0] || key.column > range.column[1]) {
      return spreadsheetSortError('column-out-of-range');
    }
    if (key.sortOn !== undefined) {
      if (
        (key.position !== 'top' && key.position !== 'bottom') ||
        key.customList !== undefined ||
        key.direction !== undefined
      ) {
        return spreadsheetSortError('invalid-appearance');
      }
      if (key.sortOn === 'cell-color' || key.sortOn === 'font-color') {
        const color = normalizeSpreadsheetSortColor(key.color);
        if (color === undefined || key.icon !== undefined) {
          return spreadsheetSortError('invalid-appearance');
        }
        const normalized: SpreadsheetSortKey = {
          column: key.column,
          sortOn: key.sortOn,
          color,
          position: key.position,
        };
        if (!appendSpreadsheetSortKey(keys, normalized)) {
          return spreadsheetSortError('duplicate-key');
        }
        continue;
      }
      if (key.sortOn === 'icon' && key.color === undefined) {
        const icon = normalizeSpreadsheetSortIcon(key.icon);
        if (icon) {
          const normalized: SpreadsheetSortKey = {
            column: key.column,
            sortOn: 'icon',
            icon,
            position: key.position,
          };
          if (!appendSpreadsheetSortKey(keys, normalized)) {
            return spreadsheetSortError('duplicate-key');
          }
          continue;
        }
      }
      return spreadsheetSortError('invalid-appearance');
    }
    if (key.customList !== undefined) {
      if (key.direction !== undefined) {
        return spreadsheetSortError('invalid-custom-list');
      }
      const customList = validateSpreadsheetSortCustomList(key.customList);
      if (!customList.ok) return spreadsheetSortError('invalid-custom-list');
      if (
        !appendSpreadsheetSortKey(keys, {
          column: key.column,
          customList: customList.entries,
        })
      ) {
        return spreadsheetSortError('duplicate-key');
      }
      continue;
    }
    if (key.direction !== 'ascending' && key.direction !== 'descending') {
      return spreadsheetSortError('invalid-direction');
    }
    if (
      !appendSpreadsheetSortKey(keys, {
        column: key.column,
        direction: key.direction,
      })
    ) {
      return spreadsheetSortError('duplicate-key');
    }
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
  appearanceRows?: SpreadsheetSortAppearanceRows,
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
  const resolvedAppearanceRows =
    appearanceRows ?? createSpreadsheetSortDirectAppearanceRows(source);
  if (
    keys.some((key) => key.sortOn !== undefined) &&
    !spreadsheetSortAppearanceRowsMatch(source, resolvedAppearanceRows)
  ) {
    return spreadsheetSortError('invalid-appearance');
  }

  const header = hasHeader ? source[0] : undefined;
  const rows = source.slice(hasHeader ? 1 : 0);
  const appearanceOffset = hasHeader ? 1 : 0;
  const compiledKeys = keys.map((key) =>
    compileSpreadsheetSortKey(key, range.column[0]),
  );
  if (
    compiledKeys.some(
      (key) =>
        key.kind === 'appearance' &&
        !resolvedAppearanceRows
          .slice(appearanceOffset)
          .some((row) =>
            spreadsheetSortCellMatchesAppearance(row[key.offset], key.target),
          ),
    )
  ) {
    return spreadsheetSortError('invalid-appearance');
  }
  const sorted = rows
    .map((cells, index) => ({
      appearance: resolvedAppearanceRows[index + appearanceOffset] ?? [],
      cells,
      index,
    }))
    .sort((left, right) => {
      for (const key of compiledKeys) {
        const order = compareSpreadsheetSortCells(
          left.cells[key.offset] ?? null,
          right.cells[key.offset] ?? null,
          left.appearance[key.offset],
          right.appearance[key.offset],
          key,
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
  leftAppearance: SpreadsheetSortAppearanceRows[number][number] | undefined,
  rightAppearance: SpreadsheetSortAppearanceRows[number][number] | undefined,
  key: SpreadsheetSortCompiledKey,
): number {
  if (key.kind === 'appearance') {
    const leftMatches = spreadsheetSortCellMatchesAppearance(
      leftAppearance,
      key.target,
    );
    const rightMatches = spreadsheetSortCellMatchesAppearance(
      rightAppearance,
      key.target,
    );
    if (leftMatches === rightMatches) return 0;
    const matchedFirst = key.position === 'top';
    return leftMatches === matchedFirst ? -1 : 1;
  }
  const leftValue = spreadsheetSortValue(left);
  const rightValue = spreadsheetSortValue(right);
  if (leftValue === null) return rightValue === null ? 0 : 1;
  if (rightValue === null) return -1;
  if (key.kind === 'custom-list') {
    const leftRank = key.customRanks.get(
      spreadsheetSortCustomListMatchKey(leftValue),
    );
    const rightRank = key.customRanks.get(
      spreadsheetSortCustomListMatchKey(rightValue),
    );
    if (leftRank !== undefined || rightRank !== undefined) {
      if (leftRank === undefined) return 1;
      if (rightRank === undefined) return -1;
      return leftRank - rightRank;
    }
  }
  const order =
    typeof leftValue === 'number' && typeof rightValue === 'number'
      ? leftValue - rightValue
      : spreadsheetSortCollator.compare(String(leftValue), String(rightValue));
  return key.kind === 'value' && key.direction === 'descending'
    ? -order
    : order;
}

type SpreadsheetSortCompiledKey =
  | {
      kind: 'appearance';
      offset: number;
      position: SpreadsheetSortAppearancePosition;
      target: SpreadsheetSortAppearanceTarget;
    }
  | {
      customRanks: ReadonlyMap<string, number>;
      kind: 'custom-list';
      offset: number;
    }
  | {
      direction: SpreadsheetSortDirection;
      kind: 'value';
      offset: number;
    };

function compileSpreadsheetSortKey(
  key: SpreadsheetSortKey,
  firstColumn: number,
): SpreadsheetSortCompiledKey {
  if (key.sortOn === 'cell-color' || key.sortOn === 'font-color') {
    return {
      kind: 'appearance',
      offset: key.column - firstColumn,
      position: key.position,
      target: { kind: key.sortOn, color: key.color },
    };
  }
  if (key.sortOn === 'icon') {
    return {
      kind: 'appearance',
      offset: key.column - firstColumn,
      position: key.position,
      target: { kind: 'icon', icon: key.icon },
    };
  }
  if (key.customList !== undefined) {
    return {
      kind: 'custom-list',
      offset: key.column - firstColumn,
      customRanks: new Map(
        key.customList.map((entry, index) => [
          spreadsheetSortCustomListMatchKey(entry),
          index,
        ]),
      ),
    };
  }
  return {
    kind: 'value',
    offset: key.column - firstColumn,
    direction: key.direction ?? 'ascending',
  };
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

function spreadsheetSortAppearanceRowsMatch(
  rows: readonly (readonly (Cell | null)[])[],
  appearances: SpreadsheetSortAppearanceRows,
): boolean {
  return (
    rows.length === appearances.length &&
    rows.every(
      (row, index) => row.length === (appearances[index]?.length ?? -1),
    )
  );
}

function appendSpreadsheetSortKey(
  keys: SpreadsheetSortKey[],
  key: SpreadsheetSortKey,
): boolean {
  const identity = spreadsheetSortKeyIdentity(key);
  if (
    keys.some((candidate) => spreadsheetSortKeyIdentity(candidate) === identity)
  ) {
    return false;
  }
  keys.push(key);
  return true;
}

function spreadsheetSortKeyIdentity(key: SpreadsheetSortKey): string {
  if (key.sortOn === 'cell-color' || key.sortOn === 'font-color') {
    return `${key.column}:${spreadsheetSortAppearanceTargetValue({
      kind: key.sortOn,
      color: key.color,
    })}`;
  }
  if (key.sortOn === 'icon') {
    return `${key.column}:${spreadsheetSortAppearanceTargetValue({
      kind: 'icon',
      icon: key.icon,
    })}`;
  }
  return `${key.column}:values`;
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
  if (cell?.f || typeof cell?.v !== 'string') return '';
  const value = cell.v;
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
    'duplicate-key': '同一列不能重复使用相同的值或外观排序条件。',
    'formula-reference-out-of-range':
      '排序会使相对公式引用超出工作表范围，因此未应用任何更改。',
    'invalid-appearance':
      '外观排序条件或当前颜色/图标快照无效，因此未应用任何更改。',
    'invalid-custom-list': '自定义排序序列无效，请检查项目数量、长度和重复项。',
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
