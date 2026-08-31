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
import {
  createSpreadsheetSortDirectAppearanceRows,
  normalizeSpreadsheetSortColor,
  normalizeSpreadsheetSortIcon,
  type SpreadsheetSortAppearanceKind,
  type SpreadsheetSortAppearancePosition,
  type SpreadsheetSortAppearanceRows,
  type SpreadsheetSortIconTarget,
  spreadsheetSortAppearanceTargetValue,
} from './spreadsheet-sort-appearance';
import {
  DEFAULT_SPREADSHEET_SORT_TEXT_OPTIONS,
  isSpreadsheetSortTextMethod,
  type SpreadsheetSortTextOptions,
} from './spreadsheet-sort-collation';
import {
  mergeSpreadsheetSortCustomLists,
  type SpreadsheetSortCustomList,
  validateSpreadsheetSortCustomList,
} from './spreadsheet-sort-custom-list';

export type {
  SpreadsheetSortAppearanceRows,
  SpreadsheetSortCellAppearance,
  SpreadsheetSortIconTarget,
} from './spreadsheet-sort-appearance';
export type {
  SpreadsheetSortTextMethod,
  SpreadsheetSortTextOptions,
} from './spreadsheet-sort-collation';
export type { SpreadsheetSortCustomList } from './spreadsheet-sort-custom-list';

export const MAX_SPREADSHEET_SORT_KEYS = 64;
export const MAX_SPREADSHEET_SORT_CELLS = 1_000_000;

const MAX_SPREADSHEET_SORT_HEADER_LENGTH = 48;

export type SpreadsheetSortDirection = 'ascending' | 'descending';
export type SpreadsheetSortOrientation = 'top-to-bottom' | 'left-to-right';

export type SpreadsheetSortKey =
  | {
      color?: never;
      customList?: never;
      direction: SpreadsheetSortDirection;
      icon?: never;
      index: number;
      position?: never;
      sortOn?: never;
    }
  | {
      color?: never;
      customList: readonly string[];
      direction?: never;
      icon?: never;
      index: number;
      position?: never;
      sortOn?: never;
    }
  | {
      color: string | null;
      customList?: never;
      direction?: never;
      icon?: never;
      index: number;
      position: SpreadsheetSortAppearancePosition;
      sortOn: Exclude<SpreadsheetSortAppearanceKind, 'icon'>;
    }
  | {
      color?: never;
      customList?: never;
      direction?: never;
      icon: SpreadsheetSortIconTarget;
      index: number;
      position: SpreadsheetSortAppearancePosition;
      sortOn: 'icon';
    };

export interface SpreadsheetSortOptions extends SpreadsheetSortTextOptions {
  orientation: SpreadsheetSortOrientation;
}

export interface SpreadsheetSortDialogValue extends SpreadsheetSortOptions {
  hasHeader: boolean;
  keys: SpreadsheetSortKey[];
}

export interface SpreadsheetSortTarget {
  activeColumn: number;
  activeRow: number;
  range: SpreadsheetCellRange;
  scope?: SpreadsheetSortOwnedScope;
}

export type SpreadsheetSortOwnedScope =
  | { hasHeader: true; kind: 'auto-filter' }
  | { hasHeader: boolean; kind: 'table'; tableId: string };

export interface SpreadsheetSortOwnedRange {
  range: SpreadsheetCellRange;
  scope: SpreadsheetSortOwnedScope;
}

export interface SpreadsheetSortRangeCandidate {
  available: boolean;
  range: SpreadsheetCellRange;
  scope?: SpreadsheetSortOwnedScope;
}

export type SpreadsheetSortIntent =
  | { type: 'custom' }
  | { type: 'quick'; direction: SpreadsheetSortDirection };

export interface SpreadsheetSortOpenRequest {
  activeColumn: number;
  activeRow: number;
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
  ownedScope?: SpreadsheetSortOwnedScope;
  selectedRangeReference: string;
  sheetName: string;
}

export type SpreadsheetSortRequest = Omit<
  SpreadsheetSortDialogValue,
  keyof SpreadsheetSortTextOptions
> &
  Partial<SpreadsheetSortTextOptions> & {
    range: SpreadsheetCellRange;
    scope?: SpreadsheetSortOwnedScope;
    sheetId: string;
  };

export interface SpreadsheetSortNormalizedRequest
  extends SpreadsheetSortDialogValue {
  range: SpreadsheetCellRange;
  scope?: SpreadsheetSortOwnedScope;
  sheetId: string;
}

export interface SpreadsheetSortField {
  index: number;
  label: string;
}

export interface SpreadsheetSortDialogSource {
  activeRow: number;
  appearanceRows: SpreadsheetSortAppearanceRows;
  columns: SpreadsheetSortField[];
  customLists: readonly SpreadsheetSortCustomList[];
  range: SpreadsheetCellRange;
  rangeReference: string;
  rows: SpreadsheetSortField[];
  scope?: SpreadsheetSortOwnedScope;
  sheetId: string;
  sheetName: string;
  value: SpreadsheetSortDialogValue;
}

export type SpreadsheetSortErrorCode =
  | 'column-out-of-range'
  | 'duplicate-key'
  | 'formula-reference-out-of-range'
  | 'invalid-appearance'
  | 'invalid-case-sensitivity'
  | 'invalid-custom-list'
  | 'invalid-direction'
  | 'invalid-header'
  | 'invalid-matrix'
  | 'invalid-orientation'
  | 'invalid-range'
  | 'invalid-scope'
  | 'invalid-text-method'
  | 'missing-key'
  | 'not-enough-columns'
  | 'not-enough-rows'
  | 'range-too-large'
  | 'row-out-of-range'
  | 'too-many-keys'
  | 'unsupported-text-method'
  | 'unsupported-linked-cell';

export type SpreadsheetSortValidationResult =
  | { ok: true; request: SpreadsheetSortNormalizedRequest }
  | { code: SpreadsheetSortErrorCode; message: string; ok: false };

export type SpreadsheetSortResult =
  | { ok: true; rows: (Cell | null)[][]; sourceIndexes: number[] }
  | { code: SpreadsheetSortErrorCode; message: string; ok: false };

export function createSpreadsheetSortRangePlan(
  sheet: WorkSpreadsheetSheet,
  selectedRange: SpreadsheetCellRange,
): SpreadsheetSortRangePlan | null {
  const selected = normalizeSpreadsheetCellRange(selectedRange);
  if (!selected) return null;
  const owned = spreadsheetSortOwnedRangeContaining(sheet, selected);
  if (owned) {
    return {
      selectedRange: selected,
      expandedRange: spreadsheetCellRangesEqual(owned.range, selected)
        ? null
        : owned.range,
    };
  }
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
    ...(request.expanded.scope ? { ownedScope: request.expanded.scope } : {}),
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
  if (!spreadsheetSortRowsMatchRange(rows, range)) {
    return null;
  }
  const hasHeader = height > 2 && spreadsheetSortRowsHaveHeader(rows);
  const resolvedHasHeader = target.scope?.hasHeader ?? hasHeader;
  const activeColumn = spreadsheetSortActiveIndex(
    target.activeColumn,
    range.column,
  );
  const activeRow = spreadsheetSortActiveIndex(target.activeRow, range.row);
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
    activeRow,
    appearanceRows: resolvedAppearanceRows,
    customLists: mergeSpreadsheetSortCustomLists(customLists),
    ...(target.scope ? { scope: target.scope } : {}),
    columns: Array.from({ length: width }, (_, offset) => {
      const index = range.column[0] + offset;
      const name = spreadsheetSortHeaderText(header[offset]);
      const coordinate = spreadsheetSortColumnLabel(index);
      return {
        index,
        label: name ? `${coordinate}（${name}）` : coordinate,
      };
    }),
    rows: Array.from({ length: height }, (_, offset) => {
      const index = range.row[0] + offset;
      return { index, label: `行 ${index + 1}` };
    }),
    value: {
      orientation: 'top-to-bottom',
      ...DEFAULT_SPREADSHEET_SORT_TEXT_OPTIONS,
      hasHeader: resolvedHasHeader,
      keys: [{ index: activeColumn, direction: 'ascending' }],
    },
  };
}

export function spreadsheetSortRowsMatchRange(
  rows: readonly (readonly (Cell | null)[])[],
  range: SpreadsheetCellRange,
): boolean {
  const normalized = normalizeSpreadsheetCellRange(range);
  if (!normalized) return false;
  const width = normalized.column[1] - normalized.column[0] + 1;
  const height = normalized.row[1] - normalized.row[0] + 1;
  return rows.length === height && rows.every((row) => row.length === width);
}

export function spreadsheetSortRowsFromSheet(
  sheet: WorkSpreadsheetSheet,
  range: SpreadsheetCellRange,
): (Cell | null)[][] | null {
  const normalized = normalizeSpreadsheetCellRange(range);
  if (!normalized || spreadsheetSortRangeError(normalized)) return null;
  const width = normalized.column[1] - normalized.column[0] + 1;
  const height = normalized.row[1] - normalized.row[0] + 1;
  const rows = Array.from({ length: height }, () =>
    Array<Cell | null>(width).fill(null),
  );
  for (let row = normalized.row[0]; row <= normalized.row[1]; row += 1) {
    const source = sheet.data?.[row];
    if (!source) continue;
    const target = rows[row - normalized.row[0]]!;
    for (
      let column = normalized.column[0];
      column <= normalized.column[1];
      column += 1
    ) {
      target[column - normalized.column[0]] = source[column] ?? null;
    }
  }
  for (const entry of sheet.celldata ?? []) {
    if (
      entry.r < normalized.row[0] ||
      entry.r > normalized.row[1] ||
      entry.c < normalized.column[0] ||
      entry.c > normalized.column[1]
    ) {
      continue;
    }
    const row = rows[entry.r - normalized.row[0]]!;
    const column = entry.c - normalized.column[0];
    if (row[column] === null) row[column] = entry.v ?? null;
  }
  return rows;
}

export function spreadsheetSortOwnedRangeForExactRange(
  sheet: WorkSpreadsheetSheet,
  range: SpreadsheetCellRange,
): SpreadsheetSortOwnedRange | null {
  const normalized = normalizeSpreadsheetCellRange(range);
  if (!normalized) return null;
  const matches = spreadsheetSortOwnedRanges(sheet).filter((candidate) =>
    spreadsheetCellRangesEqual(candidate.range, normalized),
  );
  return matches.length === 1 ? matches[0]! : null;
}

function spreadsheetSortOwnedRangeContaining(
  sheet: WorkSpreadsheetSheet,
  range: SpreadsheetCellRange,
): SpreadsheetSortOwnedRange | null {
  const matches = spreadsheetSortOwnedRanges(sheet).filter((candidate) =>
    spreadsheetSortRangeContains(candidate.range, range),
  );
  return matches.length === 1 ? matches[0]! : null;
}

function spreadsheetSortOwnedRanges(
  sheet: WorkSpreadsheetSheet,
): SpreadsheetSortOwnedRange[] {
  const ranges: SpreadsheetSortOwnedRange[] = [];
  const autoFilter = normalizeSpreadsheetCellRange(
    sheet.filter_select ?? { row: [], column: [] },
  );
  if (autoFilter && autoFilter.row[1] > autoFilter.row[0]) {
    ranges.push({
      range: autoFilter,
      scope: { kind: 'auto-filter', hasHeader: true },
    });
  }
  for (const table of sheet.tables ?? []) {
    const tableRange = normalizeSpreadsheetCellRange(table.range);
    if (!tableRange || !table.id || table.columns.length === 0) continue;
    const width = tableRange.column[1] - tableRange.column[0] + 1;
    if (width !== table.columns.length) continue;
    const endRow = tableRange.row[1] - Number(table.totalsRow);
    if (endRow <= tableRange.row[0]) continue;
    ranges.push({
      range: {
        row: [tableRange.row[0], endRow],
        column: [...tableRange.column],
      },
      scope: {
        kind: 'table',
        tableId: table.id,
        hasHeader: table.headerRow,
      },
    });
  }
  return ranges;
}

function spreadsheetSortRangeContains(
  outer: SpreadsheetCellRange,
  inner: SpreadsheetCellRange,
): boolean {
  return (
    outer.row[0] <= inner.row[0] &&
    outer.row[1] >= inner.row[1] &&
    outer.column[0] <= inner.column[0] &&
    outer.column[1] >= inner.column[1]
  );
}

export function validateSpreadsheetSortRequest(
  request: SpreadsheetSortRequest,
): SpreadsheetSortValidationResult {
  const range = normalizeSpreadsheetCellRange(request.range);
  if (
    !range ||
    typeof request.sheetId !== 'string' ||
    !request.sheetId.trim()
  ) {
    return spreadsheetSortError('invalid-range');
  }
  const scope = normalizeSpreadsheetSortOwnedScope(request.scope);
  if (request.scope !== undefined && !scope) {
    return spreadsheetSortError('invalid-scope');
  }
  if (
    request.orientation !== 'top-to-bottom' &&
    request.orientation !== 'left-to-right'
  ) {
    return spreadsheetSortError('invalid-orientation');
  }
  const caseSensitive =
    request.caseSensitive === undefined
      ? DEFAULT_SPREADSHEET_SORT_TEXT_OPTIONS.caseSensitive
      : request.caseSensitive;
  if (typeof caseSensitive !== 'boolean') {
    return spreadsheetSortError('invalid-case-sensitivity');
  }
  const textMethod =
    request.textMethod === undefined
      ? DEFAULT_SPREADSHEET_SORT_TEXT_OPTIONS.textMethod
      : request.textMethod;
  if (!isSpreadsheetSortTextMethod(textMethod)) {
    return spreadsheetSortError('invalid-text-method');
  }
  if (request.orientation === 'left-to-right' && request.hasHeader) {
    return spreadsheetSortError('invalid-header');
  }
  const rangeError = spreadsheetSortRangeError(range);
  if (rangeError) return rangeError;
  if (!Array.isArray(request.keys) || !request.keys.length) {
    return spreadsheetSortError('missing-key');
  }
  if (request.keys.length > MAX_SPREADSHEET_SORT_KEYS) {
    return spreadsheetSortError('too-many-keys');
  }

  const keyRange =
    request.orientation === 'top-to-bottom' ? range.column : range.row;
  const outOfRangeCode =
    request.orientation === 'top-to-bottom'
      ? 'column-out-of-range'
      : 'row-out-of-range';
  const keys: SpreadsheetSortKey[] = [];
  for (const key of request.keys) {
    if (
      !Number.isSafeInteger(key.index) ||
      key.index < keyRange[0] ||
      key.index > keyRange[1]
    ) {
      return spreadsheetSortError(outOfRangeCode);
    }
    if (key.sortOn !== undefined) {
      if (
        (key.position !== 'first' && key.position !== 'last') ||
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
          index: key.index,
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
            index: key.index,
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
          index: key.index,
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
        index: key.index,
        direction: key.direction,
      })
    ) {
      return spreadsheetSortError('duplicate-key');
    }
  }

  if (request.orientation === 'top-to-bottom') {
    const rowCount = range.row[1] - range.row[0] + 1;
    if (rowCount - (request.hasHeader ? 1 : 0) < 2) {
      return spreadsheetSortError('not-enough-rows');
    }
  } else if (range.column[1] - range.column[0] + 1 < 2) {
    return spreadsheetSortError('not-enough-columns');
  }

  return {
    ok: true,
    request: {
      sheetId: request.sheetId,
      range,
      orientation: request.orientation,
      caseSensitive,
      textMethod,
      hasHeader: Boolean(request.hasHeader),
      keys,
      ...(scope ? { scope } : {}),
    },
  };
}

export function spreadsheetSortFailureMessage(
  result: SpreadsheetSortValidationResult | SpreadsheetSortResult,
): string | null {
  return result.ok ? null : result.message;
}

export function spreadsheetSortAppearanceRowsMatch(
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

export function spreadsheetSortError(
  code: SpreadsheetSortErrorCode,
): Extract<SpreadsheetSortValidationResult, { ok: false }> {
  const messages: Record<SpreadsheetSortErrorCode, string> = {
    'column-out-of-range': '排序列必须位于当前选定区域内。',
    'duplicate-key': '同一行或列不能重复使用相同的值或外观排序条件。',
    'formula-reference-out-of-range':
      '排序会使相对公式引用超出工作表范围，因此未应用任何更改。',
    'invalid-appearance':
      '外观排序条件或当前颜色/图标快照无效，因此未应用任何更改。',
    'invalid-case-sensitivity': '请选择有效的大小写比较规则。',
    'invalid-custom-list': '自定义排序序列无效，请检查项目数量、长度和重复项。',
    'invalid-direction': '请选择有效的升序或降序次序。',
    'invalid-header': '按行排序会移动所有选定列，不能保留标题列。',
    'invalid-matrix': '排序只能应用到已完整读取的矩形区域。',
    'invalid-orientation': '请选择有效的按列或按行排序方向。',
    'invalid-range': '请选择一个有效的连续单元格区域。',
    'invalid-scope': '排序区域的表格或筛选所有权已变化，请重新打开排序。',
    'invalid-text-method': '请选择拼音排序或笔画排序。',
    'missing-key': '请至少添加一个排序条件。',
    'not-enough-columns': '当前区域没有足够的数据列可供按行排序。',
    'not-enough-rows': '当前区域没有足够的数据行可供排序。',
    'range-too-large': `一次最多可排序 ${MAX_SPREADSHEET_SORT_CELLS.toLocaleString('en-US')} 个单元格。`,
    'row-out-of-range': '排序行必须位于当前选定区域内。',
    'too-many-keys': `一次最多可设置 ${MAX_SPREADSHEET_SORT_KEYS} 个排序条件。`,
    'unsupported-text-method':
      '当前运行环境不支持所选中文排序方法，因此未应用任何更改。',
    'unsupported-linked-cell':
      '当前区域包含坐标关联的超链接，尚不能安全地随排序移动。',
  };
  return { ok: false, code, message: messages[code] };
}

function normalizeSpreadsheetSortOwnedScope(
  value: SpreadsheetSortOwnedScope | undefined,
): SpreadsheetSortOwnedScope | null {
  if (value === undefined) return null;
  if (!value || typeof value !== 'object' || Array.isArray(value)) return null;
  if (
    value.kind === 'auto-filter' &&
    value.hasHeader === true &&
    Object.keys(value).every((key) => key === 'kind' || key === 'hasHeader')
  ) {
    return { kind: 'auto-filter', hasHeader: true };
  }
  if (
    value.kind === 'table' &&
    typeof value.tableId === 'string' &&
    Boolean(value.tableId.trim()) &&
    typeof value.hasHeader === 'boolean' &&
    Object.keys(value).every(
      (key) => key === 'kind' || key === 'tableId' || key === 'hasHeader',
    )
  ) {
    return {
      kind: 'table',
      tableId: value.tableId,
      hasHeader: value.hasHeader,
    };
  }
  return null;
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
    return `${key.index}:${spreadsheetSortAppearanceTargetValue({
      kind: key.sortOn,
      color: key.color,
    })}`;
  }
  if (key.sortOn === 'icon') {
    return `${key.index}:${spreadsheetSortAppearanceTargetValue({
      kind: 'icon',
      icon: key.icon,
    })}`;
  }
  return `${key.index}:values`;
}

function spreadsheetSortActiveIndex(
  active: number,
  range: readonly [number, number],
): number {
  return active >= range[0] && active <= range[1] ? active : range[0];
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
