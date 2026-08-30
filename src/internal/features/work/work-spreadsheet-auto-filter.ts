import type { Cell } from '@fortune-sheet/core';
import type {
  WorkSpreadsheetCustomFilterCondition,
  WorkSpreadsheetDynamicFilter,
  WorkSpreadsheetFilter,
  WorkSpreadsheetFilterCriteria,
  WorkSpreadsheetSheet,
} from './work-types';
import { workSpreadsheetFilterTextIsBounded } from './work-spreadsheet-filter-contract';
import {
  type WorkSpreadsheetDynamicFilterContext,
  workSpreadsheetDynamicFilterMatcher,
} from './work-spreadsheet-dynamic-filter';
import { workSpreadsheetWildcardMatcher } from './work-spreadsheet-wildcard';

const AUTO_FILTER_CRITERIA_KIND = 'a3s-office-auto-filter-criteria';
const AUTO_FILTER_CRITERIA_VERSION = 1;

const SINGLE_VALUE_FILTER_TYPES = new Set([
  'equals',
  'not-equals',
  'contains',
  'does-not-contain',
  'begins-with',
  'does-not-begin-with',
  'ends-with',
  'does-not-end-with',
  'matches-wildcard',
  'does-not-match-wildcard',
  'greater-than',
  'greater-than-or-equal',
  'less-than',
  'less-than-or-equal',
]);

const DYNAMIC_FILTER_KINDS = new Set([
  'above-average',
  'below-average',
  'tomorrow',
  'today',
  'yesterday',
  'next-week',
  'this-week',
  'last-week',
  'next-month',
  'this-month',
  'last-month',
  'next-quarter',
  'this-quarter',
  'last-quarter',
  'next-year',
  'this-year',
  'last-year',
  'year-to-date',
  'quarter-1',
  'quarter-2',
  'quarter-3',
  'quarter-4',
  'month-1',
  'month-2',
  'month-3',
  'month-4',
  'month-5',
  'month-6',
  'month-7',
  'month-8',
  'month-9',
  'month-10',
  'month-11',
  'month-12',
]);

export interface WorkSpreadsheetAutoFilterRange {
  column: [number, number];
  row: [number, number];
}

interface WorkSpreadsheetAutoFilterCriteriaPayload {
  criteria: WorkSpreadsheetFilterCriteria;
  kind: typeof AUTO_FILTER_CRITERIA_KIND;
  manualHiddenOverlap?: string[];
  version: typeof AUTO_FILTER_CRITERIA_VERSION;
}

interface WorkSpreadsheetAutoFilterColumnState {
  caljs: WorkSpreadsheetAutoFilterCriteriaPayload;
  cindex: number;
  edc: number;
  edr: number;
  optionstate: true;
  rowhidden: Record<string, 0>;
  stc: number;
  str: number;
}

export function normalizedWorkSpreadsheetAutoFilterRange(
  selection: WorkSpreadsheetSheet['filter_select'],
): WorkSpreadsheetAutoFilterRange | null {
  const rowStart = finiteIndex(selection?.row?.[0]);
  const rowEnd = finiteIndex(selection?.row?.[1]);
  const columnStart = finiteIndex(selection?.column?.[0]);
  const columnEnd = finiteIndex(selection?.column?.[1]);
  if (
    rowStart === null ||
    rowEnd === null ||
    columnStart === null ||
    columnEnd === null
  ) {
    return null;
  }
  return {
    row: [Math.min(rowStart, rowEnd), Math.max(rowStart, rowEnd)],
    column: [
      Math.min(columnStart, columnEnd),
      Math.max(columnStart, columnEnd),
    ],
  };
}

export function workSpreadsheetAutoFilterCriteria(
  sheet: WorkSpreadsheetSheet | undefined,
  column: number,
): WorkSpreadsheetFilterCriteria | null {
  const range = normalizedWorkSpreadsheetAutoFilterRange(sheet?.filter_select);
  if (!sheet || !range || !columnInRange(range, column)) return null;
  return criteriaFromEntry(sheet.filter?.[String(column - range.column[0])]);
}

export function workSpreadsheetAutoFilterCriteriaEntries(
  sheet: WorkSpreadsheetSheet,
): WorkSpreadsheetFilter[] {
  const range = normalizedWorkSpreadsheetAutoFilterRange(sheet.filter_select);
  if (!range) return [];
  const entries: WorkSpreadsheetFilter[] = [];
  for (const [key, value] of Object.entries(sheet.filter ?? {})) {
    const column = Number(key);
    const criteria = criteriaFromEntry(value);
    if (
      !criteria ||
      !Number.isSafeInteger(column) ||
      column < 0 ||
      column > range.column[1] - range.column[0]
    ) {
      continue;
    }
    entries.push({ column, criteria });
  }
  return entries.sort((left, right) => left.column - right.column);
}

export function workSpreadsheetSheetWithAutoFilterCriteria(
  sheet: WorkSpreadsheetSheet,
  column: number,
  criteria: WorkSpreadsheetFilterCriteria,
  context: WorkSpreadsheetDynamicFilterContext = { now: new Date() },
): WorkSpreadsheetSheet | null {
  const normalizedCriteria = normalizeWorkSpreadsheetFilterCriteria(criteria);
  const range = normalizedWorkSpreadsheetAutoFilterRange(sheet.filter_select);
  if (!normalizedCriteria || !range || !columnInRange(range, column)) {
    return null;
  }
  const rowhidden = hiddenRowsForCriteria(
    sheet,
    range,
    column,
    normalizedCriteria,
    context,
  );
  if (!rowhidden) return null;
  const manuallyHiddenRows = workSpreadsheetAutoFilterManuallyHiddenRows(sheet);
  const filter = {
    ...(sheet.filter ?? {}),
    [String(column - range.column[0])]: filterColumnState(
      range,
      column,
      normalizedCriteria,
      rowhidden,
    ),
  };
  return sheetWithFilterState(sheet, filter, manuallyHiddenRows);
}

export function workSpreadsheetSheetWithoutAutoFilterCriteria(
  sheet: WorkSpreadsheetSheet,
  column: number,
): WorkSpreadsheetSheet | null {
  const range = normalizedWorkSpreadsheetAutoFilterRange(sheet.filter_select);
  if (!range || !columnInRange(range, column)) return null;
  const key = String(column - range.column[0]);
  if (!Object.hasOwn(sheet.filter ?? {}, key)) return null;
  const filter = { ...(sheet.filter ?? {}) };
  delete filter[key];
  return sheetWithFilterState(
    sheet,
    filter,
    workSpreadsheetAutoFilterManuallyHiddenRows(sheet),
  );
}

export function workSpreadsheetAutoFilterManuallyHiddenRows(
  sheet: WorkSpreadsheetSheet,
): Set<string> {
  const owned = filterHiddenRows(sheet.filter);
  const manuallyHidden = new Set(
    Object.keys(sheet.config?.rowhidden ?? {}).filter((row) => !owned.has(row)),
  );
  for (const entry of Object.values(sheet.filter ?? {})) {
    const overlap = criteriaPayloadFromEntry(entry)?.manualHiddenOverlap;
    if (!Array.isArray(overlap)) continue;
    for (const row of overlap) {
      if (validHiddenRowKey(row)) manuallyHidden.add(row);
    }
  }
  return manuallyHidden;
}

export function workSpreadsheetSheetWithImportedAutoFilterCriteria(
  sheet: WorkSpreadsheetSheet,
  entries: readonly WorkSpreadsheetFilter[],
  context: WorkSpreadsheetDynamicFilterContext = { now: new Date() },
): WorkSpreadsheetSheet {
  const range = normalizedWorkSpreadsheetAutoFilterRange(sheet.filter_select);
  if (!range || !entries.length) return sheet;
  const width = range.column[1] - range.column[0] + 1;
  const filter: NonNullable<WorkSpreadsheetSheet['filter']> = {};
  const owned = new Set<string>();
  for (const entry of entries) {
    const criteria = normalizeWorkSpreadsheetFilterCriteria(entry.criteria);
    if (
      !criteria ||
      !Number.isSafeInteger(entry.column) ||
      entry.column < 0 ||
      entry.column >= width
    ) {
      continue;
    }
    const column = range.column[0] + entry.column;
    const rowhidden =
      hiddenRowsForCriteria(sheet, range, column, criteria, context) ?? {};
    for (const row of Object.keys(rowhidden)) owned.add(row);
    filter[String(entry.column)] = filterColumnState(
      range,
      column,
      criteria,
      rowhidden,
    );
  }
  if (!Object.keys(filter).length) return sheet;
  const manuallyHiddenRows = new Set(
    Object.keys(sheet.config?.rowhidden ?? {}).filter((row) => !owned.has(row)),
  );
  return sheetWithFilterState(sheet, filter, manuallyHiddenRows);
}

export function normalizeWorkSpreadsheetFilterCriteria(
  criteria: unknown,
): WorkSpreadsheetFilterCriteria | null {
  if (!criteria || typeof criteria !== 'object' || Array.isArray(criteria)) {
    return null;
  }
  const value = criteria as Record<string, unknown>;
  const type = value.type;
  if (type === 'blanks' || type === 'non-blanks') return { type };
  if (type === 'values') {
    if (
      !Array.isArray(value.values) ||
      typeof value.includeBlanks !== 'boolean'
    ) {
      return null;
    }
    const values = value.values.filter(
      (item): item is string => typeof item === 'string',
    );
    if (
      values.length !== value.values.length ||
      (!values.length && !value.includeBlanks)
    ) {
      return null;
    }
    return {
      type,
      values: [...new Set(values)],
      includeBlanks: value.includeBlanks,
    };
  }
  if (type === 'between' || type === 'not-between') {
    return typeof value.lower === 'string' && typeof value.upper === 'string'
      ? { type, lower: value.lower, upper: value.upper }
      : null;
  }
  if (type === 'compound') {
    if (
      (value.conjunction !== 'and' && value.conjunction !== 'or') ||
      !Array.isArray(value.conditions) ||
      value.conditions.length !== 2
    ) {
      return null;
    }
    const first = normalizeWorkSpreadsheetCustomFilterCondition(
      value.conditions[0],
    );
    const second = normalizeWorkSpreadsheetCustomFilterCondition(
      value.conditions[1],
    );
    return first && second
      ? {
          type,
          conjunction: value.conjunction,
          conditions: [first, second],
        }
      : null;
  }
  if (type === 'top' || type === 'bottom') {
    return Number.isSafeInteger(value.count) &&
      Number(value.count) >= 1 &&
      Number(value.count) <= 500
      ? { type, count: Number(value.count) }
      : null;
  }
  if (type === 'top-percent' || type === 'bottom-percent') {
    return Number.isSafeInteger(value.percent) &&
      Number(value.percent) >= 1 &&
      Number(value.percent) <= 100
      ? { type, percent: Number(value.percent) }
      : null;
  }
  if (type === 'dynamic') {
    return typeof value.kind === 'string' &&
      DYNAMIC_FILTER_KINDS.has(value.kind)
      ? { type, kind: value.kind as WorkSpreadsheetDynamicFilter }
      : null;
  }
  return normalizeWorkSpreadsheetCustomFilterCondition(value);
}

function normalizeWorkSpreadsheetCustomFilterCondition(
  condition: unknown,
): WorkSpreadsheetCustomFilterCondition | null {
  if (!condition || typeof condition !== 'object' || Array.isArray(condition)) {
    return null;
  }
  const value = condition as Record<string, unknown>;
  if (
    typeof value.type !== 'string' ||
    !SINGLE_VALUE_FILTER_TYPES.has(value.type) ||
    typeof value.value !== 'string'
  ) {
    return null;
  }
  if (
    (value.type === 'matches-wildcard' ||
      value.type === 'does-not-match-wildcard') &&
    !workSpreadsheetFilterTextIsBounded(value.value)
  ) {
    return null;
  }
  return {
    type: value.type,
    value: value.value,
  } as WorkSpreadsheetCustomFilterCondition;
}

function sheetWithFilterState(
  sheet: WorkSpreadsheetSheet,
  filter: NonNullable<WorkSpreadsheetSheet['filter']>,
  manuallyHiddenRows: ReadonlySet<string>,
): WorkSpreadsheetSheet {
  const owned = filterHiddenRows(filter);
  const overlap = [...manuallyHiddenRows]
    .filter((row) => owned.has(row))
    .sort(compareHiddenRowKeys);
  const rowhidden: Record<string, 0> = {};
  for (const row of manuallyHiddenRows) rowhidden[row] = 0;
  for (const row of owned) rowhidden[row] = 0;
  return {
    ...sheet,
    filter: filterWithManualHiddenOverlap(filter, overlap),
    config: { ...(sheet.config ?? {}), rowhidden },
  };
}

function filterWithManualHiddenOverlap(
  filter: NonNullable<WorkSpreadsheetSheet['filter']>,
  manualHiddenOverlap: string[],
): NonNullable<WorkSpreadsheetSheet['filter']> {
  return Object.fromEntries(
    Object.entries(filter).map(([column, entry]) => {
      const payload = criteriaPayloadFromEntry(entry);
      if (!payload || !entry || typeof entry !== 'object') {
        return [column, entry];
      }
      return [
        column,
        {
          ...entry,
          caljs: {
            ...payload,
            manualHiddenOverlap: [...manualHiddenOverlap],
          },
        },
      ];
    }),
  );
}

function filterColumnState(
  range: WorkSpreadsheetAutoFilterRange,
  column: number,
  criteria: WorkSpreadsheetFilterCriteria,
  rowhidden: Record<string, 0>,
): WorkSpreadsheetAutoFilterColumnState {
  return {
    caljs: {
      criteria,
      kind: AUTO_FILTER_CRITERIA_KIND,
      version: AUTO_FILTER_CRITERIA_VERSION,
    },
    cindex: column,
    edc: range.column[1],
    edr: range.row[1],
    optionstate: true,
    rowhidden,
    stc: range.column[0],
    str: range.row[0],
  };
}

function hiddenRowsForCriteria(
  sheet: WorkSpreadsheetSheet,
  range: WorkSpreadsheetAutoFilterRange,
  column: number,
  criteria: WorkSpreadsheetFilterCriteria,
  context: WorkSpreadsheetDynamicFilterContext,
): Record<string, 0> | null {
  if (
    criteria.type === 'top' ||
    criteria.type === 'top-percent' ||
    criteria.type === 'bottom' ||
    criteria.type === 'bottom-percent'
  ) {
    return hiddenRowsForRankCriteria(sheet, range, column, criteria);
  }
  if (criteria.type === 'dynamic') {
    return hiddenRowsForDynamicCriteria(
      sheet,
      range,
      column,
      criteria.kind,
      context,
    );
  }
  const matches = filterMatcher(criteria);
  if (!matches) return null;
  const cellAt = sheetCellReader(sheet);
  const hidden: Record<string, 0> = {};
  for (let row = range.row[0] + 1; row <= range.row[1]; row += 1) {
    if (!matches(cellAt(row, column))) hidden[String(row)] = 0;
  }
  return hidden;
}

function hiddenRowsForDynamicCriteria(
  sheet: WorkSpreadsheetSheet,
  range: WorkSpreadsheetAutoFilterRange,
  column: number,
  kind: WorkSpreadsheetDynamicFilter,
  context: WorkSpreadsheetDynamicFilterContext,
): Record<string, 0> | null {
  const cellAt = sheetCellReader(sheet);
  const cells = {
    *[Symbol.iterator](): Iterator<Cell | null> {
      for (let row = range.row[0] + 1; row <= range.row[1]; row += 1) {
        yield cellAt(row, column);
      }
    },
  };
  const matches = workSpreadsheetDynamicFilterMatcher(kind, cells, context);
  if (!matches) return null;
  const hidden: Record<string, 0> = {};
  for (let row = range.row[0] + 1; row <= range.row[1]; row += 1) {
    if (!matches(cellAt(row, column))) hidden[String(row)] = 0;
  }
  return hidden;
}

function hiddenRowsForRankCriteria(
  sheet: WorkSpreadsheetSheet,
  range: WorkSpreadsheetAutoFilterRange,
  column: number,
  criteria: Extract<
    WorkSpreadsheetFilterCriteria,
    {
      type: 'top' | 'top-percent' | 'bottom' | 'bottom-percent';
    }
  >,
): Record<string, 0> {
  const cellAt = sheetCellReader(sheet);
  const values: number[] = [];
  for (let row = range.row[0] + 1; row <= range.row[1]; row += 1) {
    const value = numericCellValue(cellAt(row, column));
    if (value !== null) values.push(value);
  }
  const top = criteria.type === 'top' || criteria.type === 'top-percent';
  values.sort((left, right) => (top ? right - left : left - right));
  const requested =
    criteria.type === 'top' || criteria.type === 'bottom'
      ? criteria.count
      : Math.ceil((values.length * criteria.percent) / 100);
  const boundary = values[Math.min(requested, values.length) - 1];
  const hidden: Record<string, 0> = {};
  for (let row = range.row[0] + 1; row <= range.row[1]; row += 1) {
    const value = numericCellValue(cellAt(row, column));
    if (
      value === null ||
      boundary === undefined ||
      (top ? value < boundary : value > boundary)
    ) {
      hidden[String(row)] = 0;
    }
  }
  return hidden;
}

function filterMatcher(
  criteria: WorkSpreadsheetFilterCriteria,
): ((cell: Cell | null) => boolean) | null {
  if (criteria.type === 'blanks') return cellIsBlank;
  if (criteria.type === 'non-blanks') return (cell) => !cellIsBlank(cell);
  if (criteria.type === 'values') {
    const values = new Set(criteria.values.map(comparableText));
    return (cell) => {
      const value = cellValue(cell);
      if (value === null) return criteria.includeBlanks;
      return values.has(comparableText(String(value)));
    };
  }
  if (criteria.type === 'between' || criteria.type === 'not-between') {
    return (cell) => {
      const lower = compareCell(cell, criteria.lower);
      const upper = compareCell(cell, criteria.upper);
      const between =
        lower !== null && upper !== null && lower >= 0 && upper <= 0;
      return criteria.type === 'between' ? between : !between;
    };
  }
  if (criteria.type === 'compound') {
    const first = customFilterMatcher(criteria.conditions[0]);
    const second = customFilterMatcher(criteria.conditions[1]);
    return criteria.conjunction === 'and'
      ? (cell) => first(cell) && second(cell)
      : (cell) => first(cell) || second(cell);
  }
  if (
    criteria.type === 'top' ||
    criteria.type === 'top-percent' ||
    criteria.type === 'bottom' ||
    criteria.type === 'bottom-percent' ||
    criteria.type === 'dynamic'
  ) {
    return null;
  }
  return customFilterMatcher(criteria);
}

function customFilterMatcher(
  criteria: WorkSpreadsheetCustomFilterCondition,
): (cell: Cell | null) => boolean {
  if (
    criteria.type === 'matches-wildcard' ||
    criteria.type === 'does-not-match-wildcard'
  ) {
    const matchesWildcard = workSpreadsheetWildcardMatcher(criteria.value);
    const matches = (cell: Cell | null) => {
      const value = cellValue(cell);
      return value !== null && matchesWildcard(String(value));
    };
    return criteria.type === 'matches-wildcard'
      ? matches
      : (cell) => !matches(cell);
  }
  const expected = criteria.value;
  if (criteria.type === 'contains') {
    return (cell) => cellText(cell).includes(comparableText(expected));
  }
  if (criteria.type === 'does-not-contain') {
    return (cell) => !cellText(cell).includes(comparableText(expected));
  }
  if (criteria.type === 'begins-with') {
    return (cell) => cellText(cell).startsWith(comparableText(expected));
  }
  if (criteria.type === 'does-not-begin-with') {
    return (cell) => !cellText(cell).startsWith(comparableText(expected));
  }
  if (criteria.type === 'ends-with') {
    return (cell) => cellText(cell).endsWith(comparableText(expected));
  }
  if (criteria.type === 'does-not-end-with') {
    return (cell) => !cellText(cell).endsWith(comparableText(expected));
  }
  return (cell) => {
    const order = compareCell(cell, expected);
    if (order === null) return false;
    if (criteria.type === 'equals') return order === 0;
    if (criteria.type === 'not-equals') return order !== 0;
    if (criteria.type === 'greater-than') return order > 0;
    if (criteria.type === 'greater-than-or-equal') return order >= 0;
    if (criteria.type === 'less-than') return order < 0;
    return order <= 0;
  };
}

function compareCell(cell: Cell | null, expected: string): number | null {
  const value = cellValue(cell);
  if (value === null) return null;
  const numericExpected = finiteNumber(expected);
  if (typeof value === 'number' && numericExpected !== null) {
    return value - numericExpected;
  }
  return comparableText(String(value)).localeCompare(
    comparableText(expected),
    'zh-CN',
    { numeric: true, sensitivity: 'base' },
  );
}

function cellIsBlank(cell: Cell | null): boolean {
  return cellValue(cell) === null;
}

function cellText(cell: Cell | null): string {
  const value = cellValue(cell);
  return value === null ? '' : comparableText(String(value));
}

function cellValue(cell: Cell | null): string | number | boolean | null {
  const value = cell?.v ?? cell?.m;
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (typeof value === 'string') return value === '' ? null : value;
  if (typeof value === 'boolean') return value;
  return null;
}

function numericCellValue(cell: Cell | null): number | null {
  const value = cellValue(cell);
  return typeof value === 'number' ? value : null;
}

function comparableText(value: string): string {
  return value.normalize('NFKC').toLocaleLowerCase('zh-CN');
}

function finiteNumber(value: string): number | null {
  if (!value.trim()) return null;
  const numeric = Number(value);
  return Number.isFinite(numeric) ? numeric : null;
}

function criteriaFromEntry(
  entry: unknown,
): WorkSpreadsheetFilterCriteria | null {
  const payload = criteriaPayloadFromEntry(entry);
  return payload
    ? normalizeWorkSpreadsheetFilterCriteria(payload.criteria)
    : null;
}

function criteriaPayloadFromEntry(
  entry: unknown,
): WorkSpreadsheetAutoFilterCriteriaPayload | null {
  if (!entry || typeof entry !== 'object') return null;
  const caljs = (entry as { caljs?: unknown }).caljs;
  if (!caljs || typeof caljs !== 'object') return null;
  const payload = caljs as Partial<WorkSpreadsheetAutoFilterCriteriaPayload>;
  if (
    payload.kind !== AUTO_FILTER_CRITERIA_KIND ||
    payload.version !== AUTO_FILTER_CRITERIA_VERSION
  ) {
    return null;
  }
  return normalizeWorkSpreadsheetFilterCriteria(payload.criteria)
    ? (payload as WorkSpreadsheetAutoFilterCriteriaPayload)
    : null;
}

function filterHiddenRows(filter: WorkSpreadsheetSheet['filter']): Set<string> {
  const rows = new Set<string>();
  for (const value of Object.values(filter ?? {})) {
    if (!value || typeof value !== 'object') continue;
    const rowhidden = (value as { rowhidden?: unknown }).rowhidden;
    if (!rowhidden || typeof rowhidden !== 'object') continue;
    for (const row of Object.keys(rowhidden)) rows.add(row);
  }
  return rows;
}

function sheetCellReader(
  sheet: WorkSpreadsheetSheet,
): (row: number, column: number) => Cell | null {
  if (sheet.data) {
    return (row, column) => sheet.data?.[row]?.[column] ?? null;
  }
  const cells = new Map(
    (sheet.celldata ?? []).map((entry) => [
      `${entry.r}:${entry.c}`,
      entry.v ?? null,
    ]),
  );
  return (row, column) => cells.get(`${row}:${column}`) ?? null;
}

function columnInRange(
  range: WorkSpreadsheetAutoFilterRange,
  column: number,
): boolean {
  return (
    Number.isSafeInteger(column) &&
    column >= range.column[0] &&
    column <= range.column[1]
  );
}

function finiteIndex(value: unknown): number | null {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0
    ? value
    : null;
}

function validHiddenRowKey(value: unknown): value is string {
  return (
    typeof value === 'string' &&
    /^(?:0|[1-9]\d*)$/.test(value) &&
    Number.isSafeInteger(Number(value))
  );
}

function compareHiddenRowKeys(left: string, right: string): number {
  return Number(left) - Number(right);
}
