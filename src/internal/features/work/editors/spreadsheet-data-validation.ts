import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetDataValidationErrorStyle,
  WorkSpreadsheetDataValidationItem,
  WorkSpreadsheetDataValidationRange,
  WorkSpreadsheetSheet,
} from '../work-types';
import {
  boundedSpreadsheetDataValidationText,
  normalizeSpreadsheetDataValidationErrorStyle,
  normalizeSpreadsheetDateValidationBoundary,
  SPREADSHEET_DATA_VALIDATION_ERROR_LIMIT,
  SPREADSHEET_DATA_VALIDATION_FORMULA_LIMIT,
  SPREADSHEET_DATA_VALIDATION_HINT_LIMIT,
  SPREADSHEET_DATA_VALIDATION_TITLE_LIMIT,
} from '../work-spreadsheet-data-validation';
import { formatSpreadsheetCellRanges } from '../work-spreadsheet-ranges';
import { canMutateSpreadsheetCellRanges } from './spreadsheet-cell-mutation-guard';
import {
  normalizeSpreadsheetCellRange,
  parseSpreadsheetCellRange,
  spreadsheetCellRangeArea,
  spreadsheetCellRangeContains,
  spreadsheetCellRangesIntersect,
  subtractSpreadsheetCellRange,
  type SpreadsheetCellRange,
} from './spreadsheet-cell-range';
import { resolveSpreadsheetGoToTarget } from './spreadsheet-go-to';
import { spreadsheetSheetBounds } from './spreadsheet-keyboard-navigation';
import {
  isSpreadsheetDependentListFormula,
  normalizeSpreadsheetDependentListFormula,
  resolveSpreadsheetDependentListReference,
} from './spreadsheet-data-validation-list';

export const MAX_SPREADSHEET_DATA_VALIDATION_CELLS = 10_000;

export type SpreadsheetDataValidationType =
  | 'custom'
  | 'date'
  | 'dropdown'
  | 'number'
  | 'number_integer'
  | 'text_length';

export type SpreadsheetDataValidationOperator =
  | 'between'
  | 'equal'
  | 'greaterOrEqualTo'
  | 'lessThan'
  | 'lessThanOrEqualTo'
  | 'moreThanThe'
  | 'noEarlierThan'
  | 'noLaterThan'
  | 'notBetween'
  | 'notEqualTo'
  | 'earlierThan'
  | 'laterThan';

export interface SpreadsheetDataValidationDialogValue {
  allowBlank: boolean;
  errorMessage: string;
  errorStyle: WorkSpreadsheetDataValidationErrorStyle;
  errorTitle: string;
  hintShow: boolean;
  hintTitle: string;
  hintValue: string;
  prohibitInput: boolean;
  showDropdownArrow: boolean;
  type: SpreadsheetDataValidationType;
  type2: SpreadsheetDataValidationOperator | '';
  value1: string;
  value2: string;
}

export interface SpreadsheetDataValidationTarget {
  activeCell: { row: number; column: number };
  ranges: readonly SpreadsheetCellRange[];
  sheetId: string;
}

export interface SpreadsheetDataValidationRequest
  extends SpreadsheetDataValidationTarget {
  value: SpreadsheetDataValidationDialogValue;
}

export interface SpreadsheetDataValidationDialogSource
  extends SpreadsheetDataValidationTarget {
  hasValidation: boolean;
  mixed: boolean;
  rangeReference: string;
  sheetName: string;
  value: SpreadsheetDataValidationDialogValue;
}

export type SpreadsheetDataValidationResult =
  | {
      item: WorkSpreadsheetDataValidationItem;
      ok: true;
      ranges: SpreadsheetCellRange[];
      sheet: WorkSpreadsheetSheet;
    }
  | {
      code: SpreadsheetDataValidationErrorCode;
      message: string;
      ok: false;
    };

export type SpreadsheetDataValidationErrorCode =
  | 'invalid-custom-formula'
  | 'invalid-date'
  | 'invalid-list-formula'
  | 'invalid-list-source'
  | 'invalid-number'
  | 'invalid-operator'
  | 'invalid-range'
  | 'invalid-text-length'
  | 'missing-value'
  | 'multiple-list-columns'
  | 'out-of-bounds'
  | 'protected-range'
  | 'range-too-large'
  | 'sheet-not-found'
  | 'value-order';

const defaultSpreadsheetDataValidationValue: SpreadsheetDataValidationDialogValue =
  {
    allowBlank: true,
    errorMessage: '',
    errorStyle: 'stop',
    errorTitle: '',
    hintShow: false,
    hintTitle: '',
    hintValue: '',
    prohibitInput: true,
    showDropdownArrow: true,
    type: 'dropdown',
    type2: '',
    value1: '',
    value2: '',
  };

const numericOperators: readonly SpreadsheetDataValidationOperator[] = [
  'between',
  'notBetween',
  'equal',
  'notEqualTo',
  'moreThanThe',
  'lessThan',
  'greaterOrEqualTo',
  'lessThanOrEqualTo',
];

const dateOperators: readonly SpreadsheetDataValidationOperator[] = [
  'between',
  'notBetween',
  'equal',
  'notEqualTo',
  'earlierThan',
  'noEarlierThan',
  'laterThan',
  'noLaterThan',
];

export function spreadsheetDataValidationOperators(
  type: SpreadsheetDataValidationType,
): readonly SpreadsheetDataValidationOperator[] {
  if (type === 'custom' || type === 'dropdown') return [];
  return type === 'date' ? dateOperators : numericOperators;
}

export function createSpreadsheetDataValidationDialogSource(
  content: WorkSpreadsheetContent,
  target: SpreadsheetDataValidationTarget,
): SpreadsheetDataValidationDialogSource | null {
  const resolved = resolveSpreadsheetDataValidationTarget(content, target);
  if (!resolved) return null;
  const { ranges, sheet } = resolved;
  if (spreadsheetDataValidationTargetError(sheet, ranges)) return null;
  const selection = spreadsheetDataValidationSelectionState(sheet, ranges);
  const active = spreadsheetDataValidationItemAt(
    sheet,
    target.activeCell.row,
    target.activeCell.column,
  );
  const value = spreadsheetDataValidationDialogValue(active);
  return {
    sheetId: target.sheetId,
    sheetName: sheet.name,
    ranges,
    activeCell: { ...target.activeCell },
    rangeReference: formatSpreadsheetCellRanges(ranges),
    hasValidation: selection.hasValidation,
    mixed: selection.mixed,
    value,
  };
}

export function validateSpreadsheetDataValidationRequest(
  content: WorkSpreadsheetContent,
  request: SpreadsheetDataValidationRequest,
): SpreadsheetDataValidationResult {
  const resolved = resolveSpreadsheetDataValidationTarget(content, request);
  if (!resolved) {
    return spreadsheetDataValidationError('invalid-range');
  }
  const targetError = spreadsheetDataValidationTargetError(
    resolved.sheet,
    resolved.ranges,
  );
  if (targetError) return targetError;
  const item = normalizeSpreadsheetDataValidationItem(
    content,
    resolved.sheet,
    resolved.ranges,
    request.value,
  );
  if (!item.ok) return item;
  return {
    ok: true,
    item: item.item,
    ranges: resolved.ranges,
    sheet: resolved.sheet,
  };
}

export function applySpreadsheetDataValidation(
  content: WorkSpreadsheetContent,
  request: SpreadsheetDataValidationRequest,
): WorkSpreadsheetContent | null {
  const validation = validateSpreadsheetDataValidationRequest(content, request);
  if (!validation.ok) return null;
  const cleared = clearSpreadsheetDataValidationRanges(
    validation.sheet,
    validation.ranges,
  );
  const nextRanges = appendSpreadsheetDataValidationRange(
    cleared.dataValidationRanges ?? [],
    validation.ranges,
    validation.item,
  );
  const nextSheet: WorkSpreadsheetSheet = {
    ...cleared,
    dataValidationRanges: nextRanges,
  };
  return replaceSpreadsheetDataValidationSheet(content, nextSheet);
}

export function removeSpreadsheetDataValidation(
  content: WorkSpreadsheetContent,
  target: SpreadsheetDataValidationTarget,
): WorkSpreadsheetContent | null {
  const resolved = resolveSpreadsheetDataValidationTarget(content, target);
  if (!resolved) return null;
  const targetError = spreadsheetDataValidationTargetError(
    resolved.sheet,
    resolved.ranges,
  );
  if (
    targetError ||
    !spreadsheetSelectionHasDataValidation(resolved.sheet, resolved.ranges)
  ) {
    return null;
  }
  return replaceSpreadsheetDataValidationSheet(
    content,
    clearSpreadsheetDataValidationRanges(resolved.sheet, resolved.ranges),
  );
}

export function canRemoveSpreadsheetDataValidation(
  content: WorkSpreadsheetContent,
  target: SpreadsheetDataValidationTarget,
): boolean {
  const resolved = resolveSpreadsheetDataValidationTarget(content, target);
  return Boolean(
    resolved &&
      !spreadsheetDataValidationTargetError(resolved.sheet, resolved.ranges) &&
      spreadsheetSelectionHasDataValidation(resolved.sheet, resolved.ranges),
  );
}

export function spreadsheetDataValidationFailureMessage(
  content: WorkSpreadsheetContent,
  request: SpreadsheetDataValidationRequest,
): string | null {
  const validation = validateSpreadsheetDataValidationRequest(content, request);
  return validation.ok ? null : validation.message;
}

function resolveSpreadsheetDataValidationTarget(
  content: WorkSpreadsheetContent,
  target: SpreadsheetDataValidationTarget,
): { ranges: SpreadsheetCellRange[]; sheet: WorkSpreadsheetSheet } | null {
  const sheet = content.sheets.find(
    (candidate) => candidate.id === target.sheetId,
  );
  if (!sheet) return null;
  const ranges = normalizeSpreadsheetDataValidationRanges(target.ranges);
  if (!ranges.length) return null;
  if (
    !ranges.some((range) =>
      spreadsheetCellRangeContains(
        range,
        target.activeCell.row,
        target.activeCell.column,
      ),
    )
  ) {
    return null;
  }
  return { sheet, ranges };
}

function spreadsheetDataValidationTargetError(
  sheet: WorkSpreadsheetSheet,
  ranges: readonly SpreadsheetCellRange[],
): Extract<SpreadsheetDataValidationResult, { ok: false }> | null {
  const bounds = spreadsheetSheetBounds(sheet);
  if (
    ranges.some(
      (range) =>
        range.row[1] > bounds.lastRow || range.column[1] > bounds.lastColumn,
    )
  ) {
    return spreadsheetDataValidationError('out-of-bounds');
  }
  const area = ranges.reduce(
    (total, range) => total + spreadsheetCellRangeArea(range),
    0,
  );
  if (area > MAX_SPREADSHEET_DATA_VALIDATION_CELLS) {
    return spreadsheetDataValidationError('range-too-large');
  }
  if (!canMutateSpreadsheetCellRanges(sheet, ranges)) {
    return spreadsheetDataValidationError('protected-range');
  }
  return null;
}

function normalizeSpreadsheetDataValidationItem(
  content: WorkSpreadsheetContent,
  sheet: WorkSpreadsheetSheet,
  ranges: readonly SpreadsheetCellRange[],
  value: SpreadsheetDataValidationDialogValue,
):
  | { item: WorkSpreadsheetDataValidationItem; ok: true }
  | Extract<SpreadsheetDataValidationResult, { ok: false }> {
  const operators = spreadsheetDataValidationOperators(value.type);
  const type2 =
    value.type === 'dropdown' || value.type === 'custom' ? '' : value.type2;
  if (
    value.type !== 'dropdown' &&
    value.type !== 'custom' &&
    !operators.includes(type2 as SpreadsheetDataValidationOperator)
  ) {
    return spreadsheetDataValidationError('invalid-operator');
  }
  const normalizedValues = normalizeSpreadsheetDataValidationValues(
    content,
    sheet,
    ranges,
    value.type,
    type2,
    value.value1,
    value.value2,
  );
  if (!normalizedValues.ok) return normalizedValues;
  return {
    ok: true,
    item: {
      type: value.type,
      type2,
      rangeTxt: formatSpreadsheetCellRanges(
        ranges.map(cloneSpreadsheetCellRange),
      ),
      value1: normalizedValues.value1,
      value2: normalizedValues.value2,
      validity: '',
      remote: false,
      allowBlank: Boolean(value.allowBlank),
      showDropdownArrow:
        value.type === 'dropdown' ? Boolean(value.showDropdownArrow) : true,
      prohibitInput: Boolean(value.prohibitInput),
      errorStyle: normalizeSpreadsheetDataValidationErrorStyle(
        value.errorStyle,
      ),
      errorTitle: boundedSpreadsheetDataValidationText(
        value.errorTitle,
        SPREADSHEET_DATA_VALIDATION_TITLE_LIMIT,
      ),
      errorMessage: boundedSpreadsheetDataValidationText(
        value.errorMessage,
        SPREADSHEET_DATA_VALIDATION_ERROR_LIMIT,
      ),
      hintShow: Boolean(value.hintShow),
      hintTitle: boundedSpreadsheetDataValidationText(
        value.hintTitle,
        SPREADSHEET_DATA_VALIDATION_TITLE_LIMIT,
      ),
      hintValue: boundedSpreadsheetDataValidationText(
        value.hintValue,
        SPREADSHEET_DATA_VALIDATION_HINT_LIMIT,
      ),
      checked: false,
    },
  };
}

function normalizeSpreadsheetDataValidationValues(
  content: WorkSpreadsheetContent,
  sheet: WorkSpreadsheetSheet,
  ranges: readonly SpreadsheetCellRange[],
  type: SpreadsheetDataValidationType,
  type2: SpreadsheetDataValidationOperator | '',
  rawValue1: string,
  rawValue2: string,
):
  | { ok: true; value1: string; value2: string }
  | Extract<SpreadsheetDataValidationResult, { ok: false }> {
  const value1 = rawValue1.trim();
  const value2 = rawValue2.trim();
  if (!value1) return spreadsheetDataValidationError('missing-value');
  if (type === 'dropdown') {
    return normalizeSpreadsheetListValidation(content, sheet, ranges, value1);
  }
  if (type === 'custom') {
    const formula = value1.replace(/^=/, '').trim();
    if (
      !formula ||
      Array.from(formula).length > SPREADSHEET_DATA_VALIDATION_FORMULA_LIMIT ||
      /[\u0000-\u001f\u007f]/u.test(formula)
    ) {
      return spreadsheetDataValidationError('invalid-custom-formula');
    }
    return { ok: true, value1: formula, value2: '' };
  }
  const needsSecond = type2 === 'between' || type2 === 'notBetween';
  if (needsSecond && !value2) {
    return spreadsheetDataValidationError('missing-value');
  }
  if (type === 'date') {
    const first = normalizeSpreadsheetDateValidationBoundary(value1);
    const second = needsSecond
      ? normalizeSpreadsheetDateValidationBoundary(value2)
      : null;
    if (first === null || (needsSecond && second === null)) {
      return spreadsheetDataValidationError('invalid-date');
    }
    if (second !== null && second < first) {
      return spreadsheetDataValidationError('value-order');
    }
    return { ok: true, value1: first, value2: second ?? '' };
  }
  const first = Number(value1);
  const second = needsSecond ? Number(value2) : null;
  if (!Number.isFinite(first) || (needsSecond && !Number.isFinite(second))) {
    return spreadsheetDataValidationError(
      type === 'text_length' ? 'invalid-text-length' : 'invalid-number',
    );
  }
  if (
    (type === 'number_integer' && !Number.isInteger(first)) ||
    (type === 'text_length' && (!Number.isInteger(first) || first < 0)) ||
    (second !== null &&
      ((type === 'number_integer' && !Number.isInteger(second)) ||
        (type === 'text_length' && (!Number.isInteger(second) || second < 0))))
  ) {
    return spreadsheetDataValidationError(
      type === 'text_length' ? 'invalid-text-length' : 'invalid-number',
    );
  }
  if (second !== null && second < first) {
    return spreadsheetDataValidationError('value-order');
  }
  return {
    ok: true,
    value1: String(first),
    value2: second === null ? '' : String(second),
  };
}

function normalizeSpreadsheetListValidation(
  content: WorkSpreadsheetContent,
  sheet: WorkSpreadsheetSheet,
  ranges: readonly SpreadsheetCellRange[],
  value: string,
):
  | { ok: true; value1: string; value2: string }
  | Extract<SpreadsheetDataValidationResult, { ok: false }> {
  if (isSpreadsheetDependentListFormula(value)) {
    const normalized = normalizeSpreadsheetDependentListFormula(value);
    if (!normalized.ok || !normalized.formula) {
      return spreadsheetDataValidationError('invalid-list-formula');
    }
    for (const range of ranges) {
      const resolved = resolveSpreadsheetDependentListReference(
        content,
        sheet.id ?? '',
        range.row[0],
        range.column[0],
        range.row[0],
        range.column[0],
        normalized.formula,
      );
      if (!resolved.ok) {
        return spreadsheetDataValidationError('invalid-list-formula');
      }
    }
    return { ok: true, value1: normalized.formula, value2: '' };
  }
  const looksLikeReference =
    value.startsWith('=') ||
    value.includes('!') ||
    /^\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?$/i.test(value) ||
    (content.namedRanges ?? []).some(
      (range) =>
        range.name.trim().toLocaleLowerCase() === value.toLocaleLowerCase(),
    );
  if (looksLikeReference) {
    const resolved = resolveSpreadsheetGoToTarget(
      content,
      sheet.id ?? '',
      value,
    );
    if (!resolved.ok)
      return spreadsheetDataValidationError('invalid-list-source');
    const { range } = resolved.target;
    if (range.row[0] !== range.row[1] && range.column[0] !== range.column[1]) {
      return spreadsheetDataValidationError('multiple-list-columns');
    }
    return {
      ok: true,
      value1: resolved.target.displayReference,
      value2: '',
    };
  }
  const options = value.split(',').map((option) => option.trim());
  if (
    options.some((option) => !option || option.length > 255) ||
    options.join(',').length > 255
  ) {
    return spreadsheetDataValidationError('invalid-list-source');
  }
  return { ok: true, value1: options.join(','), value2: '' };
}

function spreadsheetDataValidationSelectionState(
  sheet: WorkSpreadsheetSheet,
  ranges: readonly SpreadsheetCellRange[],
): { hasValidation: boolean; mixed: boolean } {
  let firstSignature: string | null | undefined;
  let hasValidation = false;
  for (const range of ranges) {
    for (let row = range.row[0]; row <= range.row[1]; row += 1) {
      for (
        let column = range.column[0];
        column <= range.column[1];
        column += 1
      ) {
        const item = spreadsheetDataValidationItemAt(sheet, row, column);
        const signature = item
          ? spreadsheetDataValidationSignature(item)
          : null;
        hasValidation ||= Boolean(item);
        if (firstSignature === undefined) firstSignature = signature;
        else if (firstSignature !== signature)
          return { hasValidation, mixed: true };
      }
    }
  }
  return { hasValidation, mixed: false };
}

export function spreadsheetDataValidationItemAt(
  sheet: WorkSpreadsheetSheet,
  row: number,
  column: number,
): WorkSpreadsheetDataValidationItem | null {
  const direct = spreadsheetDataValidationItem(
    (sheet.dataVerification as Record<string, unknown> | undefined)?.[
      `${row}_${column}`
    ],
  );
  if (direct) return direct;
  for (
    let index = (sheet.dataValidationRanges?.length ?? 0) - 1;
    index >= 0;
    index -= 1
  ) {
    const entry = sheet.dataValidationRanges?.[index];
    if (!entry) continue;
    const matches = entry.ranges.some((candidate) => {
      const range = parseSpreadsheetCellRange(candidate);
      return range ? spreadsheetCellRangeContains(range, row, column) : false;
    });
    if (matches) return spreadsheetDataValidationItem(entry.item);
  }
  return null;
}

function spreadsheetDataValidationDialogValue(
  value: unknown,
): SpreadsheetDataValidationDialogValue {
  const item = spreadsheetDataValidationItem(value);
  if (!item || !spreadsheetDataValidationType(item.type)) {
    return { ...defaultSpreadsheetDataValidationValue };
  }
  const type = item.type as SpreadsheetDataValidationType;
  const operators = spreadsheetDataValidationOperators(type);
  const type2 = operators.includes(
    item.type2 as SpreadsheetDataValidationOperator,
  )
    ? (item.type2 as SpreadsheetDataValidationOperator)
    : (operators[0] ?? '');
  return {
    type,
    type2,
    value1:
      type === 'date'
        ? (normalizeSpreadsheetDateValidationBoundary(item.value1) ??
          item.value1)
        : item.value1,
    value2:
      type === 'date' && item.value2
        ? (normalizeSpreadsheetDateValidationBoundary(item.value2) ??
          item.value2)
        : item.value2,
    prohibitInput: item.prohibitInput,
    allowBlank: item.allowBlank ?? true,
    showDropdownArrow: item.showDropdownArrow ?? true,
    errorStyle: item.errorStyle ?? 'stop',
    errorTitle: item.errorTitle ?? '',
    errorMessage: item.errorMessage ?? '',
    hintShow: item.hintShow,
    hintTitle: item.hintTitle ?? '',
    hintValue: item.hintValue,
  };
}

function spreadsheetDataValidationType(
  value: string,
): value is SpreadsheetDataValidationType {
  return [
    'custom',
    'date',
    'dropdown',
    'number',
    'number_integer',
    'text_length',
  ].includes(value);
}

function spreadsheetDataValidationItem(
  value: unknown,
): WorkSpreadsheetDataValidationItem | null {
  if (!isRecord(value) || typeof value.type !== 'string') return null;
  return {
    type: value.type,
    type2: typeof value.type2 === 'string' ? value.type2 : '',
    rangeTxt: typeof value.rangeTxt === 'string' ? value.rangeTxt : '',
    value1:
      typeof value.value1 === 'string'
        ? value.value1
        : String(value.value1 ?? ''),
    value2:
      typeof value.value2 === 'string'
        ? value.value2
        : String(value.value2 ?? ''),
    validity: typeof value.validity === 'string' ? value.validity : '',
    remote: Boolean(value.remote),
    allowBlank: typeof value.allowBlank === 'boolean' ? value.allowBlank : true,
    showDropdownArrow:
      typeof value.showDropdownArrow === 'boolean'
        ? value.showDropdownArrow
        : true,
    prohibitInput: Boolean(value.prohibitInput),
    errorStyle: normalizeSpreadsheetDataValidationErrorStyle(value.errorStyle),
    errorTitle: boundedSpreadsheetDataValidationText(
      value.errorTitle,
      SPREADSHEET_DATA_VALIDATION_TITLE_LIMIT,
    ),
    errorMessage: boundedSpreadsheetDataValidationText(
      value.errorMessage,
      SPREADSHEET_DATA_VALIDATION_ERROR_LIMIT,
    ),
    hintShow: Boolean(value.hintShow),
    hintTitle: boundedSpreadsheetDataValidationText(
      value.hintTitle,
      SPREADSHEET_DATA_VALIDATION_TITLE_LIMIT,
    ),
    hintValue: boundedSpreadsheetDataValidationText(
      value.hintValue,
      SPREADSHEET_DATA_VALIDATION_HINT_LIMIT,
    ),
    checked: Boolean(value.checked),
  };
}

function spreadsheetSelectionHasDataValidation(
  sheet: WorkSpreadsheetSheet,
  ranges: readonly SpreadsheetCellRange[],
): boolean {
  const direct = sheet.dataVerification as Record<string, unknown> | undefined;
  if (
    direct &&
    Object.keys(direct).some((key) => {
      const coordinate = spreadsheetValidationCoordinate(key);
      return coordinate
        ? ranges.some((range) =>
            spreadsheetCellRangeContains(
              range,
              coordinate.row,
              coordinate.column,
            ),
          )
        : false;
    })
  ) {
    return true;
  }
  return (sheet.dataValidationRanges ?? []).some((entry) =>
    entry.ranges.some((candidate) => {
      const range = parseSpreadsheetCellRange(candidate);
      return range
        ? ranges.some((selected) =>
            spreadsheetCellRangesIntersect(range, selected),
          )
        : false;
    }),
  );
}

function clearSpreadsheetDataValidationRanges(
  sheet: WorkSpreadsheetSheet,
  selected: readonly SpreadsheetCellRange[],
): WorkSpreadsheetSheet {
  const direct = {
    ...((sheet.dataVerification ?? {}) as Record<string, unknown>),
  };
  for (const key of Object.keys(direct)) {
    const coordinate = spreadsheetValidationCoordinate(key);
    if (
      coordinate &&
      selected.some((range) =>
        spreadsheetCellRangeContains(range, coordinate.row, coordinate.column),
      )
    ) {
      delete direct[key];
    }
  }
  const compact = (sheet.dataValidationRanges ?? []).flatMap((entry) => {
    const remaining = entry.ranges.flatMap((candidate) => {
      const parsed = parseSpreadsheetCellRange(candidate);
      if (!parsed) return [candidate];
      return selected.reduce<SpreadsheetCellRange[]>(
        (ranges, removed) =>
          ranges.flatMap((range) =>
            subtractSpreadsheetCellRange(range, removed),
          ),
        [parsed],
      );
    });
    return remaining.length ? [{ ...entry, ranges: remaining }] : [];
  });
  const next: WorkSpreadsheetSheet = {
    ...sheet,
    dataVerification: Object.keys(direct).length
      ? (direct as WorkSpreadsheetSheet['dataVerification'])
      : undefined,
    dataValidationRanges: compact.length ? compact : undefined,
  };
  return next;
}

function appendSpreadsheetDataValidationRange(
  existing: readonly WorkSpreadsheetDataValidationRange[],
  ranges: readonly SpreadsheetCellRange[],
  item: WorkSpreadsheetDataValidationItem,
): WorkSpreadsheetDataValidationRange[] {
  const signature = spreadsheetDataValidationSignature(item);
  const sameIndex = existing.findIndex(
    (entry) => spreadsheetDataValidationSignature(entry.item) === signature,
  );
  if (sameIndex < 0) {
    return [
      ...existing,
      {
        ranges: ranges.map(cloneSpreadsheetCellRange),
        item: { ...item },
      },
    ];
  }
  return existing.map((entry, index) =>
    index === sameIndex
      ? {
          ...entry,
          ranges: normalizeSpreadsheetDataValidationRanges([
            ...entry.ranges,
            ...ranges,
          ]),
        }
      : entry,
  );
}

function normalizeSpreadsheetDataValidationRanges(
  values: readonly SpreadsheetCellRange[],
): SpreadsheetCellRange[] {
  const result: SpreadsheetCellRange[] = [];
  for (const value of values) {
    const range = normalizeSpreadsheetCellRange(value);
    if (!range) continue;
    let remaining = [range];
    for (const existing of result) {
      remaining = remaining.flatMap((candidate) =>
        subtractSpreadsheetCellRange(candidate, existing),
      );
    }
    result.push(...remaining);
  }
  return mergeAdjacentSpreadsheetDataValidationRanges(result);
}

function mergeAdjacentSpreadsheetDataValidationRanges(
  values: readonly SpreadsheetCellRange[],
): SpreadsheetCellRange[] {
  const ranges = values.map(cloneSpreadsheetCellRange);
  let merged = true;
  while (merged) {
    merged = false;
    outer: for (let leftIndex = 0; leftIndex < ranges.length; leftIndex += 1) {
      for (
        let rightIndex = leftIndex + 1;
        rightIndex < ranges.length;
        rightIndex += 1
      ) {
        const combined = combineSpreadsheetDataValidationRanges(
          ranges[leftIndex],
          ranges[rightIndex],
        );
        if (!combined) continue;
        ranges.splice(rightIndex, 1);
        ranges[leftIndex] = combined;
        merged = true;
        break outer;
      }
    }
  }
  return ranges;
}

function combineSpreadsheetDataValidationRanges(
  left: SpreadsheetCellRange | undefined,
  right: SpreadsheetCellRange | undefined,
): SpreadsheetCellRange | null {
  if (!left || !right) return null;
  if (
    left.row[0] === right.row[0] &&
    left.row[1] === right.row[1] &&
    (left.column[1] + 1 === right.column[0] ||
      right.column[1] + 1 === left.column[0])
  ) {
    return {
      row: [...left.row],
      column: [
        Math.min(left.column[0], right.column[0]),
        Math.max(left.column[1], right.column[1]),
      ],
    };
  }
  if (
    left.column[0] === right.column[0] &&
    left.column[1] === right.column[1] &&
    (left.row[1] + 1 === right.row[0] || right.row[1] + 1 === left.row[0])
  ) {
    return {
      row: [
        Math.min(left.row[0], right.row[0]),
        Math.max(left.row[1], right.row[1]),
      ],
      column: [...left.column],
    };
  }
  return null;
}

function spreadsheetValidationCoordinate(
  value: string,
): { column: number; row: number } | null {
  const match = /^(\d+)_(\d+)$/.exec(value);
  return match ? { row: Number(match[1]), column: Number(match[2]) } : null;
}

function spreadsheetDataValidationSignature(value: unknown): string {
  const item = spreadsheetDataValidationItem(value);
  return item
    ? JSON.stringify({
        type: item.type,
        type2: item.type2,
        value1: item.value1,
        value2: item.value2,
        allowBlank: item.allowBlank,
        showDropdownArrow: item.showDropdownArrow,
        prohibitInput: item.prohibitInput,
        errorStyle: item.errorStyle,
        errorTitle: item.errorTitle,
        errorMessage: item.errorMessage,
        hintShow: item.hintShow,
        hintTitle: item.hintTitle,
        hintValue: item.hintValue,
      })
    : 'invalid';
}

function replaceSpreadsheetDataValidationSheet(
  content: WorkSpreadsheetContent,
  sheet: WorkSpreadsheetSheet,
): WorkSpreadsheetContent {
  return {
    ...content,
    sheets: content.sheets.map((candidate) =>
      candidate.id === sheet.id ? sheet : candidate,
    ),
  };
}

function cloneSpreadsheetCellRange(
  range: SpreadsheetCellRange,
): SpreadsheetCellRange {
  return { row: [...range.row], column: [...range.column] };
}

function spreadsheetDataValidationError(
  code: SpreadsheetDataValidationErrorCode,
): Extract<SpreadsheetDataValidationResult, { ok: false }> {
  const messages: Record<SpreadsheetDataValidationErrorCode, string> = {
    'invalid-custom-formula':
      '自定义公式必须是 255 个字符以内的本地公式，且不能包含控制字符。',
    'invalid-date': '请输入有效的日期、Excel 日期序号或 DATE(...) 表达式。',
    'invalid-list-formula':
      '动态下拉公式必须是本地 =INDIRECT(单元格或文本拼接)，并解析为当前工作簿内的一行或一列区域。',
    'invalid-list-source':
      '请输入不超过 255 个字符的逗号分隔列表，或有效的单行/单列区域。',
    'invalid-number': '请输入有效的数字；整数验证不能使用小数边界。',
    'invalid-operator': '请选择与当前验证类型匹配的数据条件。',
    'invalid-range': '请选择一个或多个有效的连续单元格区域。',
    'invalid-text-length': '文本长度边界必须是大于或等于 0 的整数。',
    'missing-value': '请填写验证条件所需的值。',
    'multiple-list-columns': '下拉列表来源只能是一行或一列连续单元格。',
    'out-of-bounds': '所选区域超出了工作表的有效边界。',
    'protected-range': '不能修改受保护、只读、合并或透视表区域的数据验证。',
    'range-too-large': `一次最多可设置 ${MAX_SPREADSHEET_DATA_VALIDATION_CELLS.toLocaleString('en-US')} 个单元格的数据验证。`,
    'sheet-not-found': '找不到要设置数据验证的工作表。',
    'value-order': '结束值不能小于开始值。',
  };
  return { ok: false, code, message: messages[code] };
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value && typeof value === 'object' && !Array.isArray(value));
}
