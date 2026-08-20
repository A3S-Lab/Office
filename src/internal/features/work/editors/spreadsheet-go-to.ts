import {
  formatSpreadsheetCellRanges,
  parseSpreadsheetCellRanges,
} from '../work-spreadsheet-ranges';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetNamedRange,
  WorkSpreadsheetSheet,
} from '../work-types';
import type { SpreadsheetCellRange } from './spreadsheet-cell-range';
import { spreadsheetSheetBounds } from './spreadsheet-keyboard-navigation';

const spreadsheetMaximumRows = 1_048_576;
const spreadsheetMaximumColumns = 16_384;
const spreadsheetGoToMaximumInputLength = 512;

export type SpreadsheetGoToErrorCode =
  | 'ambiguous-name'
  | 'empty'
  | 'hidden-sheet'
  | 'invalid-reference'
  | 'multiple-ranges'
  | 'out-of-bounds'
  | 'sheet-not-found'
  | 'unsupported-name';

export interface SpreadsheetGoToTarget {
  displayReference: string;
  name?: string;
  range: SpreadsheetCellRange;
  selection: SpreadsheetCellRange & {
    row_focus: number;
    column_focus: number;
  };
  sheetId: string;
  source: 'named-range' | 'reference';
}

export type SpreadsheetGoToResolution =
  | { ok: true; target: SpreadsheetGoToTarget }
  | {
      ok: false;
      code: SpreadsheetGoToErrorCode;
      message: string;
    };

interface QualifiedSpreadsheetReference {
  qualifier: string | null;
  reference: string;
}

interface ResolvedSpreadsheetRange {
  range: SpreadsheetCellRange;
  sheet: WorkSpreadsheetSheet;
}

type SpreadsheetRangeResolution =
  | { kind: 'match'; value: ResolvedSpreadsheetRange }
  | { kind: 'miss' }
  | {
      kind: 'error';
      code: SpreadsheetGoToErrorCode;
    };

export function resolveSpreadsheetGoToTarget(
  content: WorkSpreadsheetContent,
  activeSheetId: string,
  input: string,
): SpreadsheetGoToResolution {
  const value = input.trim().replace(/^=/, '').trim();
  if (!value) return spreadsheetGoToError('empty');
  if (value.length > spreadsheetGoToMaximumInputLength) {
    return spreadsheetGoToError('invalid-reference');
  }
  const activeSheet = content.sheets.find(
    (sheet) => sheet.id === activeSheetId,
  );
  if (!activeSheet) return spreadsheetGoToError('sheet-not-found');

  const qualified = splitSpreadsheetQualifiedReference(value);
  if (!qualified) return spreadsheetGoToError('invalid-reference');
  const requestedSheet = qualified.qualifier
    ? spreadsheetSheetByName(content, qualified.qualifier)
    : activeSheet;
  if (qualified.qualifier && !requestedSheet) {
    return spreadsheetGoToError('sheet-not-found');
  }
  if (requestedSheet?.hide === 1) {
    return spreadsheetGoToError('hidden-sheet');
  }

  const direct = resolveSpreadsheetDirectRange(
    content,
    requestedSheet ?? activeSheet,
    qualified,
  );
  if (direct.kind === 'error') return spreadsheetGoToError(direct.code);
  if (direct.kind === 'match') {
    return spreadsheetGoToSuccess(direct.value, 'reference');
  }

  const named = findSpreadsheetGoToName(
    content,
    qualified.reference,
    requestedSheet ?? activeSheet,
    Boolean(qualified.qualifier),
  );
  if (named.kind === 'error') return spreadsheetGoToError(named.code);
  if (!named.value) return spreadsheetGoToError('invalid-reference');
  const fallbackSheet = named.value.scopeSheetId
    ? content.sheets.find((sheet) => sheet.id === named.value?.scopeSheetId)
    : (requestedSheet ?? activeSheet);
  if (!fallbackSheet) return spreadsheetGoToError('sheet-not-found');
  if (fallbackSheet.hide === 1) return spreadsheetGoToError('hidden-sheet');
  const namedReference = splitSpreadsheetQualifiedReference(
    named.value.reference.trim().replace(/^=/, '').trim(),
  );
  if (!namedReference) return spreadsheetGoToError('unsupported-name');
  const resolved = resolveSpreadsheetDirectRange(
    content,
    fallbackSheet,
    namedReference,
  );
  if (resolved.kind === 'error') return spreadsheetGoToError(resolved.code);
  if (resolved.kind !== 'match') {
    return spreadsheetGoToError('unsupported-name');
  }
  return spreadsheetGoToSuccess(
    resolved.value,
    'named-range',
    named.value.name,
  );
}

export function spreadsheetGoToValidationMessage(
  content: WorkSpreadsheetContent,
  activeSheetId: string,
  input: string,
): string | null {
  const resolution = resolveSpreadsheetGoToTarget(
    content,
    activeSheetId,
    input,
  );
  return resolution.ok ? null : resolution.message;
}

function resolveSpreadsheetDirectRange(
  content: WorkSpreadsheetContent,
  fallbackSheet: WorkSpreadsheetSheet,
  qualified: QualifiedSpreadsheetReference,
): SpreadsheetRangeResolution {
  const sheet = qualified.qualifier
    ? spreadsheetSheetByName(content, qualified.qualifier)
    : fallbackSheet;
  if (!sheet) return { kind: 'error', code: 'sheet-not-found' };
  if (sheet.hide === 1) return { kind: 'error', code: 'hidden-sheet' };
  const ranges = parseSpreadsheetCellRanges(qualified.reference);
  if (!ranges) return { kind: 'miss' };
  if (ranges.length !== 1) {
    return { kind: 'error', code: 'multiple-ranges' };
  }
  const range = ranges[0];
  if (!range || !spreadsheetGoToRangeIsBounded(sheet, range)) {
    return { kind: 'error', code: 'out-of-bounds' };
  }
  return { kind: 'match', value: { range, sheet } };
}

function findSpreadsheetGoToName(
  content: WorkSpreadsheetContent,
  input: string,
  scopeSheet: WorkSpreadsheetSheet,
  qualified: boolean,
):
  | { kind: 'match'; value: WorkSpreadsheetNamedRange | null }
  | { kind: 'error'; code: 'ambiguous-name' } {
  const normalized = input.trim().toLocaleLowerCase();
  const matching = (content.namedRanges ?? []).filter(
    (range) => range.name.trim().toLocaleLowerCase() === normalized,
  );
  const scoped = matching.filter(
    (range) => range.scopeSheetId === scopeSheet.id,
  );
  const workbook = matching.filter((range) => !range.scopeSheetId);
  const preferred = scoped.length ? scoped : workbook;
  if (preferred.length > 1) {
    return { kind: 'error', code: 'ambiguous-name' };
  }
  if (!preferred.length && !qualified) {
    return { kind: 'match', value: null };
  }
  return { kind: 'match', value: preferred[0] ?? null };
}

function spreadsheetGoToRangeIsBounded(
  sheet: WorkSpreadsheetSheet,
  range: SpreadsheetCellRange,
): boolean {
  const bounds = spreadsheetSheetBounds(sheet);
  return (
    range.row[0] >= 0 &&
    range.row[1] < spreadsheetMaximumRows &&
    range.row[1] <= bounds.lastRow &&
    range.column[0] >= 0 &&
    range.column[1] < spreadsheetMaximumColumns &&
    range.column[1] <= bounds.lastColumn
  );
}

function spreadsheetGoToSuccess(
  resolved: ResolvedSpreadsheetRange,
  source: SpreadsheetGoToTarget['source'],
  name?: string,
): SpreadsheetGoToResolution {
  const reference = formatSpreadsheetCellRanges([resolved.range]);
  return {
    ok: true,
    target: {
      displayReference: `${quoteSpreadsheetSheetName(resolved.sheet.name)}!${reference}`,
      name,
      range: resolved.range,
      selection: {
        row: [...resolved.range.row],
        column: [...resolved.range.column],
        row_focus: resolved.range.row[0],
        column_focus: resolved.range.column[0],
      },
      sheetId: resolved.sheet.id ?? '',
      source,
    },
  };
}

function spreadsheetGoToError(
  code: SpreadsheetGoToErrorCode,
): SpreadsheetGoToResolution {
  const messages: Record<SpreadsheetGoToErrorCode, string> = {
    'ambiguous-name': '同一作用域中存在多个同名区域，无法确定定位目标。',
    empty: '请输入单元格、连续区域或已定义名称。',
    'hidden-sheet': '不能定位到隐藏工作表。',
    'invalid-reference': '请输入有效的 A1 单元格、连续区域或已定义名称。',
    'multiple-ranges': '一次只能定位到一个连续区域。',
    'out-of-bounds': '引用超出了目标工作表的有效边界。',
    'sheet-not-found': '找不到引用的工作表。',
    'unsupported-name': '该名称不是可定位的连续单元格区域。',
  };
  return { ok: false, code, message: messages[code] };
}

function splitSpreadsheetQualifiedReference(
  value: string,
): QualifiedSpreadsheetReference | null {
  let quoted = false;
  let separator = -1;
  for (let index = 0; index < value.length; index += 1) {
    const character = value[index];
    if (character === "'") {
      if (quoted && value[index + 1] === "'") {
        index += 1;
        continue;
      }
      quoted = !quoted;
      continue;
    }
    if (character !== '!' || quoted) continue;
    if (separator >= 0) return null;
    separator = index;
  }
  if (quoted) return null;
  if (separator < 0) return { qualifier: null, reference: value.trim() };
  const rawQualifier = value.slice(0, separator).trim();
  const reference = value.slice(separator + 1).trim();
  if (!rawQualifier || !reference) return null;
  const qualifier = unquoteSpreadsheetSheetName(rawQualifier);
  return qualifier ? { qualifier, reference } : null;
}

function unquoteSpreadsheetSheetName(value: string): string | null {
  if (!value.startsWith("'")) return value.includes("'") ? null : value;
  if (!value.endsWith("'") || value.length < 2) return null;
  const inner = value.slice(1, -1);
  for (let index = 0; index < inner.length; index += 1) {
    if (inner[index] !== "'") continue;
    if (inner[index + 1] !== "'") return null;
    index += 1;
  }
  return inner.replaceAll("''", "'");
}

function spreadsheetSheetByName(
  content: WorkSpreadsheetContent,
  name: string,
): WorkSpreadsheetSheet | undefined {
  const normalized = name.trim().toLocaleLowerCase();
  return content.sheets.find(
    (sheet) => sheet.name.trim().toLocaleLowerCase() === normalized,
  );
}

function quoteSpreadsheetSheetName(name: string): string {
  return `'${name.replaceAll("'", "''")}'`;
}
