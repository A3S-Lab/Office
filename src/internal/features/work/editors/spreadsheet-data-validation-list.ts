import type { Cell } from '@fortune-sheet/core';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetDataValidationItem,
  WorkSpreadsheetSheet,
} from '../work-types';
import { resolveSpreadsheetGoToTarget } from './spreadsheet-go-to';
import {
  spreadsheetCellRangeArea,
  spreadsheetCellRangeContains,
  type SpreadsheetCellRange,
} from './spreadsheet-cell-range';
import { parseSpreadsheetCellRanges } from '../work-spreadsheet-ranges';

/** Keep dependent-list parsing synchronous and bounded at the editor boundary. */
export const MAX_SPREADSHEET_DEPENDENT_LIST_FORMULA_LENGTH = 255;
export const MAX_SPREADSHEET_DEPENDENT_LIST_REFERENCE_CELLS = 1_024;
export const MAX_SPREADSHEET_DEPENDENT_LIST_MATERIALIZED_CELLS = 10_000;

const DEPENDENT_LIST_PROJECTION_KEY = '__a3sDependentListProjection';

interface DependentListProjection {
  kind: 'compact' | 'direct';
  formula: string;
}

type ProjectedValidationItem = WorkSpreadsheetDataValidationItem & {
  [DEPENDENT_LIST_PROJECTION_KEY]?: DependentListProjection;
};

export interface SpreadsheetDependentListFormulaResult {
  ok: boolean;
  formula?: string;
  message?: string;
}

export interface SpreadsheetDependentListReferenceResult {
  ok: boolean;
  empty?: boolean;
  displayReference?: string;
  message?: string;
}

/**
 * Returns true for the deliberately small, local dependent-list grammar.
 *
 * The leading `=` is optional when reading an existing workbook, but authoring
 * always writes it back so a formula cannot be confused with a literal list.
 */
export function isSpreadsheetDependentListFormula(value: unknown): boolean {
  return (
    typeof value === 'string' && /^=?\s*INDIRECT\s*\(/iu.test(value.trim())
  );
}

/**
 * Normalize an INDIRECT list source without evaluating workbook state.
 * Supported arguments are quoted text, a single-cell reference, or a bounded
 * `&` concatenation of those terms. Nested functions and external books are
 * intentionally rejected at the local editor boundary.
 */
export function normalizeSpreadsheetDependentListFormula(
  value: string,
): SpreadsheetDependentListFormulaResult {
  const source = value.trim();
  if (!source) return { ok: false, message: '动态下拉来源不能为空。' };
  if (
    Array.from(source).length > MAX_SPREADSHEET_DEPENDENT_LIST_FORMULA_LENGTH
  ) {
    return { ok: false, message: '动态下拉公式超过 255 个字符。' };
  }
  if (/[\u0000-\u001f\u007f]/u.test(source) || /[\[\]]/u.test(source)) {
    return { ok: false, message: '动态下拉公式只能引用当前工作簿。' };
  }
  const body = source.replace(/^=/, '').trim();
  const argument = indirectArgument(body);
  if (argument === null) {
    return {
      ok: false,
      message: '动态下拉公式仅支持 =INDIRECT(单元格或文本拼接)。',
    };
  }
  const terms = splitConcatenation(argument);
  if (!terms || terms.length === 0 || terms.length > 16) {
    return {
      ok: false,
      message: '动态下拉公式的拼接项数量超出本地安全边界。',
    };
  }
  if (!terms.every((term) => isDependentListTerm(term))) {
    return {
      ok: false,
      message: '动态下拉公式只能拼接文本和单元格引用。',
    };
  }
  return {
    ok: true,
    formula: `=INDIRECT(${terms.map((term) => term.trim()).join('&')})`,
  };
}

/**
 * Resolve one dependent list for a target cell. An empty driver cell is a
 * valid state and produces an empty list; malformed non-empty references fail
 * closed so the editor never silently offers the wrong options.
 */
export function resolveSpreadsheetDependentListReference(
  content: WorkSpreadsheetContent,
  sheetId: string,
  row: number,
  column: number,
  anchorRow: number,
  anchorColumn: number,
  formula: string,
): SpreadsheetDependentListReferenceResult {
  const normalized = normalizeSpreadsheetDependentListFormula(formula);
  if (!normalized.ok || !normalized.formula) {
    return { ok: false, message: normalized.message };
  }
  const body = normalized.formula.slice(1);
  const argument = indirectArgument(body);
  if (argument === null) {
    return { ok: false, message: '动态下拉公式括号不完整。' };
  }
  const terms = splitConcatenation(argument);
  if (!terms) {
    return { ok: false, message: '动态下拉公式拼接项无效。' };
  }
  let referenceText = '';
  for (const term of terms) {
    const parsed = parseDependentListTerm(term.trim());
    if (!parsed) {
      return { ok: false, message: '动态下拉公式包含不支持的表达式。' };
    }
    if (parsed.kind === 'text') {
      referenceText += parsed.value;
      continue;
    }
    const resolvedCell = resolveDependentListCell(
      content,
      sheetId,
      row,
      column,
      anchorRow,
      anchorColumn,
      parsed,
    );
    if (!resolvedCell.ok) return resolvedCell;
    referenceText += resolvedCell.value;
  }
  const trimmedReference = referenceText.trim();
  if (!trimmedReference) return { ok: true, empty: true };
  const target = resolveSpreadsheetGoToTarget(
    content,
    sheetId,
    trimmedReference,
  );
  if (!target.ok) {
    return {
      ok: false,
      message: `动态下拉来源无法解析：${target.message}`,
    };
  }
  const area = spreadsheetCellRangeArea(target.target.range);
  if (
    target.target.range.row[0] !== target.target.range.row[1] &&
    target.target.range.column[0] !== target.target.range.column[1]
  ) {
    return { ok: false, message: '动态下拉来源只能是一行或一列连续区域。' };
  }
  if (area > MAX_SPREADSHEET_DEPENDENT_LIST_REFERENCE_CELLS) {
    return {
      ok: false,
      message: `动态下拉来源最多读取 ${MAX_SPREADSHEET_DEPENDENT_LIST_REFERENCE_CELLS.toLocaleString('en-US')} 个单元格。`,
    };
  }
  return {
    ok: true,
    displayReference: target.target.displayReference,
  };
}

/**
 * Materialize only the runtime view Fortune needs. The controlled model keeps
 * the INDIRECT formula; each generated item carries a private marker so the
 * projection cleanup can restore it after Fortune emits a cell operation.
 */
export function materializeSpreadsheetDependentListsForFortune(
  sheets: readonly WorkSpreadsheetSheet[],
  namedRanges: WorkSpreadsheetContent['namedRanges'] = [],
): WorkSpreadsheetSheet[] {
  const content: WorkSpreadsheetContent = {
    type: 'spreadsheet',
    sheets: sheets.map((sheet) => sheet),
    namedRanges,
  };
  return sheets.map((sheet) => {
    let dataVerification: Record<string, unknown> | undefined;
    let dataValidationRanges: WorkSpreadsheetSheet['dataValidationRanges'];
    let changed = false;
    let materializationBudget =
      MAX_SPREADSHEET_DEPENDENT_LIST_MATERIALIZED_CELLS;
    const direct = sheet.dataVerification as
      | Record<string, unknown>
      | undefined;
    if (direct) {
      for (const [key, value] of Object.entries(direct)) {
        const coordinate = dependentListCoordinate(key);
        const item = dependentListItem(value);
        if (
          !coordinate ||
          !item ||
          !isSpreadsheetDependentListFormula(item.value1)
        ) {
          continue;
        }
        const range = dependentListAnchorRange(
          item,
          coordinate.row,
          coordinate.column,
        );
        const reference = resolveSpreadsheetDependentListReference(
          content,
          sheet.id ?? '',
          coordinate.row,
          coordinate.column,
          range.row[0],
          range.column[0],
          item.value1,
        );
        const projected = projectedDependentListItem(item, reference, 'direct');
        dataVerification ??= { ...direct };
        dataVerification[key] = projected;
        changed = true;
        materializationBudget = Math.max(0, materializationBudget - 1);
      }
    }
    for (const [entryIndex, entry] of (
      sheet.dataValidationRanges ?? []
    ).entries()) {
      if (!isSpreadsheetDependentListFormula(entry.item.value1)) continue;
      const formula = entry.item.value1;
      const projectedItem = {
        ...entry.item,
        value1: '',
        [DEPENDENT_LIST_PROJECTION_KEY]: {
          kind: 'compact' as const,
          formula,
        },
      } satisfies ProjectedValidationItem;
      dataValidationRanges ??= [...(sheet.dataValidationRanges ?? [])];
      dataValidationRanges[entryIndex] = {
        ...entry,
        item: projectedItem,
      };
      changed = true;
      for (const range of entry.ranges) {
        const area = spreadsheetCellRangeArea(range);
        if (area > materializationBudget) continue;
        materializationBudget -= area;
        for (let row = range.row[0]; row <= range.row[1]; row += 1) {
          for (
            let column = range.column[0];
            column <= range.column[1];
            column += 1
          ) {
            const key = `${row}_${column}`;
            if (direct && Object.hasOwn(direct, key)) continue;
            const reference = resolveSpreadsheetDependentListReference(
              content,
              sheet.id ?? '',
              row,
              column,
              range.row[0],
              range.column[0],
              entry.item.value1,
            );
            dataVerification ??= { ...(direct ?? {}) };
            dataVerification[key] = projectedDependentListItem(
              entry.item,
              reference,
              'compact',
            );
            changed = true;
          }
        }
      }
    }
    return changed
      ? {
          ...sheet,
          dataVerification:
            dataVerification as WorkSpreadsheetSheet['dataVerification'],
          dataValidationRanges,
        }
      : sheet;
  });
}

/** Restore controlled formulas and remove generated compact-cell overrides. */
export function restoreSpreadsheetDependentListProjections(
  sheets: readonly WorkSpreadsheetSheet[],
): WorkSpreadsheetSheet[] {
  return sheets.map((sheet) => {
    const direct = sheet.dataVerification as
      | Record<string, unknown>
      | undefined;
    let nextDirect: Record<string, unknown> | undefined;
    if (direct) {
      for (const [key, value] of Object.entries(direct)) {
        const item = dependentListItem(value) as ProjectedValidationItem | null;
        const marker = item?.[DEPENDENT_LIST_PROJECTION_KEY];
        if (!marker) continue;
        nextDirect ??= { ...direct };
        if (marker.kind === 'compact') {
          delete nextDirect[key];
        } else {
          const { [DEPENDENT_LIST_PROJECTION_KEY]: _marker, ...rest } = item;
          nextDirect[key] = { ...rest, value1: marker.formula };
        }
      }
    }
    let nextRanges: WorkSpreadsheetSheet['dataValidationRanges'];
    for (const [entryIndex, entry] of (
      sheet.dataValidationRanges ?? []
    ).entries()) {
      const item = entry.item as ProjectedValidationItem;
      const marker = item[DEPENDENT_LIST_PROJECTION_KEY];
      if (!marker) continue;
      const { [DEPENDENT_LIST_PROJECTION_KEY]: _marker, ...rest } = item;
      nextRanges ??= [...(sheet.dataValidationRanges ?? [])];
      nextRanges[entryIndex] = {
        ...entry,
        item: { ...rest, value1: marker.formula },
      };
    }
    if (!nextDirect && !nextRanges) return sheet;
    return {
      ...sheet,
      dataVerification:
        nextDirect && Object.keys(nextDirect).length
          ? (nextDirect as WorkSpreadsheetSheet['dataVerification'])
          : undefined,
      dataValidationRanges: nextRanges,
    };
  });
}

function projectedDependentListItem(
  item: ProjectedValidationItem,
  reference: SpreadsheetDependentListReferenceResult,
  kind: DependentListProjection['kind'],
): ProjectedValidationItem {
  const { [DEPENDENT_LIST_PROJECTION_KEY]: _previous, ...rest } = item;
  return {
    ...rest,
    value1: reference.ok ? (reference.displayReference ?? '') : '',
    [DEPENDENT_LIST_PROJECTION_KEY]: { kind, formula: item.value1 },
  };
}

function dependentListItem(value: unknown): ProjectedValidationItem | null {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return null;
  if (typeof (value as { type?: unknown }).type !== 'string') return null;
  return value as ProjectedValidationItem;
}

function dependentListCoordinate(
  value: string,
): { row: number; column: number } | null {
  const match = /^(\d+)_(\d+)$/.exec(value);
  if (!match) return null;
  const row = Number(match[1]);
  const column = Number(match[2]);
  return Number.isSafeInteger(row) && Number.isSafeInteger(column)
    ? { row, column }
    : null;
}

function dependentListAnchorRange(
  item: WorkSpreadsheetDataValidationItem,
  row: number,
  column: number,
): SpreadsheetCellRange {
  const ranges = parseSpreadsheetCellRanges(item.rangeTxt);
  const containing = ranges?.find((range) =>
    spreadsheetCellRangeContains(range, row, column),
  );
  return containing
    ? {
        row: [containing.row[0], containing.row[0]],
        column: [containing.column[0], containing.column[0]],
      }
    : { row: [row, row], column: [column, column] };
}

function indirectArgument(value: string): string | null {
  const match = /^INDIRECT\s*\(/iu.exec(value.trim());
  if (!match) return null;
  const start = match[0].length;
  let depth = 1;
  let quoted = false;
  for (let index = start; index < value.length; index += 1) {
    const character = value[index];
    if (character === '"') {
      if (quoted && value[index + 1] === '"') {
        index += 1;
        continue;
      }
      quoted = !quoted;
      continue;
    }
    if (quoted) continue;
    if (character === '(') depth += 1;
    if (character !== ')') continue;
    depth -= 1;
    if (depth === 0) {
      return value.slice(start, index).trim() && !value.slice(index + 1).trim()
        ? value.slice(start, index).trim()
        : null;
    }
  }
  return null;
}

function splitConcatenation(value: string): string[] | null {
  const terms: string[] = [];
  let start = 0;
  let quoted = false;
  for (let index = 0; index < value.length; index += 1) {
    const character = value[index];
    if (character === '"') {
      if (quoted && value[index + 1] === '"') {
        index += 1;
        continue;
      }
      quoted = !quoted;
    } else if (!quoted && character === '&') {
      terms.push(value.slice(start, index).trim());
      start = index + 1;
    } else if (!quoted && (character === '(' || character === ')')) {
      return null;
    }
  }
  if (quoted) return null;
  terms.push(value.slice(start).trim());
  return terms;
}

function isDependentListTerm(value: string): boolean {
  return Boolean(parseDependentListTerm(value));
}

type DependentListTerm =
  | { kind: 'text'; value: string }
  | {
      kind: 'cell';
      qualifier?: string;
      column: number;
      row: number;
      absoluteColumn: boolean;
      absoluteRow: boolean;
    };

function parseDependentListTerm(value: string): DependentListTerm | null {
  const text = value.trim();
  if (text.startsWith('"') && text.endsWith('"')) {
    const inner = text.slice(1, -1);
    for (let index = 0; index < inner.length; index += 1) {
      if (inner[index] !== '"') continue;
      if (inner[index + 1] === '"') {
        index += 1;
        continue;
      }
      return null;
    }
    return { kind: 'text', value: inner.replaceAll('""', '"') };
  }
  const match =
    /^(?:(?:'((?:[^']|'')+)'|([^'!\[\]\s]+))!)?(\$?)([A-Z]{1,3})(\$?)([1-9]\d*)$/iu.exec(
      text,
    );
  if (!match) return null;
  const column = decodeColumn(match[4]);
  const row = Number(match[6]) - 1;
  if (column < 0 || row < 0 || column >= 16_384 || row >= 1_048_576)
    return null;
  return {
    kind: 'cell',
    qualifier: match[1] ? match[1].replaceAll("''", "'") : match[2],
    column,
    row,
    absoluteColumn: Boolean(match[3]),
    absoluteRow: Boolean(match[5]),
  };
}

function resolveDependentListCell(
  content: WorkSpreadsheetContent,
  sheetId: string,
  row: number,
  column: number,
  anchorRow: number,
  anchorColumn: number,
  term: Extract<DependentListTerm, { kind: 'cell' }>,
): { ok: true; value: string } | { ok: false; message: string } {
  const fallback = content.sheets.find((sheet) => sheet.id === sheetId);
  const sheet = term.qualifier
    ? content.sheets.find(
        (candidate) =>
          candidate.name.trim().toLocaleLowerCase() ===
          term.qualifier?.trim().toLocaleLowerCase(),
      )
    : fallback;
  if (!sheet)
    return { ok: false, message: '动态下拉公式引用了不存在的工作表。' };
  const resolvedRow = term.absoluteRow
    ? term.row
    : row + (term.row - anchorRow);
  const resolvedColumn = term.absoluteColumn
    ? term.column
    : column + (term.column - anchorColumn);
  if (
    resolvedRow < 0 ||
    resolvedColumn < 0 ||
    resolvedRow >= 1_048_576 ||
    resolvedColumn >= 16_384
  ) {
    return { ok: false, message: '动态下拉公式引用超出了工作表边界。' };
  }
  const cell = spreadsheetDependentListCellAt(
    sheet,
    resolvedRow,
    resolvedColumn,
  );
  if (!cell) return { ok: true, value: '' };
  if (typeof cell.f === 'string' && cell.f && cell.v === undefined) {
    return {
      ok: false,
      message: '动态下拉公式引用了尚未计算的公式单元格。',
    };
  }
  const displayed = cell.m ?? cell.v;
  return {
    ok: true,
    value:
      displayed === null || displayed === undefined ? '' : String(displayed),
  };
}

function spreadsheetDependentListCellAt(
  sheet: WorkSpreadsheetSheet,
  row: number,
  column: number,
): Cell | null {
  const direct = sheet.data?.[row]?.[column];
  if (direct) return direct;
  const entry = sheet.celldata?.find(
    (candidate) => candidate.r === row && candidate.c === column,
  );
  if (!entry || !entry.v || typeof entry.v !== 'object') return null;
  return entry.v as Cell;
}

function decodeColumn(value: string): number {
  let result = 0;
  for (const character of value.toUpperCase()) {
    result = result * 26 + character.charCodeAt(0) - 64;
  }
  return result - 1;
}
