import type { Cell } from '@fortune-sheet/core';
import { cloneSparseMatrix } from '../spreadsheet-sparse';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetSheet,
} from '../work-types';
import {
  coalesceXlsxRichTextRuns,
  MAX_XLSX_RICH_TEXT_FONT_NAME_CHARACTERS,
  MAX_XLSX_RICH_TEXT_FONT_SIZE,
  MAX_XLSX_RICH_TEXT_RUNS_PER_CELL,
  normalizeXlsxRichTextEditSource,
  normalizeXlsxRichTextColor,
  type XlsxRichTextRun,
} from '../work-xlsx-rich-text-model';

export type SpreadsheetRichTextFormatAttribute =
  | 'bl'
  | 'cl'
  | 'fc'
  | 'ff'
  | 'fs'
  | 'it'
  | 'un';

export type SpreadsheetRichTextToggleAttribute = 'bl' | 'cl' | 'it' | 'un';

export interface SpreadsheetTextSelection {
  end: number;
  start: number;
}

export interface SpreadsheetRichTextSelectionFormatRequest {
  attribute: SpreadsheetRichTextFormatAttribute;
  column: number;
  row: number;
  selection: SpreadsheetTextSelection;
  sheetId: string;
  /**
   * A cell reconstructed from the live editor before focus moves into a
   * ribbon popover. It is intentionally internal and still passes through all
   * text, run-count, coordinate, and formatting bounds below.
   */
  sourceCell?: Cell;
  value: unknown;
}

interface NormalizedRichTextSource {
  runs: XlsxRichTextRun[];
  text: string;
}

type SpreadsheetCellDataEntry = NonNullable<
  WorkSpreadsheetSheet['celldata']
>[number];

const formatAttributes = new Set<SpreadsheetRichTextFormatAttribute>([
  'bl',
  'cl',
  'fc',
  'ff',
  'fs',
  'it',
  'un',
]);

const MAX_SPREADSHEET_ROW = 1_048_575;
const MAX_SPREADSHEET_COLUMN = 16_383;

export function canApplySpreadsheetRichTextSelectionFormat(
  content: WorkSpreadsheetContent,
  request: SpreadsheetRichTextSelectionFormatRequest,
): boolean {
  return resolveSpreadsheetRichTextSelectionFormat(content, request) !== null;
}

export function applySpreadsheetRichTextSelectionFormat(
  content: WorkSpreadsheetContent,
  request: SpreadsheetRichTextSelectionFormatRequest,
): WorkSpreadsheetContent | null {
  const resolved = resolveSpreadsheetRichTextSelectionFormat(content, request);
  if (!resolved) return null;
  const sheets = [...content.sheets];
  sheets[resolved.sheetIndex] = replaceSpreadsheetRichTextCell(
    resolved.sheet,
    request.row,
    request.column,
    resolved.cell,
  );
  return { ...content, sheets };
}

export function formatSpreadsheetRichTextSelection(
  cell: Cell,
  selection: SpreadsheetTextSelection,
  attribute: SpreadsheetRichTextFormatAttribute,
  value: unknown,
): Cell | null {
  const normalizedValue = normalizeFormatValue(attribute, value);
  const source = normalizedRichTextSource(cell);
  if (
    !source ||
    normalizedValue === INVALID_FORMAT_VALUE ||
    !validTextSelection(source.text, selection)
  ) {
    return null;
  }

  const splitRuns: XlsxRichTextRun[] = [];
  let offset = 0;
  for (const run of source.runs) {
    const runStart = offset;
    const runEnd = runStart + run.v.length;
    offset = runEnd;
    const overlapStart = Math.max(runStart, selection.start);
    const overlapEnd = Math.min(runEnd, selection.end);
    if (overlapStart >= overlapEnd) {
      splitRuns.push({ ...run });
      continue;
    }
    appendRunSegment(splitRuns, run, runStart, overlapStart);
    appendRunSegment(
      splitRuns,
      patchRichTextRun(run, attribute, normalizedValue),
      overlapStart,
      overlapEnd,
      runStart,
    );
    appendRunSegment(splitRuns, run, overlapEnd, runEnd, runStart);
  }

  const runs = coalesceXlsxRichTextRuns(splitRuns);
  if (!runs.length || runs.length > MAX_XLSX_RICH_TEXT_RUNS_PER_CELL) {
    return null;
  }
  const next: Cell = {
    ...cell,
    ct: { ...cell.ct, s: runs, t: 'inlineStr' },
    v: source.text,
  };
  delete next.f;
  delete next.m;
  return next;
}

export function spreadsheetRichTextSelectionToggleValue(
  cell: Cell,
  selection: SpreadsheetTextSelection,
  attribute: SpreadsheetRichTextToggleAttribute,
): 0 | 1 | null {
  const source = normalizedRichTextSource(cell);
  if (!source || !validTextSelection(source.text, selection)) return null;
  let offset = 0;
  let selectedRuns = 0;
  let allActive = true;
  for (const run of source.runs) {
    const runStart = offset;
    const runEnd = runStart + run.v.length;
    offset = runEnd;
    if (
      Math.max(runStart, selection.start) >= Math.min(runEnd, selection.end)
    ) {
      continue;
    }
    selectedRuns += 1;
    if (
      attribute === 'un'
        ? Number(run.un) === 0 || run.un === undefined
        : Number(run[attribute]) !== 1
    ) {
      allActive = false;
    }
  }
  return selectedRuns ? (allActive ? 0 : 1) : null;
}

const INVALID_FORMAT_VALUE = Symbol('invalid-format-value');

function resolveSpreadsheetRichTextSelectionFormat(
  content: WorkSpreadsheetContent,
  request: SpreadsheetRichTextSelectionFormatRequest,
): {
  cell: Cell;
  sheet: WorkSpreadsheetSheet;
  sheetIndex: number;
} | null {
  if (
    !request.sheetId ||
    !Number.isInteger(request.row) ||
    request.row < 0 ||
    request.row > MAX_SPREADSHEET_ROW ||
    !Number.isInteger(request.column) ||
    request.column < 0 ||
    request.column > MAX_SPREADSHEET_COLUMN ||
    !formatAttributes.has(request.attribute)
  ) {
    return null;
  }
  const sheetIndex = content.sheets.findIndex(
    (candidate) => candidate.id === request.sheetId,
  );
  const sheet = content.sheets[sheetIndex];
  if (!sheet || sheetIndex < 0) return null;
  const sourceCell =
    request.sourceCell ?? sheetCellAt(sheet, request.row, request.column);
  if (!sourceCell) return null;
  const cell = formatSpreadsheetRichTextSelection(
    sourceCell,
    request.selection,
    request.attribute,
    request.value,
  );
  return cell ? { cell, sheet, sheetIndex } : null;
}

function normalizedRichTextSource(cell: Cell): NormalizedRichTextSource | null {
  const source = normalizeXlsxRichTextEditSource(cell);
  return source?.text ? source : null;
}

function normalizeFormatValue(
  attribute: SpreadsheetRichTextFormatAttribute,
  value: unknown,
): unknown | typeof INVALID_FORMAT_VALUE {
  if (attribute === 'bl' || attribute === 'cl' || attribute === 'it') {
    const toggle = Number(value);
    return Number.isSafeInteger(toggle) && (toggle === 0 || toggle === 1)
      ? toggle
      : INVALID_FORMAT_VALUE;
  }
  if (attribute === 'un') {
    const underline = Number(value);
    return Number.isSafeInteger(underline) && underline >= 0 && underline <= 4
      ? underline
      : INVALID_FORMAT_VALUE;
  }
  if (attribute === 'ff') {
    return typeof value === 'string' &&
      value.trim() &&
      value.trim().length <= MAX_XLSX_RICH_TEXT_FONT_NAME_CHARACTERS
      ? value.trim()
      : INVALID_FORMAT_VALUE;
  }
  if (attribute === 'fs') {
    return typeof value === 'number' &&
      Number.isFinite(value) &&
      value >= 1 &&
      value <= MAX_XLSX_RICH_TEXT_FONT_SIZE
      ? value
      : INVALID_FORMAT_VALUE;
  }
  if (value === undefined) return undefined;
  return normalizeXlsxRichTextColor(value) ?? INVALID_FORMAT_VALUE;
}

function patchRichTextRun(
  run: XlsxRichTextRun,
  attribute: SpreadsheetRichTextFormatAttribute,
  value: unknown,
): XlsxRichTextRun {
  const next = { ...run };
  if (attribute === 'fc') {
    delete next.a3sXlsxColorOrigin;
    if (value === undefined) delete next.fc;
    else next.fc = value as string;
  } else if (attribute === 'ff') {
    next.ff = value as string;
  } else if (attribute === 'fs') {
    next.fs = value as number;
  } else if (attribute === 'un') {
    next.un = value as 0 | 1 | 2 | 3 | 4;
  } else {
    next[attribute] = value as 0 | 1;
  }
  return next;
}

function appendRunSegment(
  target: XlsxRichTextRun[],
  source: XlsxRichTextRun,
  absoluteStart: number,
  absoluteEnd: number,
  sourceStart = absoluteStart,
): void {
  if (absoluteStart >= absoluteEnd) return;
  const value = source.v.slice(
    absoluteStart - sourceStart,
    absoluteEnd - sourceStart,
  );
  if (value) target.push({ ...source, v: value });
}

function validTextSelection(
  text: string,
  selection: SpreadsheetTextSelection,
): boolean {
  return Boolean(
    Number.isInteger(selection.start) &&
      Number.isInteger(selection.end) &&
      selection.start >= 0 &&
      selection.start < selection.end &&
      selection.end <= text.length &&
      !splitsSurrogatePair(text, selection.start) &&
      !splitsSurrogatePair(text, selection.end),
  );
}

function splitsSurrogatePair(text: string, offset: number): boolean {
  if (offset <= 0 || offset >= text.length) return false;
  const previous = text.charCodeAt(offset - 1);
  const next = text.charCodeAt(offset);
  return (
    previous >= 0xd800 && previous <= 0xdbff && next >= 0xdc00 && next <= 0xdfff
  );
}

function replaceSpreadsheetRichTextCell(
  sheet: WorkSpreadsheetSheet,
  row: number,
  column: number,
  cell: Cell,
): WorkSpreadsheetSheet {
  if (sheet.data !== undefined) {
    const data = cloneSparseMatrix(sheet.data);
    const values = data[row] ?? [];
    data[row] = values;
    values[column] = cell;
    return {
      ...sheet,
      column: Math.max(sheet.column ?? 0, column + 1),
      data,
      row: Math.max(sheet.row ?? 0, row + 1),
    };
  }
  const entries = [...(sheet.celldata ?? [])];
  const index = entries.findIndex(
    (candidate) => candidate.r === row && candidate.c === column,
  );
  const entry: SpreadsheetCellDataEntry = { r: row, c: column, v: cell };
  if (index >= 0) entries[index] = entry;
  else entries.push(entry);
  entries.sort((left, right) => left.r - right.r || left.c - right.c);
  return {
    ...sheet,
    celldata: entries,
    column: Math.max(sheet.column ?? 0, column + 1),
    row: Math.max(sheet.row ?? 0, row + 1),
  };
}

function sheetCellAt(
  sheet: WorkSpreadsheetSheet,
  row: number,
  column: number,
): Cell | null {
  const matrix = sheet.data?.[row]?.[column];
  if (matrix) return matrix;
  return (
    sheet.celldata?.find(
      (candidate) => candidate.r === row && candidate.c === column,
    )?.v ?? null
  );
}
