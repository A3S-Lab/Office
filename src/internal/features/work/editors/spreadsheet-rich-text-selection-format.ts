import type { Cell } from '@fortune-sheet/core';
import { cloneSparseMatrix } from '../spreadsheet-sparse';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetSheet,
} from '../work-types';
import {
  normalizeXlsxSemanticColorOrigin,
  xlsxCellStyleOrigin,
  xlsxSemanticColorMatchesValue,
} from '../work-xlsx-cell-style-origin';
import { xlsxRgbColor } from '../work-xlsx-cell-style-xml';
import {
  MAX_XLSX_RICH_TEXT_CELL_CHARACTERS,
  MAX_XLSX_RICH_TEXT_RUNS_PER_CELL,
  type XlsxRichTextRun,
} from '../work-xlsx-rich-text';

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
const MAX_SPREADSHEET_FONT_NAME_CHARACTERS = 128;
const MAX_SPREADSHEET_FONT_SIZE = 409;

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

  const runs = coalesceRichTextRuns(splitRuns);
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
  if (cell.f) return null;
  if (cell.ct?.t === 'inlineStr' && Array.isArray(cell.ct.s)) {
    if (
      !cell.ct.s.length ||
      cell.ct.s.length > MAX_XLSX_RICH_TEXT_RUNS_PER_CELL
    ) {
      return null;
    }
    const runs: XlsxRichTextRun[] = [];
    let characterCount = 0;
    for (const candidate of cell.ct.s) {
      const run = normalizeRichTextRun(candidate);
      if (!run) return null;
      if (!run.v) continue;
      characterCount += run.v.length;
      if (characterCount > MAX_XLSX_RICH_TEXT_CELL_CHARACTERS) return null;
      runs.push(run);
    }
    const text = runs.map((run) => run.v).join('');
    return runs.length && text ? { runs, text } : null;
  }
  if (
    typeof cell.v !== 'string' ||
    !cell.v ||
    cell.v.length > MAX_XLSX_RICH_TEXT_CELL_CHARACTERS ||
    !validXmlText(cell.v)
  ) {
    return null;
  }
  return { runs: [plainTextBaseRun(cell, cell.v)], text: cell.v };
}

function plainTextBaseRun(cell: Cell, text: string): XlsxRichTextRun {
  const run: XlsxRichTextRun = { v: text };
  if (Number(cell.bl) === 1) run.bl = 1;
  if (Number(cell.it) === 1) run.it = 1;
  if (Number(cell.cl) === 1) run.cl = 1;
  const underline = Number(cell.un);
  if (Number.isSafeInteger(underline) && underline >= 1 && underline <= 4) {
    run.un = underline as 1 | 2 | 3 | 4;
  }
  if (
    typeof cell.ff === 'string' &&
    cell.ff.trim() &&
    cell.ff.trim().length <= MAX_SPREADSHEET_FONT_NAME_CHARACTERS
  ) {
    run.ff = cell.ff.trim();
  }
  const size = Number(cell.fs);
  if (Number.isFinite(size) && size >= 1 && size <= MAX_SPREADSHEET_FONT_SIZE) {
    run.fs = size;
  }
  const color = normalizeColor(cell.fc);
  if (color) {
    run.fc = color;
    const origin = normalizeXlsxSemanticColorOrigin(
      xlsxCellStyleOrigin(cell)?.fontColor,
    );
    if (origin && xlsxSemanticColorMatchesValue(origin, color)) {
      run.a3sXlsxColorOrigin = origin;
    }
  }
  return run;
}

function normalizeRichTextRun(value: unknown): XlsxRichTextRun | null {
  if (
    !isRecord(value) ||
    typeof value.v !== 'string' ||
    !validXmlText(value.v)
  ) {
    return null;
  }
  const run: XlsxRichTextRun = { v: value.v };
  copyToggle(value, run, 'bl');
  copyToggle(value, run, 'it');
  copyToggle(value, run, 'cl');
  if (
    typeof value.ff === 'string' &&
    value.ff.trim() &&
    value.ff.trim().length <= MAX_SPREADSHEET_FONT_NAME_CHARACTERS
  ) {
    run.ff = value.ff.trim();
  }
  if (
    typeof value.fs === 'number' &&
    Number.isFinite(value.fs) &&
    value.fs >= 1 &&
    value.fs <= MAX_SPREADSHEET_FONT_SIZE
  ) {
    run.fs = value.fs;
  }
  const color = normalizeColor(value.fc);
  if (color) run.fc = color;
  const underline = Number(value.un);
  if (
    Number.isSafeInteger(underline) &&
    underline >= 0 &&
    underline <= 4 &&
    value.un !== undefined
  ) {
    run.un = underline as 0 | 1 | 2 | 3 | 4;
  }
  const origin = normalizeXlsxSemanticColorOrigin(value.a3sXlsxColorOrigin);
  if (origin && color && xlsxSemanticColorMatchesValue(origin, color)) {
    run.a3sXlsxColorOrigin = origin;
  }
  return run;
}

function copyToggle(
  source: Record<string, unknown>,
  target: XlsxRichTextRun,
  key: 'bl' | 'cl' | 'it',
): void {
  if (Number(source[key]) === 1) target[key] = 1;
  else if (Number(source[key]) === 0 && source[key] !== undefined) {
    target[key] = 0;
  }
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
      value.trim().length <= MAX_SPREADSHEET_FONT_NAME_CHARACTERS
      ? value.trim()
      : INVALID_FORMAT_VALUE;
  }
  if (attribute === 'fs') {
    return typeof value === 'number' &&
      Number.isFinite(value) &&
      value >= 1 &&
      value <= MAX_SPREADSHEET_FONT_SIZE
      ? value
      : INVALID_FORMAT_VALUE;
  }
  if (value === undefined) return undefined;
  return normalizeColor(value) ?? INVALID_FORMAT_VALUE;
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

function coalesceRichTextRuns(
  source: readonly XlsxRichTextRun[],
): XlsxRichTextRun[] {
  const result: XlsxRichTextRun[] = [];
  for (const run of source) {
    if (!run.v) continue;
    const previous = result.at(-1);
    if (previous && sameRunStyle(previous, run)) previous.v += run.v;
    else result.push({ ...run });
  }
  return result;
}

function sameRunStyle(left: XlsxRichTextRun, right: XlsxRichTextRun): boolean {
  return (
    left.bl === right.bl &&
    left.cl === right.cl &&
    left.fc === right.fc &&
    left.ff === right.ff &&
    left.fs === right.fs &&
    left.it === right.it &&
    left.un === right.un &&
    JSON.stringify(left.a3sXlsxColorOrigin) ===
      JSON.stringify(right.a3sXlsxColorOrigin)
  );
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

function normalizeColor(value: unknown): string | null {
  const rgb = xlsxRgbColor(value);
  return rgb ? `#${rgb.slice(-6).toLowerCase()}` : null;
}

function validXmlText(value: string): boolean {
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    if (
      code === 0x09 ||
      code === 0x0a ||
      code === 0x0d ||
      (code >= 0x20 && code <= 0xd7ff) ||
      (code >= 0xe000 && code <= 0xfffd)
    ) {
      continue;
    }
    if (
      code >= 0xd800 &&
      code <= 0xdbff &&
      index + 1 < value.length &&
      value.charCodeAt(index + 1) >= 0xdc00 &&
      value.charCodeAt(index + 1) <= 0xdfff
    ) {
      index += 1;
      continue;
    }
    return false;
  }
  return true;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
