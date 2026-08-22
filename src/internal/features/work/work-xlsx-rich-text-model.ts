import type { Cell } from '@fortune-sheet/core';
import {
  normalizeXlsxSemanticColorOrigin,
  type XlsxSemanticColorOrigin,
} from './work-xlsx-cell-style-origin';
import { xlsxRgbColor } from './work-xlsx-cell-style-xml';
import type { SpreadsheetUnderlineCellValue } from './work-spreadsheet-underline';

export const MAX_XLSX_RICH_TEXT_CELL_CHARACTERS = 32_767;
export const MAX_XLSX_RICH_TEXT_RUNS_PER_CELL = 512;
export const MAX_XLSX_RICH_TEXT_CELLS = 10_000;
export const MAX_XLSX_RICH_TEXT_RUNS = 100_000;
export const MAX_XLSX_RICH_TEXT_FONT_NAME_CHARACTERS = 128;
export const MAX_XLSX_RICH_TEXT_FONT_SIZE = 409;

export interface XlsxRichTextRun {
  a3sXlsxColorOrigin?: XlsxSemanticColorOrigin;
  bl?: 0 | 1;
  cl?: 0 | 1;
  fc?: string;
  ff?: string;
  fs?: number;
  it?: 0 | 1;
  un?: SpreadsheetUnderlineCellValue;
  v: string;
}

export interface NormalizedXlsxRichTextCell {
  runs: XlsxRichTextRun[];
  text: string;
}

export function normalizeXlsxRichTextCell(
  cell: Cell,
): NormalizedXlsxRichTextCell | null {
  if (cell.f || cell.ct?.t !== 'inlineStr' || !Array.isArray(cell.ct.s)) {
    return null;
  }
  if (
    !cell.ct.s.length ||
    cell.ct.s.length > MAX_XLSX_RICH_TEXT_RUNS_PER_CELL
  ) {
    return null;
  }
  const runs: XlsxRichTextRun[] = [];
  let characterCount = 0;
  for (const candidate of cell.ct.s) {
    const run = normalizeXlsxRichTextRun(candidate);
    if (!run) return null;
    if (!run.v) continue;
    characterCount += run.v.length;
    if (characterCount > MAX_XLSX_RICH_TEXT_CELL_CHARACTERS) return null;
    runs.push(run);
  }
  const text = runs.map((run) => run.v).join('');
  return runs.length && text ? { runs, text } : null;
}

export function normalizeXlsxRichTextRun(
  value: unknown,
): XlsxRichTextRun | null {
  if (
    !isRecord(value) ||
    typeof value.v !== 'string' ||
    !validXlsxRichText(value.v)
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
    value.ff.trim().length <= MAX_XLSX_RICH_TEXT_FONT_NAME_CHARACTERS
  ) {
    run.ff = value.ff.trim();
  }
  if (
    typeof value.fs === 'number' &&
    Number.isFinite(value.fs) &&
    value.fs >= 1 &&
    value.fs <= MAX_XLSX_RICH_TEXT_FONT_SIZE
  ) {
    run.fs = value.fs;
  }
  const color = normalizeXlsxRichTextColor(value.fc);
  if (color) run.fc = color;
  const underline = Number(value.un);
  if (
    Number.isSafeInteger(underline) &&
    underline >= 0 &&
    underline <= 4 &&
    value.un !== undefined
  ) {
    run.un = underline as SpreadsheetUnderlineCellValue;
  }
  const colorOrigin = normalizeXlsxSemanticColorOrigin(
    value.a3sXlsxColorOrigin,
  );
  if (colorOrigin) run.a3sXlsxColorOrigin = colorOrigin;
  return run;
}

export function normalizeXlsxRichTextColor(value: unknown): string | null {
  const rgb = xlsxRgbColor(value);
  return rgb ? `#${rgb.slice(-6).toLowerCase()}` : null;
}

export function coalesceXlsxRichTextRuns(
  source: readonly XlsxRichTextRun[],
): XlsxRichTextRun[] {
  const result: XlsxRichTextRun[] = [];
  for (const run of source) {
    if (!run.v) continue;
    const previous = result.at(-1);
    if (previous && sameXlsxRichTextRunStyle(previous, run)) {
      previous.v += run.v;
    } else {
      result.push({ ...run });
    }
  }
  return result;
}

export function sameXlsxRichTextRunStyle(
  left: XlsxRichTextRun,
  right: XlsxRichTextRun,
): boolean {
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

export function validXlsxRichText(value: string): boolean {
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
      isHighSurrogate(code) &&
      index + 1 < value.length &&
      isLowSurrogate(value.charCodeAt(index + 1))
    ) {
      index += 1;
      continue;
    }
    return false;
  }
  return true;
}

export function isHighSurrogate(value: number): boolean {
  return value >= 0xd800 && value <= 0xdbff;
}

export function isLowSurrogate(value: number): boolean {
  return value >= 0xdc00 && value <= 0xdfff;
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

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
