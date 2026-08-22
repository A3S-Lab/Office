import type { Cell, Selection } from '@fortune-sheet/core';
import { normalizeCssColor } from '../work-css-color';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetSheet,
} from '../work-types';
import type { SpreadsheetRichTextPasteIntent } from '../work-xlsx-rich-text-edit';
import {
  coalesceXlsxRichTextRuns,
  MAX_XLSX_RICH_TEXT_CELL_CHARACTERS,
  MAX_XLSX_RICH_TEXT_FONT_NAME_CHARACTERS,
  MAX_XLSX_RICH_TEXT_FONT_SIZE,
  MAX_XLSX_RICH_TEXT_RUNS_PER_CELL,
  normalizeXlsxRichTextCell,
  normalizeXlsxRichTextEditSource,
  sameXlsxRichTextRunStyle,
  validXlsxRichText,
  type XlsxRichTextRun,
} from '../work-xlsx-rich-text-model';
import { captureSpreadsheetRichTextDomSelection } from './spreadsheet-rich-text-dom-selection';

const MAX_SPREADSHEET_RICH_TEXT_CLIPBOARD_HTML_CHARACTERS = 256_000;

interface PendingSpreadsheetRichTextPaste {
  intent: SpreadsheetRichTextPasteIntent;
  sourceText: string;
}

interface ParsedSpreadsheetRichTextClipboard {
  runs: XlsxRichTextRun[];
  text: string;
}

type RichTextRunStyle = Omit<XlsxRichTextRun, 'v' | 'a3sXlsxColorOrigin'>;

const pendingPastes = new WeakMap<
  WorkSpreadsheetSheet,
  Map<string, PendingSpreadsheetRichTextPaste>
>();

const ignoredClipboardElements = new Set([
  'HEAD',
  'META',
  'NOSCRIPT',
  'SCRIPT',
  'STYLE',
  'TITLE',
]);

const blockClipboardElements = new Set([
  'ADDRESS',
  'BLOCKQUOTE',
  'DIV',
  'H1',
  'H2',
  'H3',
  'H4',
  'H5',
  'H6',
  'LI',
  'P',
  'PRE',
]);

export function parseSpreadsheetRichTextClipboard(
  html: string,
  plainText: string,
): ParsedSpreadsheetRichTextClipboard | null {
  const text = normalizeClipboardText(plainText);
  if (
    !html ||
    html.length > MAX_SPREADSHEET_RICH_TEXT_CLIPBOARD_HTML_CHARACTERS ||
    !text ||
    text.length > MAX_XLSX_RICH_TEXT_CELL_CHARACTERS ||
    !validXlsxRichText(text)
  ) {
    return null;
  }
  const template = document.createElement('template');
  template.innerHTML = html;
  const runs: XlsxRichTextRun[] = [];
  const state = { characterCount: 0, invalid: false };
  for (const child of template.content.childNodes) {
    collectClipboardRuns(child, {}, runs, state);
    if (state.invalid) return null;
  }
  trimTrailingClipboardNewline(runs, text);
  const normalized = coalesceXlsxRichTextRuns(runs);
  if (
    !normalized.length ||
    normalized.length > MAX_XLSX_RICH_TEXT_RUNS_PER_CELL ||
    normalized.map((run) => run.v).join('') !== text ||
    !normalized.some(spreadsheetRichTextRunHasFormatting)
  ) {
    return null;
  }
  return { runs: normalized, text };
}

export function captureSpreadsheetRichTextPaste({
  clipboardData,
  content,
  root,
  selection,
  sheetId,
  target,
}: {
  clipboardData: Pick<DataTransfer, 'getData'>;
  content: WorkSpreadsheetContent;
  root: HTMLElement | null;
  selection: Selection | null | undefined;
  sheetId: string;
  target: EventTarget | null;
}): boolean {
  const domSelection = captureSpreadsheetRichTextDomSelection({ root, target });
  const sheet = content.sheets.find((candidate) => candidate.id === sheetId);
  if (!domSelection || !sheet || !selection) return false;
  const row = focusedAxisIndex(selection.row_focus, selection.row);
  const column = focusedAxisIndex(selection.column_focus, selection.column);
  const cell = spreadsheetCellAt(sheet, row, column);
  const source = normalizeXlsxRichTextEditSource(cell);
  const key = spreadsheetRichTextPasteKey(row, column);
  if (
    !source ||
    domSelection.editor.textContent !== source.text ||
    domSelection.end > source.text.length
  ) {
    clearPendingPaste(sheet, key);
    return false;
  }
  const parsed = parseSpreadsheetRichTextClipboard(
    clipboardData.getData('text/html'),
    clipboardData.getData('text/plain'),
  );
  if (!parsed) {
    clearPendingPaste(sheet, key);
    return false;
  }
  stageSpreadsheetRichTextPaste(sheet, row, column, source.text, {
    end: domSelection.end,
    runs: parsed.runs,
    start: domSelection.start,
    text: parsed.text,
  });
  return true;
}

export function stageSpreadsheetRichTextPaste(
  sheet: WorkSpreadsheetSheet,
  row: number,
  column: number,
  sourceText: string,
  intent: SpreadsheetRichTextPasteIntent,
): void {
  if (!validCoordinate(row) || !validCoordinate(column)) return;
  const key = spreadsheetRichTextPasteKey(row, column);
  const entries = pendingPastes.get(sheet) ?? new Map();
  pendingPastes.set(sheet, entries);
  entries.set(key, {
    intent: {
      ...intent,
      runs: intent.runs.map((run) => ({ ...run })),
    },
    sourceText,
  });
}

export function takeSpreadsheetRichTextPaste(
  sheet: WorkSpreadsheetSheet,
  row: number,
  column: number,
  current: Cell,
): SpreadsheetRichTextPasteIntent | undefined {
  const intent = peekSpreadsheetRichTextPaste(sheet, row, column, current);
  if (intent) consumeSpreadsheetRichTextPaste(sheet, row, column);
  return intent;
}

export function peekSpreadsheetRichTextPaste(
  sheet: WorkSpreadsheetSheet,
  row: number,
  column: number,
  current: Cell,
): SpreadsheetRichTextPasteIntent | undefined {
  const entries = pendingPastes.get(sheet);
  const key = spreadsheetRichTextPasteKey(row, column);
  const pending = entries?.get(key);
  if (!entries || !pending) return undefined;
  const currentText = spreadsheetRichTextCellText(current);
  if (currentText === pending.sourceText) return undefined;
  return currentText === null ? undefined : pending.intent;
}

export function consumeSpreadsheetRichTextPaste(
  sheet: WorkSpreadsheetSheet,
  row: number,
  column: number,
): void {
  clearPendingPaste(sheet, spreadsheetRichTextPasteKey(row, column));
}

function collectClipboardRuns(
  node: Node,
  inherited: RichTextRunStyle,
  runs: XlsxRichTextRun[],
  state: { characterCount: number; invalid: boolean },
): void {
  if (state.invalid) return;
  if (node.nodeType === Node.TEXT_NODE) {
    appendClipboardText(runs, inherited, node.textContent ?? '', state);
    return;
  }
  if (
    !(node instanceof HTMLElement) ||
    ignoredClipboardElements.has(node.tagName)
  ) {
    return;
  }
  const style = spreadsheetClipboardElementStyle(node, inherited);
  if (node.tagName === 'BR') {
    appendClipboardText(runs, style, '\n', state);
    return;
  }
  for (const child of node.childNodes) {
    collectClipboardRuns(child, style, runs, state);
    if (state.invalid) return;
  }
  if (
    blockClipboardElements.has(node.tagName) &&
    hasFollowingClipboardContent(node)
  ) {
    appendClipboardText(runs, style, '\n', state);
  }
}

function appendClipboardText(
  runs: XlsxRichTextRun[],
  style: RichTextRunStyle,
  source: string,
  state: { characterCount: number; invalid: boolean },
): void {
  const value = normalizeClipboardText(source);
  if (!value) return;
  state.characterCount += value.length;
  if (
    state.characterCount > MAX_XLSX_RICH_TEXT_CELL_CHARACTERS ||
    !validXlsxRichText(value)
  ) {
    state.invalid = true;
    return;
  }
  const run: XlsxRichTextRun = { ...style, v: value };
  const previous = runs.at(-1);
  if (previous && sameXlsxRichTextRunStyle(previous, run)) previous.v += value;
  else runs.push(run);
  if (runs.length > MAX_XLSX_RICH_TEXT_RUNS_PER_CELL) state.invalid = true;
}

function spreadsheetClipboardElementStyle(
  element: HTMLElement,
  inherited: RichTextRunStyle,
): RichTextRunStyle {
  const style = { ...inherited };
  if (element.matches('b, strong')) style.bl = 1;
  if (element.matches('i, em')) style.it = 1;
  if (element.matches('s, strike, del')) style.cl = 1;
  if (element.matches('u')) style.un = 1;
  const declarations = cssDeclarations(element.getAttribute('style'));
  applyFontWeight(style, declarations.get('font-weight'));
  applyToggle(style, 'it', declarations.get('font-style'), 'italic');
  applyTextDecoration(style, declarations);
  applyFontFamily(style, declarations.get('font-family'));
  applyFontSize(style, declarations.get('font-size'));
  applyFontColor(style, declarations.get('color'));
  applyFortuneToggle(style, 'cl', declarations.get('lucky-strike'));
  applyFortuneUnderline(style, declarations.get('lucky-underline'));
  if (declarations.get('border-bottom')?.trim().toLowerCase() !== 'none') {
    if (declarations.has('border-bottom')) style.un = 1;
  }
  if (element.tagName === 'FONT') {
    applyFontFamily(style, element.getAttribute('face') ?? undefined);
    applyFontColor(style, element.getAttribute('color') ?? undefined);
    applyHtmlFontSize(style, element.getAttribute('size'));
  }
  return style;
}

function cssDeclarations(value: string | null): Map<string, string> {
  const declarations = new Map<string, string>();
  for (const candidate of value?.split(';') ?? []) {
    const separator = candidate.indexOf(':');
    if (separator <= 0) continue;
    const key = candidate.slice(0, separator).trim().toLowerCase();
    const declaration = candidate.slice(separator + 1).trim();
    if (key && declaration) declarations.set(key, declaration);
  }
  return declarations;
}

function applyFontWeight(
  style: RichTextRunStyle,
  value: string | undefined,
): void {
  const normalized = value?.trim().toLowerCase();
  if (!normalized) return;
  if (normalized === 'bold' || normalized === 'bolder') style.bl = 1;
  else if (normalized === 'normal' || normalized === 'lighter') style.bl = 0;
  else if (/^\d+$/.test(normalized))
    style.bl = Number(normalized) >= 600 ? 1 : 0;
}

function applyToggle(
  style: RichTextRunStyle,
  key: 'it',
  value: string | undefined,
  enabled: string,
): void {
  const normalized = value?.trim().toLowerCase();
  if (!normalized) return;
  style[key] = normalized === enabled || normalized === 'oblique' ? 1 : 0;
}

function applyTextDecoration(
  style: RichTextRunStyle,
  declarations: ReadonlyMap<string, string>,
): void {
  const value =
    declarations.get('text-decoration-line') ??
    declarations.get('text-decoration');
  const normalized = value?.toLowerCase();
  if (!normalized) return;
  if (normalized.includes('none')) {
    style.cl = 0;
    style.un = 0;
    return;
  }
  if (normalized.includes('line-through')) style.cl = 1;
  if (normalized.includes('underline')) {
    style.un = declarations.get('text-decoration-style') === 'double' ? 2 : 1;
  }
}

function applyFontFamily(
  style: RichTextRunStyle,
  value: string | undefined,
): void {
  const family = value
    ?.split(',')[0]
    ?.trim()
    .replace(/^(['"])(.*)\1$/, '$2');
  if (family && family.length <= MAX_XLSX_RICH_TEXT_FONT_NAME_CHARACTERS) {
    style.ff = family;
  }
}

function applyFontSize(
  style: RichTextRunStyle,
  value: string | undefined,
): void {
  const match = /^([\d.]+)\s*(pt|px)$/i.exec(value?.trim() ?? '');
  if (!match) return;
  const number = Number(match[1]);
  const size = match[2]?.toLowerCase() === 'px' ? number * 0.75 : number;
  if (
    Number.isFinite(size) &&
    size >= 1 &&
    size <= MAX_XLSX_RICH_TEXT_FONT_SIZE
  ) {
    style.fs = Math.round(size * 100) / 100;
  }
}

function applyFontColor(
  style: RichTextRunStyle,
  value: string | undefined,
): void {
  const color = normalizeCssColor(value);
  if (color && color !== 'transparent') style.fc = color;
}

function applyFortuneToggle(
  style: RichTextRunStyle,
  key: 'cl',
  value: string | undefined,
): void {
  const toggle = Number(value);
  if (Number.isSafeInteger(toggle) && (toggle === 0 || toggle === 1)) {
    style[key] = toggle;
  }
}

function applyFortuneUnderline(
  style: RichTextRunStyle,
  value: string | undefined,
): void {
  const underline = Number(value);
  if (Number.isSafeInteger(underline) && underline >= 0 && underline <= 4) {
    style.un = underline as 0 | 1 | 2 | 3 | 4;
  }
}

function applyHtmlFontSize(
  style: RichTextRunStyle,
  value: string | null,
): void {
  const index = Number(value);
  const size = [0, 8, 10, 12, 14, 18, 24, 36][index];
  if (size) style.fs = size;
}

function spreadsheetRichTextRunHasFormatting(run: XlsxRichTextRun): boolean {
  return ['bl', 'cl', 'fc', 'ff', 'fs', 'it', 'un'].some(
    (key) => run[key as keyof XlsxRichTextRun] !== undefined,
  );
}

function trimTrailingClipboardNewline(
  runs: XlsxRichTextRun[],
  expectedText: string,
): void {
  const actual = runs.map((run) => run.v).join('');
  if (!actual.endsWith('\n') || expectedText.endsWith('\n')) return;
  const last = runs.at(-1);
  if (!last) return;
  last.v = last.v.slice(0, -1);
  if (!last.v) runs.pop();
}

function hasFollowingClipboardContent(element: HTMLElement): boolean {
  let sibling = element.nextSibling;
  while (sibling) {
    if (sibling.nodeType === Node.TEXT_NODE) {
      if (sibling.textContent) return true;
    } else if (
      sibling instanceof HTMLElement &&
      !ignoredClipboardElements.has(sibling.tagName)
    ) {
      return true;
    }
    sibling = sibling.nextSibling;
  }
  return false;
}

function spreadsheetRichTextCellText(cell: Cell): string | null {
  const rich = normalizeXlsxRichTextCell(cell);
  if (typeof cell.v === 'string' && cell.v !== rich?.text) return cell.v;
  return rich?.text ?? (typeof cell.v === 'string' ? cell.v : null);
}

function spreadsheetCellAt(
  sheet: WorkSpreadsheetSheet,
  row: number,
  column: number,
): Cell | null {
  return (
    sheet.data?.[row]?.[column] ??
    sheet.celldata?.find(
      (candidate) => candidate.r === row && candidate.c === column,
    )?.v ??
    null
  );
}

function focusedAxisIndex(value: unknown, axis: readonly number[]): number {
  const minimum = finiteIndex(axis[0], 0);
  const maximum = finiteIndex(axis[1], minimum);
  const focus = finiteIndex(value, minimum);
  return Math.min(
    Math.max(minimum, maximum),
    Math.max(Math.min(minimum, maximum), focus),
  );
}

function finiteIndex(value: unknown, fallback: number): number {
  return typeof value === 'number' && Number.isFinite(value)
    ? Math.max(0, Math.trunc(value))
    : fallback;
}

function normalizeClipboardText(value: string): string {
  return value.replace(/\r\n?/g, '\n').replace(/\u00a0/g, ' ');
}

function clearPendingPaste(sheet: WorkSpreadsheetSheet, key: string): void {
  const entries = pendingPastes.get(sheet);
  entries?.delete(key);
  if (entries && !entries.size) pendingPastes.delete(sheet);
}

function spreadsheetRichTextPasteKey(row: number, column: number): string {
  return `${row}:${column}`;
}

function validCoordinate(value: number): boolean {
  return Number.isSafeInteger(value) && value >= 0;
}
