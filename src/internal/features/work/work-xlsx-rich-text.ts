import type { Cell, Sheet } from '@fortune-sheet/core';
import { sparseArrayEntries } from './spreadsheet-sparse';
import {
  spreadsheetUnderlineCellValue,
  spreadsheetUnderlineCellValueFromXlsx,
  spreadsheetUnderlineStyle,
  type SpreadsheetUnderlineStyle,
} from './work-spreadsheet-underline';
import {
  attribute,
  descendants,
  directChild,
  directChildren,
} from './work-ooxml-package';
import {
  applyXlsxSemanticColorOrigin,
  normalizeXlsxSemanticColorOrigin,
  readXlsxSemanticColorOrigin,
  xlsxSemanticColorMatchesValue,
  type XlsxCellStyleOrigin,
  type XlsxSemanticPalette,
} from './work-xlsx-cell-style-origin';
import { activeXlsxSemanticColorOrigin } from './work-xlsx-cell-style-values';
import { xlsxRgbColor } from './work-xlsx-cell-style-xml';
import {
  createXlsxColorResolver,
  resolveXlsxColor,
  type XlsxColorResolver,
} from './work-xlsx-colors';
import { decodeXlsxCellAddress, xlsxCellAddress } from './work-xlsx-worksheet';
import {
  MAX_XLSX_RICH_TEXT_CELL_CHARACTERS,
  MAX_XLSX_RICH_TEXT_CELLS,
  MAX_XLSX_RICH_TEXT_FONT_NAME_CHARACTERS,
  MAX_XLSX_RICH_TEXT_FONT_SIZE,
  MAX_XLSX_RICH_TEXT_RUNS,
  MAX_XLSX_RICH_TEXT_RUNS_PER_CELL,
  normalizeXlsxRichTextCell,
  normalizeXlsxRichTextColor,
  type NormalizedXlsxRichTextCell,
  type XlsxRichTextRun,
  validXlsxRichText,
} from './work-xlsx-rich-text-model';

export {
  MAX_XLSX_RICH_TEXT_CELL_CHARACTERS,
  MAX_XLSX_RICH_TEXT_CELLS,
  MAX_XLSX_RICH_TEXT_RUNS,
  MAX_XLSX_RICH_TEXT_RUNS_PER_CELL,
  type XlsxRichTextRun,
} from './work-xlsx-rich-text-model';

const MAX_XLSX_SHARED_RICH_TEXT_ITEMS = 10_000;

export interface XlsxRichTextCell {
  column: number;
  row: number;
  runs: XlsxRichTextRun[];
  text: string;
}

export interface XlsxRichTextReadContext {
  readonly colors: XlsxColorResolver;
  readonly hasRichSharedStrings: boolean;
  readonly sharedStrings: ReadonlyMap<number, readonly XlsxRichTextRun[]>;
  remainingCells: number;
  remainingRuns: number;
}

interface XlsxRichTextContextOptions {
  sharedStrings: Document | null;
  styles: Document | null;
  theme: Document | null;
}

interface SpreadsheetRichTextFontPatch {
  bold?: boolean;
  fontColor?: string;
  fontFamily?: string;
  fontSize?: number;
  italic?: boolean;
  strike?: boolean;
  underline?: SpreadsheetUnderlineStyle;
}

export function createXlsxRichTextReadContext(
  options: XlsxRichTextContextOptions,
): XlsxRichTextReadContext {
  const colors = createXlsxColorResolver(options.styles, options.theme);
  const sharedStrings = readRichSharedStrings(options.sharedStrings, colors);
  return {
    colors,
    hasRichSharedStrings: sharedStrings.size > 0,
    remainingCells: MAX_XLSX_RICH_TEXT_CELLS,
    remainingRuns: MAX_XLSX_RICH_TEXT_RUNS,
    sharedStrings,
  };
}

export function readXlsxRichTextCells(
  worksheet: Document,
  context: XlsxRichTextReadContext,
): XlsxRichTextCell[] {
  const result: XlsxRichTextCell[] = [];
  const seen = new Set<string>();
  for (const element of descendants(worksheet, 'c')) {
    if (context.remainingCells <= 0 || context.remainingRuns <= 0) break;
    if (directChild(element, 'f')) continue;
    const reference = attribute(element, 'r');
    if (!reference) continue;
    const coordinate = decodeXlsxCellAddress(reference);
    if (
      !coordinate ||
      coordinate.column > 16_383 ||
      coordinate.row > 1_048_575 ||
      seen.has(reference)
    ) {
      continue;
    }
    seen.add(reference);

    const type = attribute(element, 't');
    let source: readonly XlsxRichTextRun[] | null = null;
    if (type === 's') {
      const index = nonNegativeInteger(directChild(element, 'v')?.textContent);
      source =
        index === null ? null : (context.sharedStrings.get(index) ?? null);
    } else if (type === 'inlineStr') {
      const inlineString = directChild(element, 'is');
      source = inlineString
        ? readRichTextRuns(inlineString, context.colors)
        : null;
    }
    if (!source?.length || source.length > context.remainingRuns) {
      continue;
    }
    const runs = source.map((run) => ({ ...run }));
    const text = runs.map((run) => run.v).join('');
    if (!text || text.length > MAX_XLSX_RICH_TEXT_CELL_CHARACTERS) continue;
    context.remainingCells -= 1;
    context.remainingRuns -= runs.length;
    result.push({ ...coordinate, runs, text });
  }
  return result;
}

export function applyImportedXlsxRichText(
  cell: Cell,
  richText: XlsxRichTextCell | undefined,
): Cell {
  if (!richText) return cell;
  const next: Cell = {
    ...cell,
    ct: {
      ...cell.ct,
      s: richText.runs.map((run) => ({ ...run })),
      t: 'inlineStr',
    },
    v: richText.text,
  };
  delete next.m;
  return next;
}

export function patchSpreadsheetRichTextFontRuns(
  cell: Cell,
  patch: SpreadsheetRichTextFontPatch,
): Cell {
  if (!fontPatchHasValues(patch) || cell.ct?.t !== 'inlineStr') return cell;
  const source = cell.ct.s;
  if (
    !Array.isArray(source) ||
    !source.length ||
    source.some(
      (run) =>
        !isRecord(run) ||
        typeof run.v !== 'string' ||
        !validXlsxRichText(run.v),
    )
  ) {
    return cell;
  }
  const normalizedColor =
    patch.fontColor === undefined
      ? undefined
      : normalizeXlsxRichTextColor(patch.fontColor);
  const runs = source.map((run) => {
    const next = { ...run } as Record<string, unknown>;
    if (patch.fontFamily !== undefined) next.ff = patch.fontFamily.trim();
    if (patch.fontSize !== undefined) next.fs = patch.fontSize;
    if (normalizedColor !== undefined) {
      next.fc = normalizedColor;
      delete next.a3sXlsxColorOrigin;
    }
    if (patch.bold !== undefined) next.bl = patch.bold ? 1 : 0;
    if (patch.italic !== undefined) next.it = patch.italic ? 1 : 0;
    if (patch.underline !== undefined) {
      next.un = spreadsheetUnderlineCellValue(patch.underline);
    }
    if (patch.strike !== undefined) next.cl = patch.strike ? 1 : 0;
    return next;
  });
  return { ...cell, ct: { ...cell.ct, s: runs } };
}

export function xlsxRichTextCellText(cell: Cell): string | null {
  return normalizeXlsxRichTextCell(cell)?.text ?? null;
}

export function sheetHasXlsxRichTextCells(sheet: Sheet): boolean {
  for (const [, row] of sparseArrayEntries(sheet.data)) {
    for (const [, cell] of sparseArrayEntries(row)) {
      if (cell && normalizeXlsxRichTextCell(cell)) return true;
    }
  }
  return false;
}

export function xlsxRichTextStyleOrigins(sheet: Sheet): XlsxCellStyleOrigin[] {
  const origins: XlsxCellStyleOrigin[] = [];
  for (const [, row] of sparseArrayEntries(sheet.data)) {
    for (const [, cell] of sparseArrayEntries(row)) {
      if (!cell) continue;
      for (const run of normalizeXlsxRichTextCell(cell)?.runs ?? []) {
        const fontColor = normalizeXlsxSemanticColorOrigin(
          run.a3sXlsxColorOrigin,
        );
        if (
          fontColor &&
          run.fc &&
          xlsxSemanticColorMatchesValue(fontColor, run.fc)
        ) {
          origins.push({ fontColor });
        }
      }
    }
  }
  return origins;
}

export function writeXlsxRichTextCells(
  worksheet: Document,
  sheet: Sheet,
  semanticPalette?: XlsxSemanticPalette,
): void {
  const elements = new Map(
    descendants(worksheet, 'c').flatMap((element) => {
      const reference = attribute(element, 'r');
      return reference ? [[reference, element] as const] : [];
    }),
  );
  let remainingCells = MAX_XLSX_RICH_TEXT_CELLS;
  let remainingRuns = MAX_XLSX_RICH_TEXT_RUNS;
  for (const [row, values] of sparseArrayEntries(sheet.data)) {
    for (const [column, cell] of sparseArrayEntries(values)) {
      if (remainingCells <= 0 || remainingRuns <= 0) return;
      if (!cell) continue;
      const richText = normalizeXlsxRichTextCell(cell);
      if (!richText || richText.runs.length > remainingRuns) continue;
      const element = elements.get(xlsxCellAddress(row, column));
      if (!element) continue;
      writeRichTextCell(element, richText, semanticPalette);
      remainingCells -= 1;
      remainingRuns -= richText.runs.length;
    }
  }
}

function readRichSharedStrings(
  document: Document | null,
  colors: XlsxColorResolver,
): ReadonlyMap<number, readonly XlsxRichTextRun[]> {
  if (!document) return new Map();
  const result = new Map<number, readonly XlsxRichTextRun[]>();
  let parsedRuns = 0;
  for (const [index, item] of directChildren(
    document.documentElement,
    'si',
  ).entries()) {
    if (
      result.size >= MAX_XLSX_SHARED_RICH_TEXT_ITEMS ||
      parsedRuns >= MAX_XLSX_RICH_TEXT_RUNS
    ) {
      break;
    }
    const runs = readRichTextRuns(item, colors);
    if (!runs || parsedRuns + runs.length > MAX_XLSX_RICH_TEXT_RUNS) continue;
    result.set(index, runs);
    parsedRuns += runs.length;
  }
  return result;
}

function readRichTextRuns(
  container: Element,
  colors: XlsxColorResolver,
): XlsxRichTextRun[] | null {
  const elements = directChildren(container, 'r');
  if (
    directChild(container, 't') ||
    !elements.length ||
    elements.length > MAX_XLSX_RICH_TEXT_RUNS_PER_CELL
  ) {
    return null;
  }
  const runs: XlsxRichTextRun[] = [];
  let characterCount = 0;
  for (const element of elements) {
    const textElement = directChild(element, 't');
    if (
      !textElement ||
      textElement.children.length ||
      directChildren(element, 't').length !== 1
    ) {
      return null;
    }
    const value = textElement.textContent ?? '';
    if (!value) continue;
    characterCount += value.length;
    if (
      characterCount > MAX_XLSX_RICH_TEXT_CELL_CHARACTERS ||
      !validXlsxRichText(value)
    ) {
      return null;
    }
    runs.push(readRichTextRun(element, value, colors));
  }
  return runs.length ? runs : null;
}

function readRichTextRun(
  element: Element,
  value: string,
  colors: XlsxColorResolver,
): XlsxRichTextRun {
  const properties = directChild(element, 'rPr');
  if (!properties) return { v: value };
  const run: XlsxRichTextRun = { v: value };
  const font =
    directChild(properties, 'rFont') ?? directChild(properties, 'name');
  const fontName = attribute(font ?? properties, 'val')?.trim();
  if (fontName && fontName.length <= MAX_XLSX_RICH_TEXT_FONT_NAME_CHARACTERS) {
    run.ff = fontName;
  }
  if (xlsxToggleEnabled(directChild(properties, 'b'))) run.bl = 1;
  if (xlsxToggleEnabled(directChild(properties, 'i'))) run.it = 1;
  if (xlsxToggleEnabled(directChild(properties, 'strike'))) run.cl = 1;
  const size = finiteNumber(
    attribute(directChild(properties, 'sz') ?? properties, 'val'),
  );
  if (size !== null && size >= 1 && size <= MAX_XLSX_RICH_TEXT_FONT_SIZE) {
    run.fs = size;
  }
  const underline = directChild(properties, 'u');
  if (underline) {
    const value = spreadsheetUnderlineCellValueFromXlsx(
      attribute(underline, 'val'),
    );
    if (value) run.un = value;
  }
  const colorElement = directChild(properties, 'color');
  const color = resolveXlsxColor(colorElement, colors);
  if (color) run.fc = color;
  const colorOrigin = readXlsxSemanticColorOrigin(colorElement, colors);
  if (colorOrigin) run.a3sXlsxColorOrigin = colorOrigin;
  return run;
}

function writeRichTextCell(
  element: Element,
  richText: NormalizedXlsxRichTextCell,
  semanticPalette: XlsxSemanticPalette | undefined,
): void {
  for (const child of directChildren(element)) {
    if (child.localName === 'v' || child.localName === 'is') child.remove();
  }
  element.setAttribute('t', 'inlineStr');
  const document = element.ownerDocument;
  const namespace = document.documentElement.namespaceURI;
  const inlineString = document.createElementNS(namespace, 'is');
  for (const run of richText.runs) {
    const runElement = document.createElementNS(namespace, 'r');
    const properties = writeRichTextRunProperties(
      document,
      run,
      semanticPalette,
    );
    if (properties) runElement.append(properties);
    const text = document.createElementNS(namespace, 't');
    if (/^\s|\s$/.test(run.v)) {
      text.setAttributeNS(
        'http://www.w3.org/XML/1998/namespace',
        'xml:space',
        'preserve',
      );
    }
    text.textContent = run.v;
    runElement.append(text);
    inlineString.append(runElement);
  }
  element.insertBefore(
    inlineString,
    directChildren(element).find((child) => child.localName === 'extLst') ??
      null,
  );
}

function writeRichTextRunProperties(
  document: Document,
  run: XlsxRichTextRun,
  semanticPalette: XlsxSemanticPalette | undefined,
): Element | null {
  const namespace = document.documentElement.namespaceURI;
  const properties = document.createElementNS(namespace, 'rPr');
  const appendValue = (name: string, value: string) => {
    const element = document.createElementNS(namespace, name);
    element.setAttribute('val', value);
    properties.append(element);
  };
  const appendToggle = (name: string, enabled: boolean) => {
    if (!enabled) return;
    const element = document.createElementNS(namespace, name);
    element.setAttribute('val', '1');
    properties.append(element);
  };
  if (run.ff) appendValue('rFont', run.ff);
  appendToggle('b', run.bl === 1);
  appendToggle('i', run.it === 1);
  appendToggle('strike', run.cl === 1);
  if (run.fc) {
    const color = document.createElementNS(namespace, 'color');
    const semanticOrigin = activeXlsxSemanticColorOrigin(
      normalizeXlsxSemanticColorOrigin(run.a3sXlsxColorOrigin),
      run.fc,
      semanticPalette,
    );
    if (semanticOrigin) applyXlsxSemanticColorOrigin(color, semanticOrigin);
    else {
      const rgb = xlsxRgbColor(run.fc);
      if (rgb) color.setAttribute('rgb', rgb);
    }
    if (color.attributes.length) properties.append(color);
  }
  if (run.fs !== undefined) appendValue('sz', String(run.fs));
  if (run.un !== undefined && run.un !== 0) {
    appendValue('u', spreadsheetUnderlineStyle(run.un));
  }
  return properties.children.length ? properties : null;
}

function fontPatchHasValues(patch: SpreadsheetRichTextFontPatch): boolean {
  return (
    patch.fontFamily !== undefined ||
    patch.fontSize !== undefined ||
    patch.fontColor !== undefined ||
    patch.bold !== undefined ||
    patch.italic !== undefined ||
    patch.underline !== undefined ||
    patch.strike !== undefined
  );
}

function xlsxToggleEnabled(element: Element | undefined): boolean {
  if (!element) return false;
  const value = attribute(element, 'val')?.trim().toLowerCase();
  return value !== '0' && value !== 'false' && value !== 'off';
}

function finiteNumber(value: string | null): number | null {
  if (value === null || !value.trim()) return null;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : null;
}

function nonNegativeInteger(value: string | null | undefined): number | null {
  if (value === null || value === undefined || !/^\d+$/.test(value.trim())) {
    return null;
  }
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) ? parsed : null;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
