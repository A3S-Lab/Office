import type { Cell, Sheet } from '@fortune-sheet/core';
import { sparseArrayEntries } from './spreadsheet-sparse';
import {
  collectXlsxCellBorders,
  type XlsxCellBorder,
  type XlsxCellBorderStyle,
  xlsxCellBorderKey,
} from './work-xlsx-cell-borders';
import {
  hasXlsxDirectCellStyle,
  type XlsxDirectCellStyleWriter,
} from './work-xlsx-cell-style-writer';
import {
  createXlsxColorResolver,
  resolveXlsxColor,
  type XlsxColorResolver,
} from './work-xlsx-colors';
import {
  attribute,
  descendants,
  directChild,
  directChildren,
} from './work-ooxml-package';

export { XlsxDirectCellStyleWriter } from './work-xlsx-cell-style-writer';

const xlsxCellBorderStyles = new Set<XlsxCellBorderStyle>([
  'dashDot',
  'dashDotDot',
  'dashed',
  'dotted',
  'double',
  'hair',
  'medium',
  'mediumDashDot',
  'mediumDashDotDot',
  'mediumDashed',
  'slantDashDot',
  'thick',
  'thin',
]);

export interface XlsxDirectCellStyle {
  border?: XlsxCellBorder;
  column: number;
  row: number;
  style: Partial<Cell>;
}

export function readXlsxDirectCellStyles(
  worksheet: Document,
  styles: Document | null,
  theme: Document | null = null,
): XlsxDirectCellStyle[] {
  if (!styles) return [];
  const colors = createXlsxColorResolver(styles, theme);
  const fonts = directChildren(
    directChild(styles.documentElement, 'fonts') ?? styles,
    'font',
  );
  const fills = directChildren(
    directChild(styles.documentElement, 'fills') ?? styles,
    'fill',
  );
  const borders = directChildren(
    directChild(styles.documentElement, 'borders') ?? styles,
    'border',
  );
  const cellXfs = directChildren(
    directChild(styles.documentElement, 'cellXfs') ?? styles,
    'xf',
  );
  if (!cellXfs.length) return [];

  return descendants(worksheet, 'c').flatMap((cell) => {
    const coordinate = decodeCell(attribute(cell, 'r'));
    const styleId = nonNegativeInteger(attribute(cell, 's'));
    const xf = styleId === null ? undefined : cellXfs[styleId];
    if (!coordinate || !xf) return [];
    const style = readDirectCellStyle(xf, fonts, fills, colors);
    const border = readDirectCellBorder(xf, borders, colors);
    return Object.keys(style).length || border
      ? [{ ...coordinate, border, style }]
      : [];
  });
}

export function sheetHasDirectCellStyles(sheet: Sheet): boolean {
  if (
    Array.isArray(sheet.config?.borderInfo) &&
    sheet.config.borderInfo.length
  ) {
    return true;
  }
  for (const [, row] of sparseArrayEntries(sheet.data)) {
    for (const [, cell] of sparseArrayEntries(row)) {
      if (cell && hasXlsxDirectCellStyle(cell)) return true;
    }
  }
  return false;
}

export function writeXlsxDirectCellStyles(
  worksheet: Document,
  sheet: Sheet,
  styles: XlsxDirectCellStyleWriter,
): void {
  const borders = collectXlsxCellBorders(sheet);
  const cells = new Map(
    descendants(worksheet, 'c').flatMap((element) => {
      const reference = attribute(element, 'r');
      return reference ? [[reference, element] as const] : [];
    }),
  );
  for (const [row, values] of sparseArrayEntries(sheet.data)) {
    for (const [column, cell] of sparseArrayEntries(values)) {
      const border = borders.get(xlsxCellBorderKey(row, column));
      if (!cell || (!hasXlsxDirectCellStyle(cell) && !border)) continue;
      const element = cells.get(encodeCell(row, column));
      if (!element) continue;
      const baseStyleId = nonNegativeInteger(attribute(element, 's')) ?? 0;
      const styleId = styles.styleId(baseStyleId, cell, border);
      if (styleId) element.setAttribute('s', String(styleId));
      else element.removeAttribute('s');
    }
  }
}

function readDirectCellStyle(
  xf: Element,
  fonts: Element[],
  fills: Element[],
  colors: XlsxColorResolver,
): Partial<Cell> {
  const style: Partial<Cell> = {};
  const fontId = nonNegativeInteger(attribute(xf, 'fontId')) ?? 0;
  const font = fonts[fontId];
  if (font && (fontId !== 0 || booleanAttribute(xf, 'applyFont'))) {
    readDirectFontStyle(style, font, colors);
  }
  const fillId = nonNegativeInteger(attribute(xf, 'fillId')) ?? 0;
  const fill = fills[fillId];
  if (fill && (fillId !== 0 || booleanAttribute(xf, 'applyFill'))) {
    readDirectFillStyle(style, fill, colors);
  }
  const alignment = directChild(xf, 'alignment');
  if (alignment) readDirectAlignmentStyle(style, alignment);
  return style;
}

function readDirectFontStyle(
  style: Partial<Cell>,
  font: Element,
  colors: XlsxColorResolver,
): void {
  const bold = directChild(font, 'b');
  if (bold && toggleEnabled(bold)) style.bl = 1;
  const italic = directChild(font, 'i');
  if (italic && toggleEnabled(italic)) style.it = 1;
  const strike = directChild(font, 'strike');
  if (strike && toggleEnabled(strike)) style.cl = 1;
  const underline = directChild(font, 'u');
  if (underline && underlineEnabled(underline)) style.un = 1;

  const name = attribute(directChild(font, 'name') ?? font, 'val')?.trim();
  if (name) style.ff = name;
  const size = Number(attribute(directChild(font, 'sz') ?? font, 'val'));
  if (Number.isFinite(size) && size > 0) style.fs = size;
  const color = resolveXlsxColor(directChild(font, 'color'), colors);
  if (color) style.fc = color;
}

function readDirectFillStyle(
  style: Partial<Cell>,
  fill: Element,
  colors: XlsxColorResolver,
): void {
  const pattern = directChild(fill, 'patternFill');
  if (!pattern || attribute(pattern, 'patternType') !== 'solid') return;
  const color = resolveXlsxColor(directChild(pattern, 'fgColor'), colors);
  if (color) style.bg = color;
}

function readDirectAlignmentStyle(
  style: Partial<Cell>,
  alignment: Element,
): void {
  const horizontal = attribute(alignment, 'horizontal');
  if (horizontal === 'center') style.ht = 0;
  else if (horizontal === 'left') style.ht = 1;
  else if (horizontal === 'right') style.ht = 2;

  const vertical = attribute(alignment, 'vertical');
  if (vertical === 'center') style.vt = 0;
  else if (vertical === 'top') style.vt = 1;
  else if (vertical === 'bottom') style.vt = 2;

  if (booleanAttribute(alignment, 'wrapText')) style.tb = '2';
  const rotation = nonNegativeInteger(attribute(alignment, 'textRotation'));
  if (rotation !== null && rotation <= 180) style.tr = String(rotation);
}

function readDirectCellBorder(
  xf: Element,
  borders: Element[],
  colors: XlsxColorResolver,
): XlsxCellBorder | undefined {
  const borderId = nonNegativeInteger(attribute(xf, 'borderId')) ?? 0;
  const source = borders[borderId];
  if (!source || (borderId === 0 && !booleanAttribute(xf, 'applyBorder'))) {
    return undefined;
  }
  const border: XlsxCellBorder = {};
  readBorderLine(border, 'left', directChild(source, 'left'), colors);
  readBorderLine(border, 'right', directChild(source, 'right'), colors);
  readBorderLine(border, 'top', directChild(source, 'top'), colors);
  readBorderLine(border, 'bottom', directChild(source, 'bottom'), colors);
  readBorderLine(border, 'diagonal', directChild(source, 'diagonal'), colors);
  if (border.diagonal) {
    border.diagonalUp = booleanAttribute(source, 'diagonalUp');
    border.diagonalDown = booleanAttribute(source, 'diagonalDown');
  }
  return Object.keys(border).length ? border : undefined;
}

function readBorderLine(
  border: XlsxCellBorder,
  name: 'bottom' | 'diagonal' | 'left' | 'right' | 'top',
  element: Element | undefined,
  colors: XlsxColorResolver,
): void {
  if (!element) return;
  const style = attribute(element, 'style');
  if (!style || !xlsxCellBorderStyles.has(style as XlsxCellBorderStyle)) {
    return;
  }
  border[name] = {
    color: resolveXlsxColor(directChild(element, 'color'), colors) ?? '#000000',
    style: style as XlsxCellBorderStyle,
  };
}

function toggleEnabled(element: Element): boolean {
  const value = attribute(element, 'val')?.trim().toLowerCase();
  return value !== '0' && value !== 'false' && value !== 'off';
}

function underlineEnabled(element: Element): boolean {
  const value = attribute(element, 'val')?.trim().toLowerCase();
  return value !== '0' && value !== 'false' && value !== 'none';
}

function booleanAttribute(element: Element, name: string): boolean {
  const value = attribute(element, name)?.trim().toLowerCase();
  return value === '1' || value === 'true' || value === 'on';
}

function encodeCell(row: number, column: number): string {
  let value = column + 1;
  let label = '';
  while (value > 0) {
    value -= 1;
    label = String.fromCharCode(65 + (value % 26)) + label;
    value = Math.floor(value / 26);
  }
  return `${label}${row + 1}`;
}

function decodeCell(
  reference: string | null,
): { column: number; row: number } | null {
  const match = /^([A-Z]{1,3})([1-9]\d{0,6})$/i.exec(reference ?? '');
  if (!match) return null;
  let column = 0;
  for (const character of match[1].toUpperCase()) {
    column = column * 26 + character.charCodeAt(0) - 64;
  }
  const row = Number(match[2]);
  return column <= 16_384 && row <= 1_048_576
    ? { column: column - 1, row: row - 1 }
    : null;
}

function nonNegativeInteger(value: string | null): number | null {
  if (value === null || !/^\d+$/.test(value)) return null;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) ? parsed : null;
}
