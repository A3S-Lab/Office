import type { Cell } from '@fortune-sheet/core';
import { attribute, directChild, directChildren } from './work-ooxml-package';
import { spreadsheetTextOrientationXlsxValueFromCell } from './work-spreadsheet-text-orientation';
import {
  type SpreadsheetUnderlineStyle,
  spreadsheetUnderlineCellValueFromXlsx,
  spreadsheetUnderlineStyle,
} from './work-spreadsheet-underline';
import type { XlsxCellBorder } from './work-xlsx-cell-borders';
import {
  type XlsxSemanticColorOrigin,
  type XlsxSemanticPalette,
  xlsxColorElementMatchesOrigin,
  xlsxSemanticColorMatchesValue,
  xlsxSemanticColorOriginSupported,
} from './work-xlsx-cell-style-origin';
import {
  type XlsxDirectAlignmentStyle,
  xlsxRgbColor,
} from './work-xlsx-cell-style-xml';
import { resolveXlsxColor, type XlsxColorResolver } from './work-xlsx-colors';

export interface XlsxDirectFontStyle {
  bold?: boolean;
  color?: string;
  italic?: boolean;
  name?: string;
  size?: number;
  strike?: boolean;
  underline?: SpreadsheetUnderlineStyle;
}

export function hasXlsxDirectFontStyle(cell: Cell): boolean {
  return Boolean(
    cell.bl !== undefined ||
      cell.it !== undefined ||
      cell.un !== undefined ||
      cell.cl !== undefined ||
      cell.ff !== undefined ||
      cell.fs !== undefined ||
      cell.fc !== undefined,
  );
}

export function directXlsxFontStyle(cell: Cell): XlsxDirectFontStyle {
  const color = cell.fc !== undefined ? xlsxRgbColor(cell.fc) : null;
  return {
    ...(cell.bl !== undefined ? { bold: Number(cell.bl) === 1 } : {}),
    ...(color ? { color } : {}),
    ...(cell.it !== undefined ? { italic: Number(cell.it) === 1 } : {}),
    ...(typeof cell.ff === 'string' && cell.ff.trim()
      ? { name: cell.ff.trim() }
      : {}),
    ...(typeof cell.fs === 'number' && Number.isFinite(cell.fs) && cell.fs > 0
      ? { size: cell.fs }
      : {}),
    ...(cell.cl !== undefined ? { strike: Number(cell.cl) === 1 } : {}),
    ...(cell.un !== undefined
      ? { underline: spreadsheetUnderlineStyle(cell.un) }
      : {}),
  };
}

export function directXlsxAlignment(
  cell: Cell,
): XlsxDirectAlignmentStyle | null {
  const alignment: XlsxDirectAlignmentStyle = {};
  if (cell.ht !== undefined) {
    alignment.horizontal =
      Number(cell.ht) === 0
        ? 'center'
        : Number(cell.ht) === 2
          ? 'right'
          : 'left';
  }
  if (cell.vt !== undefined) {
    alignment.vertical =
      Number(cell.vt) === 0
        ? 'center'
        : Number(cell.vt) === 1
          ? 'top'
          : 'bottom';
  }
  if (cell.tb !== undefined) alignment.wrapText = cell.tb === '2';
  const rotation = spreadsheetTextOrientationXlsxValueFromCell(cell);
  if (rotation !== null) alignment.textRotation = rotation;
  return Object.keys(alignment).length ? alignment : null;
}

export function xlsxStyleCollectionIndex(
  collection: Element,
  childName: string,
  value: string | null,
): number {
  const index = nonNegativeInteger(value) ?? 0;
  return directChildren(collection, childName)[index] ? index : 0;
}

export function xlsxColorMatches(
  element: Element | undefined,
  value: string,
  colors: XlsxColorResolver,
  fallback?: string,
): boolean {
  const rgb = xlsxRgbColor(value);
  const resolved = resolveXlsxColor(element, colors) ?? fallback;
  return Boolean(
    rgb &&
      resolved &&
      `#${rgb.slice(-6).toLowerCase()}` === resolved.toLowerCase(),
  );
}

export function xlsxToggleEnabled(element: Element | undefined): boolean {
  if (!element) return false;
  const value = attribute(element, 'val')?.trim().toLowerCase();
  return value !== '0' && value !== 'false' && value !== 'off';
}

export function xlsxUnderlineStyle(
  element: Element | undefined,
): SpreadsheetUnderlineStyle {
  return element
    ? spreadsheetUnderlineStyle(
        spreadsheetUnderlineCellValueFromXlsx(attribute(element, 'val')),
      )
    : 'none';
}

export function xlsxBorderLineMatches(
  element: Element | undefined,
  line: NonNullable<XlsxCellBorder['top']> | null,
  colors: XlsxColorResolver,
  semanticColor?: XlsxSemanticColorOrigin,
): boolean {
  const style = element ? attribute(element, 'style') : null;
  if (!line) return !style;
  return (
    style === line.style &&
    (semanticColor
      ? xlsxColorElementMatchesOrigin(
          element ? directChild(element, 'color') : undefined,
          semanticColor,
        )
      : xlsxColorMatches(
          element ? directChild(element, 'color') : undefined,
          line.color,
          colors,
          '#000000',
        ))
  );
}

export function activeXlsxSemanticColorOrigin(
  origin: XlsxSemanticColorOrigin | undefined,
  value: unknown,
  palette: XlsxSemanticPalette | undefined,
): XlsxSemanticColorOrigin | undefined {
  return origin &&
    xlsxSemanticColorMatchesValue(origin, value) &&
    xlsxSemanticColorOriginSupported(origin, palette)
    ? origin
    : undefined;
}

export function xlsxBooleanAttribute(
  element: Element | undefined,
  name: string,
): boolean {
  if (!element) return false;
  const value = attribute(element, name)?.trim().toLowerCase();
  return value === '1' || value === 'true' || value === 'on';
}

export function xlsxAlignmentMatches(
  xf: Element,
  style: XlsxDirectAlignmentStyle,
): boolean {
  const alignment = directChild(xf, 'alignment');
  if (
    style.horizontal !== undefined &&
    attribute(alignment ?? xf, 'horizontal') !== style.horizontal
  ) {
    return false;
  }
  if (
    style.vertical !== undefined &&
    attribute(alignment ?? xf, 'vertical') !== style.vertical
  ) {
    return false;
  }
  if (
    style.wrapText !== undefined &&
    (!alignment ||
      attribute(alignment, 'wrapText') === null ||
      xlsxBooleanAttribute(alignment, 'wrapText') !== style.wrapText)
  ) {
    return false;
  }
  if (style.textRotation !== undefined) {
    const rotation = nonNegativeInteger(
      alignment ? attribute(alignment, 'textRotation') : null,
    );
    if (rotation !== style.textRotation) return false;
  }
  return true;
}

function nonNegativeInteger(value: string | null): number | null {
  if (value === null || !/^\d+$/.test(value)) return null;
  const parsed = Number(value);
  return Number.isSafeInteger(parsed) ? parsed : null;
}
