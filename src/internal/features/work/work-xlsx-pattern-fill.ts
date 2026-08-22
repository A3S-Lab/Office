import type { Cell } from '@fortune-sheet/core';
import { attribute, directChild } from './work-ooxml-package';
import {
  normalizeXlsxSemanticColorOrigin,
  readXlsxSemanticColorOrigin,
  type XlsxSemanticColorOrigin,
  xlsxSemanticColorOriginKey,
} from './work-xlsx-cell-style-origin';
import { xlsxRgbColor } from './work-xlsx-cell-style-xml';
import { resolveXlsxColor, type XlsxColorResolver } from './work-xlsx-colors';

export const XLSX_PATTERN_FILL_CELL_KEY = 'a3sXlsxPatternFill' as const;

export const xlsxPatternFillTypes = [
  'darkDown',
  'darkGray',
  'darkGrid',
  'darkHorizontal',
  'darkTrellis',
  'darkUp',
  'darkVertical',
  'gray0625',
  'gray125',
  'lightDown',
  'lightGray',
  'lightGrid',
  'lightHorizontal',
  'lightTrellis',
  'lightUp',
  'lightVertical',
  'mediumGray',
] as const;

export type XlsxPatternFillType = (typeof xlsxPatternFillTypes)[number];

export interface XlsxPatternFill {
  backgroundColor: string;
  backgroundColorOrigin?: XlsxSemanticColorOrigin;
  foregroundColor: string;
  foregroundColorOrigin?: XlsxSemanticColorOrigin;
  patternType: XlsxPatternFillType;
}

type CellWithXlsxPatternFill = Cell & {
  [XLSX_PATTERN_FILL_CELL_KEY]?: unknown;
};

const xlsxPatternFillTypeSet = new Set<string>(xlsxPatternFillTypes);

export function readXlsxPatternFill(
  pattern: Element | undefined,
  colors: XlsxColorResolver,
): XlsxPatternFill | undefined {
  if (!pattern) return undefined;
  const patternType = attribute(pattern, 'patternType');
  if (!patternType || !isXlsxPatternFillType(patternType)) {
    return undefined;
  }
  const foreground = directChild(pattern, 'fgColor');
  const background = directChild(pattern, 'bgColor');
  const foregroundColor = foreground
    ? resolveXlsxColor(foreground, colors)
    : '#000000';
  const backgroundColor = background
    ? resolveXlsxColor(background, colors)
    : '#ffffff';
  const normalizedForegroundColor = normalizedColor(foregroundColor);
  const normalizedBackgroundColor = normalizedColor(backgroundColor);
  if (!normalizedForegroundColor || !normalizedBackgroundColor) {
    return undefined;
  }
  const foregroundColorOrigin = readXlsxSemanticColorOrigin(foreground, colors);
  const backgroundColorOrigin = readXlsxSemanticColorOrigin(background, colors);
  return {
    backgroundColor: normalizedBackgroundColor,
    ...(backgroundColorOrigin ? { backgroundColorOrigin } : {}),
    foregroundColor: normalizedForegroundColor,
    ...(foregroundColorOrigin ? { foregroundColorOrigin } : {}),
    patternType,
  };
}

export function withXlsxPatternFill(
  cell: Cell,
  fill: XlsxPatternFill | undefined,
): Cell {
  const normalized = normalizeXlsxPatternFill(fill);
  return normalized
    ? ({
        ...cell,
        [XLSX_PATTERN_FILL_CELL_KEY]: normalized,
      } as CellWithXlsxPatternFill)
    : cell;
}

export function deleteXlsxPatternFill(cell: Cell): void {
  delete (cell as CellWithXlsxPatternFill)[XLSX_PATTERN_FILL_CELL_KEY];
}

export function xlsxPatternFill(
  cell: Cell | null | undefined,
): XlsxPatternFill | undefined {
  return normalizeXlsxPatternFill(
    (cell as CellWithXlsxPatternFill | null | undefined)?.[
      XLSX_PATTERN_FILL_CELL_KEY
    ],
  );
}

export function activeXlsxPatternFill(
  cell: Cell | null | undefined,
): XlsxPatternFill | undefined {
  const fill = xlsxPatternFill(cell);
  return fill && normalizedColor(cell?.bg) === fill.backgroundColor
    ? fill
    : undefined;
}

export function normalizeXlsxPatternFill(
  value: unknown,
): XlsxPatternFill | undefined {
  if (!isRecord(value) || !isXlsxPatternFillType(value.patternType)) {
    return undefined;
  }
  const foregroundColor = normalizedColor(value.foregroundColor);
  const backgroundColor = normalizedColor(value.backgroundColor);
  if (!foregroundColor || !backgroundColor) return undefined;
  const foregroundColorOrigin = normalizeXlsxSemanticColorOrigin(
    value.foregroundColorOrigin,
  );
  const backgroundColorOrigin = normalizeXlsxSemanticColorOrigin(
    value.backgroundColorOrigin,
  );
  return {
    backgroundColor,
    ...(backgroundColorOrigin ? { backgroundColorOrigin } : {}),
    foregroundColor,
    ...(foregroundColorOrigin ? { foregroundColorOrigin } : {}),
    patternType: value.patternType,
  };
}

export function xlsxPatternFillSemanticColors(
  fill: XlsxPatternFill | undefined,
): XlsxSemanticColorOrigin[] {
  return fill
    ? [fill.foregroundColorOrigin, fill.backgroundColorOrigin].filter(
        (value): value is XlsxSemanticColorOrigin => Boolean(value),
      )
    : [];
}

export function xlsxPatternFillKey(fill: XlsxPatternFill | undefined): string {
  return fill
    ? [
        fill.patternType,
        fill.foregroundColor,
        xlsxSemanticColorOriginKey(fill.foregroundColorOrigin),
        fill.backgroundColor,
        xlsxSemanticColorOriginKey(fill.backgroundColorOrigin),
      ].join(':')
    : '';
}

function isXlsxPatternFillType(value: unknown): value is XlsxPatternFillType {
  return typeof value === 'string' && xlsxPatternFillTypeSet.has(value);
}

function normalizedColor(value: unknown): string | null {
  const rgb = xlsxRgbColor(value);
  return rgb ? `#${rgb.slice(-6).toLowerCase()}` : null;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
