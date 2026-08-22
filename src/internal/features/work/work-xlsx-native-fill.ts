import type { Cell } from '@fortune-sheet/core';
import {
  activeXlsxGradientFill,
  deleteXlsxGradientFill,
  XLSX_GRADIENT_FILL_CELL_KEY,
  xlsxGradientFillKey,
  xlsxGradientFillSemanticColors,
  type XlsxGradientFill,
} from './work-xlsx-gradient-fill';
import {
  activeXlsxPatternFill,
  deleteXlsxPatternFill,
  XLSX_PATTERN_FILL_CELL_KEY,
  xlsxPatternFillKey,
  xlsxPatternFillSemanticColors,
  type XlsxPatternFill,
} from './work-xlsx-pattern-fill';
import type { XlsxSemanticColorOrigin } from './work-xlsx-cell-style-origin';

export const xlsxNativeFillCellKeys = [
  XLSX_PATTERN_FILL_CELL_KEY,
  XLSX_GRADIENT_FILL_CELL_KEY,
] as const;

export type XlsxNativeFillCellKey = (typeof xlsxNativeFillCellKeys)[number];

export type XlsxNativeFill =
  | { kind: 'gradient'; value: XlsxGradientFill }
  | { kind: 'pattern'; value: XlsxPatternFill };

export function activeXlsxNativeFill(
  cell: Cell | null | undefined,
): XlsxNativeFill | undefined {
  const gradient = activeXlsxGradientFill(cell);
  const pattern = activeXlsxPatternFill(cell);
  if (Boolean(gradient) === Boolean(pattern)) return undefined;
  return gradient
    ? { kind: 'gradient', value: gradient }
    : { kind: 'pattern', value: pattern! };
}

export function deleteXlsxNativeFills(cell: Cell): void {
  deleteXlsxGradientFill(cell);
  deleteXlsxPatternFill(cell);
}

export function xlsxNativeFillSemanticColors(
  fill: XlsxNativeFill | undefined,
): XlsxSemanticColorOrigin[] {
  if (!fill) return [];
  return fill.kind === 'gradient'
    ? xlsxGradientFillSemanticColors(fill.value)
    : xlsxPatternFillSemanticColors(fill.value);
}

export function xlsxNativeFillKey(fill: XlsxNativeFill | undefined): string {
  if (!fill) return '';
  return fill.kind === 'gradient'
    ? `gradient:${xlsxGradientFillKey(fill.value)}`
    : `pattern:${xlsxPatternFillKey(fill.value)}`;
}
