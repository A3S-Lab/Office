import type { Cell } from '@fortune-sheet/core';
import {
  normalizeXlsxGradientFill,
  withXlsxGradientFill,
  xlsxGradientFillFallbackColor,
} from '../work-xlsx-gradient-fill';
import {
  activeXlsxNativeFill,
  deleteXlsxNativeFills,
  type XlsxNativeFill,
} from '../work-xlsx-native-fill';
import {
  normalizeXlsxPatternFill,
  withXlsxPatternFill,
} from '../work-xlsx-pattern-fill';

export type SpreadsheetCellFillFormat =
  | { kind: 'none' }
  | { color: string; kind: 'solid' }
  | XlsxNativeFill;

export function spreadsheetCellFillFormat(
  cell: Cell | null | undefined,
): SpreadsheetCellFillFormat {
  const nativeFill = activeXlsxNativeFill(cell);
  if (nativeFill) return nativeFill;
  const color = normalizeSpreadsheetFillColor(cell?.bg);
  return color ? { color, kind: 'solid' } : { kind: 'none' };
}

export function normalizeSpreadsheetCellFillFormat(
  value: unknown,
): SpreadsheetCellFillFormat | null {
  if (!isRecord(value)) return null;
  if (value.kind === 'none') return { kind: 'none' };
  if (value.kind === 'solid') {
    const color = normalizeSpreadsheetFillColor(value.color);
    return color ? { color, kind: 'solid' } : null;
  }
  if (value.kind === 'pattern') {
    const fill = normalizeXlsxPatternFill(value.value);
    return fill ? { kind: 'pattern', value: fill } : null;
  }
  if (value.kind === 'gradient') {
    const fill = normalizeXlsxGradientFill(value.value);
    return fill ? { kind: 'gradient', value: fill } : null;
  }
  return null;
}

export function withSpreadsheetCellFillFormat(
  cell: Cell,
  fill: SpreadsheetCellFillFormat,
): Cell | null {
  const normalized = normalizeSpreadsheetCellFillFormat(fill);
  if (!normalized) return null;
  const next = { ...cell };
  deleteXlsxNativeFills(next);
  if (normalized.kind === 'none') {
    delete next.bg;
    return next;
  }
  if (normalized.kind === 'solid') {
    next.bg = normalized.color;
    return next;
  }
  if (normalized.kind === 'pattern') {
    next.bg = normalized.value.backgroundColor;
    return withXlsxPatternFill(next, normalized.value);
  }
  next.bg = xlsxGradientFillFallbackColor(normalized.value);
  return withXlsxGradientFill(next, normalized.value);
}

function normalizeSpreadsheetFillColor(value: unknown): string | null {
  if (typeof value !== 'string') return null;
  const color = value.trim().toLowerCase();
  if (/^#[0-9a-f]{6}$/.test(color)) return color;
  if (!/^#[0-9a-f]{3}$/.test(color)) return null;
  return `#${[...color.slice(1)]
    .map((character) => character.repeat(2))
    .join('')}`;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
