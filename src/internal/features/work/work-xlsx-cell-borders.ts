import type { Sheet } from '@fortune-sheet/core';
import { sparseArrayEntries } from './spreadsheet-sparse';

export type XlsxCellBorderStyle =
  | 'dashDot'
  | 'dashDotDot'
  | 'dashed'
  | 'dotted'
  | 'double'
  | 'hair'
  | 'medium'
  | 'mediumDashDot'
  | 'mediumDashDotDot'
  | 'mediumDashed'
  | 'slantDashDot'
  | 'thick'
  | 'thin';

export interface XlsxCellBorderLine {
  color: string;
  style: XlsxCellBorderStyle;
}

export interface XlsxCellBorder {
  bottom?: XlsxCellBorderLine | null;
  diagonal?: XlsxCellBorderLine | null;
  diagonalDown?: boolean;
  diagonalUp?: boolean;
  left?: XlsxCellBorderLine | null;
  right?: XlsxCellBorderLine | null;
  top?: XlsxCellBorderLine | null;
}

export interface XlsxPositionedCellBorder {
  border?: XlsxCellBorder;
  column: number;
  row: number;
}

interface SpreadsheetBorderRange {
  column: readonly [number, number];
  row: readonly [number, number];
}

interface SpreadsheetBorderTarget {
  column: number;
  row: number;
}

type UnknownRecord = Record<string, unknown>;

const xlsxStyleByFortuneStyle: Readonly<Record<string, XlsxCellBorderStyle>> = {
  '1': 'thin',
  '2': 'hair',
  '3': 'dotted',
  '4': 'dashed',
  '5': 'dashDot',
  '6': 'dashDotDot',
  '7': 'double',
  '8': 'medium',
  '9': 'mediumDashed',
  '10': 'mediumDashDot',
  '11': 'mediumDashDotDot',
  '12': 'slantDashDot',
  '13': 'thick',
};

const fortuneStyleByXlsxStyle = new Map(
  Object.entries(xlsxStyleByFortuneStyle).map(([fortune, xlsx]) => [
    xlsx,
    fortune,
  ]),
);

export function collectXlsxCellBorders(
  sheet: Sheet,
): ReadonlyMap<string, XlsxCellBorder> {
  const targets = spreadsheetBorderTargets(sheet);
  const borders = new Map<string, XlsxCellBorder>();
  const source = Array.isArray(sheet.config?.borderInfo)
    ? (sheet.config.borderInfo as unknown[])
    : [];

  for (const candidate of source) {
    if (!isRecord(candidate)) continue;
    if (candidate.rangeType === 'cell') {
      applySpreadsheetCellBorderRecord(candidate, targets, borders);
      continue;
    }
    if (
      candidate.rangeType !== 'range' ||
      typeof candidate.borderType !== 'string' ||
      !Array.isArray(candidate.range)
    ) {
      continue;
    }
    const line = spreadsheetBorderLine(candidate);
    for (const rangeCandidate of candidate.range) {
      const range = spreadsheetBorderRange(rangeCandidate);
      if (!range) continue;
      for (const target of targets.values()) {
        if (!spreadsheetBorderRangeContains(range, target)) continue;
        applySpreadsheetRangeBorder(
          borders,
          target,
          range,
          candidate.borderType,
          line,
        );
      }
    }
  }
  return borders;
}

export function fortuneBorderInfoFromXlsxCells(
  entries: readonly XlsxPositionedCellBorder[],
): unknown[] {
  return entries.flatMap(({ border, column, row }) => {
    if (!border) return [];
    const value: UnknownRecord = { col_index: column, row_index: row };
    setFortuneBorderLine(value, 'l', border.left);
    setFortuneBorderLine(value, 'r', border.right);
    setFortuneBorderLine(value, 't', border.top);
    setFortuneBorderLine(value, 'b', border.bottom);
    if (border.diagonal) {
      setFortuneBorderLine(value, 's', border.diagonal);
    }
    return Object.keys(value).length > 2 ? [{ rangeType: 'cell', value }] : [];
  });
}

export function xlsxCellBorderKey(row: number, column: number): string {
  return `${row}_${column}`;
}

export function xlsxCellBorderLineFromFortune(
  value: unknown,
): XlsxCellBorderLine | null | undefined {
  if (value === null) return null;
  if (!isRecord(value)) return undefined;
  const style = xlsxStyleByFortuneStyle[String(value.style ?? '')];
  const color = normalizeRgbColor(value.color);
  return style && color ? { color, style } : undefined;
}

export function fortuneBorderStyle(style: XlsxCellBorderStyle): string {
  return fortuneStyleByXlsxStyle.get(style) ?? '1';
}

function spreadsheetBorderTargets(
  sheet: Sheet,
): ReadonlyMap<string, SpreadsheetBorderTarget> {
  const targets = new Map<string, SpreadsheetBorderTarget>();
  for (const [row, values] of sparseArrayEntries(sheet.data)) {
    for (const [column, cell] of sparseArrayEntries(values)) {
      if (cell) targets.set(xlsxCellBorderKey(row, column), { column, row });
    }
  }
  for (const candidate of Array.isArray(sheet.config?.borderInfo)
    ? (sheet.config.borderInfo as unknown[])
    : []) {
    if (!isRecord(candidate) || candidate.rangeType !== 'cell') continue;
    const value = candidate.value;
    if (!isRecord(value)) continue;
    const row = nonNegativeInteger(value.row_index);
    const column = nonNegativeInteger(value.col_index);
    if (row === null || column === null) continue;
    targets.set(xlsxCellBorderKey(row, column), { column, row });
  }
  return targets;
}

function applySpreadsheetCellBorderRecord(
  record: UnknownRecord,
  targets: ReadonlyMap<string, SpreadsheetBorderTarget>,
  borders: Map<string, XlsxCellBorder>,
): void {
  if (!isRecord(record.value)) return;
  const row = nonNegativeInteger(record.value.row_index);
  const column = nonNegativeInteger(record.value.col_index);
  if (row === null || column === null) return;
  const key = xlsxCellBorderKey(row, column);
  if (!targets.has(key)) return;
  const border = { ...(borders.get(key) ?? {}) };
  applyCellLine(border, 'left', record.value.l);
  applyCellLine(border, 'right', record.value.r);
  applyCellLine(border, 'top', record.value.t);
  applyCellLine(border, 'bottom', record.value.b);
  const diagonal = xlsxCellBorderLineFromFortune(record.value.s);
  if (diagonal !== undefined) {
    border.diagonal = diagonal;
    border.diagonalUp = diagonal !== null;
    border.diagonalDown = false;
  }
  borders.set(key, border);
}

function applyCellLine(
  border: XlsxCellBorder,
  side: 'bottom' | 'left' | 'right' | 'top',
  value: unknown,
): void {
  const line = xlsxCellBorderLineFromFortune(value);
  if (line !== undefined) border[side] = line;
}

function applySpreadsheetRangeBorder(
  borders: Map<string, XlsxCellBorder>,
  target: SpreadsheetBorderTarget,
  range: SpreadsheetBorderRange,
  borderType: string,
  line: XlsxCellBorderLine | null,
): void {
  const key = xlsxCellBorderKey(target.row, target.column);
  const border = { ...(borders.get(key) ?? {}) };
  const onTop = target.row === range.row[0];
  const onBottom = target.row === range.row[1];
  const onLeft = target.column === range.column[0];
  const onRight = target.column === range.column[1];
  switch (borderType) {
    case 'border-top':
      if (onTop) border.top = line;
      break;
    case 'border-bottom':
      if (onBottom) border.bottom = line;
      break;
    case 'border-left':
      if (onLeft) border.left = line;
      break;
    case 'border-right':
      if (onRight) border.right = line;
      break;
    case 'border-none':
      border.top = null;
      border.bottom = null;
      border.left = null;
      border.right = null;
      border.diagonal = null;
      border.diagonalUp = false;
      border.diagonalDown = false;
      break;
    case 'border-all':
      border.top = line;
      border.bottom = line;
      border.left = line;
      border.right = line;
      break;
    case 'border-outside':
      if (onTop) border.top = line;
      if (onBottom) border.bottom = line;
      if (onLeft) border.left = line;
      if (onRight) border.right = line;
      break;
    case 'border-inside':
      if (!onTop) border.top = line;
      if (!onBottom) border.bottom = line;
      if (!onLeft) border.left = line;
      if (!onRight) border.right = line;
      break;
    case 'border-horizontal':
      if (!onTop) border.top = line;
      if (!onBottom) border.bottom = line;
      break;
    case 'border-vertical':
      if (!onLeft) border.left = line;
      if (!onRight) border.right = line;
      break;
    case 'border-slash':
      border.diagonal = line;
      border.diagonalUp = line !== null;
      border.diagonalDown = false;
      break;
    default:
      return;
  }
  borders.set(key, border);
}

function spreadsheetBorderLine(
  record: UnknownRecord,
): XlsxCellBorderLine | null {
  const style = xlsxStyleByFortuneStyle[String(record.style ?? '')];
  const color = normalizeRgbColor(record.color);
  return style && color ? { color, style } : null;
}

function spreadsheetBorderRange(value: unknown): SpreadsheetBorderRange | null {
  if (!isRecord(value)) return null;
  const row = spreadsheetBorderAxis(value.row);
  const column = spreadsheetBorderAxis(value.column);
  return row && column ? { column, row } : null;
}

function spreadsheetBorderAxis(
  value: unknown,
): readonly [number, number] | null {
  if (!Array.isArray(value) || value.length !== 2) return null;
  const start = nonNegativeInteger(value[0]);
  const end = nonNegativeInteger(value[1]);
  return start !== null && end !== null && start <= end ? [start, end] : null;
}

function spreadsheetBorderRangeContains(
  range: SpreadsheetBorderRange,
  target: SpreadsheetBorderTarget,
): boolean {
  return (
    target.row >= range.row[0] &&
    target.row <= range.row[1] &&
    target.column >= range.column[0] &&
    target.column <= range.column[1]
  );
}

function setFortuneBorderLine(
  target: UnknownRecord,
  name: 'b' | 'l' | 'r' | 's' | 't',
  line: XlsxCellBorderLine | null | undefined,
): void {
  if (!line) return;
  target[name] = {
    color: normalizeRgbColor(line.color) ?? '#000000',
    style: fortuneBorderStyle(line.style),
  };
}

function normalizeRgbColor(value: unknown): string | null {
  if (typeof value !== 'string') return null;
  const color = value.trim().toLowerCase();
  if (/^#[0-9a-f]{6}$/.test(color)) return color;
  if (!/^#[0-9a-f]{3}$/.test(color)) return null;
  return `#${[...color.slice(1)]
    .map((character) => character.repeat(2))
    .join('')}`;
}

function nonNegativeInteger(value: unknown): number | null {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0
    ? value
    : null;
}

function isRecord(value: unknown): value is UnknownRecord {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}
