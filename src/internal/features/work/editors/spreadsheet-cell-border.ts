import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetSheet,
} from '../work-types';
import {
  spreadsheetCellValueWithDiagonalBorder,
  spreadsheetDiagonalBorderFromCellValue,
  type WorkSpreadsheetDiagonalBorder,
} from '../work-spreadsheet-diagonal-borders';
import {
  normalizeSpreadsheetCellRange,
  parseSpreadsheetCellRange,
  type SpreadsheetCellRange,
  type SpreadsheetCellRangeInput,
  spreadsheetCellRangeArea,
  spreadsheetCellRangeContains,
  spreadsheetCellRangesIntersect,
  subtractSpreadsheetCellRange,
} from './spreadsheet-cell-range';

export const spreadsheetCellBorderTargets = [
  'top',
  'bottom',
  'left',
  'right',
  'none',
  'all',
  'outside',
  'inside',
  'horizontal',
  'vertical',
  'diagonalDown',
  'diagonalUp',
] as const;

export type SpreadsheetCellBorderTarget =
  (typeof spreadsheetCellBorderTargets)[number];

export const spreadsheetCellBorderStyles = [
  'thin',
  'dotted',
  'dashed',
  'dash-dot',
  'dash-dot-dot',
  'medium',
  'medium-dashed',
  'medium-dash-dot',
  'medium-dash-dot-dot',
  'thick',
] as const;

export type SpreadsheetCellBorderStyle =
  (typeof spreadsheetCellBorderStyles)[number];

export interface SpreadsheetCellBorderFormat {
  target: SpreadsheetCellBorderTarget;
  color: string;
  style: SpreadsheetCellBorderStyle;
}

export interface SpreadsheetResolvedCellBorderLine {
  color: string;
  style: string;
}

export interface SpreadsheetResolvedCellBorders {
  bottom?: SpreadsheetResolvedCellBorderLine;
  diagonalDown?: SpreadsheetResolvedCellBorderLine;
  diagonalUp?: SpreadsheetResolvedCellBorderLine;
  left?: SpreadsheetResolvedCellBorderLine;
  right?: SpreadsheetResolvedCellBorderLine;
  top?: SpreadsheetResolvedCellBorderLine;
}

export interface SpreadsheetRenderableDiagonalBorders {
  down?: SpreadsheetResolvedCellBorderLine;
  up?: SpreadsheetResolvedCellBorderLine;
}

export const MAX_SPREADSHEET_DIAGONAL_BORDER_CELLS = 4_096;

type SpreadsheetNativeBorderType =
  | 'border-top'
  | 'border-bottom'
  | 'border-left'
  | 'border-right'
  | 'border-none'
  | 'border-all'
  | 'border-outside'
  | 'border-inside'
  | 'border-horizontal'
  | 'border-vertical'
  | 'border-slash';

type UnknownRecord = Record<string, unknown>;

interface SpreadsheetDiagonalBorderDirections {
  down: boolean;
  up: boolean;
}

const spreadsheetNativeBorderTypes = new Set<SpreadsheetNativeBorderType>([
  'border-top',
  'border-bottom',
  'border-left',
  'border-right',
  'border-none',
  'border-all',
  'border-outside',
  'border-inside',
  'border-horizontal',
  'border-vertical',
  'border-slash',
]);

type SpreadsheetRangeBorderTarget = Exclude<
  SpreadsheetCellBorderTarget,
  'diagonalDown' | 'diagonalUp'
>;

const spreadsheetNativeBorderTypeByTarget: Record<
  SpreadsheetRangeBorderTarget,
  SpreadsheetNativeBorderType
> = {
  top: 'border-top',
  bottom: 'border-bottom',
  left: 'border-left',
  right: 'border-right',
  none: 'border-none',
  all: 'border-all',
  outside: 'border-outside',
  inside: 'border-inside',
  horizontal: 'border-horizontal',
  vertical: 'border-vertical',
};

const spreadsheetNativeBorderStyleByName: Record<
  SpreadsheetCellBorderStyle,
  string
> = {
  thin: '1',
  dotted: '3',
  dashed: '4',
  'dash-dot': '5',
  'dash-dot-dot': '6',
  medium: '8',
  'medium-dashed': '9',
  'medium-dash-dot': '10',
  'medium-dash-dot-dot': '11',
  thick: '13',
};

export function spreadsheetNativeBorderStyle(
  style: SpreadsheetCellBorderStyle,
): string {
  return spreadsheetNativeBorderStyleByName[style];
}

export function canSetSpreadsheetCellBorders(
  content: WorkSpreadsheetContent,
  sheetId: string,
  range: SpreadsheetCellRangeInput,
  format: SpreadsheetCellBorderFormat,
): boolean {
  const normalizedRange = normalizeSpreadsheetCellRange(range);
  return Boolean(
    content.sheets.some((sheet) => sheet.id === sheetId) &&
      normalizedRange &&
      normalizeSpreadsheetBorderFormat(format) &&
      (!isSpreadsheetDiagonalBorderTarget(format.target) ||
        spreadsheetCellRangeArea(normalizedRange) <=
          MAX_SPREADSHEET_DIAGONAL_BORDER_CELLS) &&
      spreadsheetBorderInfoIsWritable(
        content.sheets.find((sheet) => sheet.id === sheetId),
      ),
  );
}

export function setSpreadsheetCellBorders(
  content: WorkSpreadsheetContent,
  sheetId: string,
  range: SpreadsheetCellRangeInput,
  format: SpreadsheetCellBorderFormat,
): WorkSpreadsheetContent | null {
  const sheetIndex = content.sheets.findIndex((sheet) => sheet.id === sheetId);
  const sheet = content.sheets[sheetIndex];
  const normalizedRange = normalizeSpreadsheetCellRange(range);
  const normalizedFormat = normalizeSpreadsheetBorderFormat(format);
  if (
    !sheet ||
    !normalizedRange ||
    !normalizedFormat ||
    !spreadsheetBorderInfoIsWritable(sheet) ||
    (isSpreadsheetDiagonalBorderTarget(normalizedFormat.target) &&
      spreadsheetCellRangeArea(normalizedRange) >
        MAX_SPREADSHEET_DIAGONAL_BORDER_CELLS)
  ) {
    return null;
  }

  const source = Array.isArray(sheet.config?.borderInfo)
    ? (sheet.config.borderInfo as unknown[])
    : [];
  if (isSpreadsheetDiagonalBorderTarget(normalizedFormat.target)) {
    const borderInfo = [
      ...compactSpreadsheetBorderInfo(source, normalizedRange, 'border-slash'),
      ...spreadsheetDiagonalBorderRecords(
        sheet,
        normalizedRange,
        normalizedFormat.target,
        normalizedFormat.color,
        spreadsheetNativeBorderStyleByName[normalizedFormat.style],
      ),
    ];
    return contentWithSpreadsheetBorderInfo(
      content,
      sheetIndex,
      sheet,
      borderInfo,
    );
  }

  const borderType =
    spreadsheetNativeBorderTypeByTarget[normalizedFormat.target];
  const compacted = compactSpreadsheetBorderInfo(
    source,
    normalizedRange,
    borderType,
  );
  const borderInfo = [
    ...compacted,
    ...spreadsheetNativeBorderRecords(
      normalizedRange,
      borderType,
      normalizedFormat.color,
      spreadsheetNativeBorderStyleByName[normalizedFormat.style],
    ),
  ];
  return contentWithSpreadsheetBorderInfo(
    content,
    sheetIndex,
    sheet,
    borderInfo,
  );
}

export function spreadsheetCellBordersAt(
  sheet: WorkSpreadsheetSheet | undefined,
  row: number,
  column: number,
): SpreadsheetResolvedCellBorders {
  const borders: SpreadsheetResolvedCellBorders = {};
  const source = Array.isArray(sheet?.config?.borderInfo)
    ? (sheet.config.borderInfo as unknown[])
    : [];
  for (const candidate of source) {
    if (!isRecord(candidate)) continue;
    if (candidate.rangeType === 'cell') {
      if (!isRecord(candidate.value)) continue;
      if (
        finiteSpreadsheetBorderIndex(candidate.value.row_index) !== row ||
        finiteSpreadsheetBorderIndex(candidate.value.col_index) !== column
      ) {
        continue;
      }
      applyResolvedCellBorderLine(borders, 'left', candidate.value.l);
      applyResolvedCellBorderLine(borders, 'right', candidate.value.r);
      applyResolvedCellBorderLine(borders, 'top', candidate.value.t);
      applyResolvedCellBorderLine(borders, 'bottom', candidate.value.b);
      applyResolvedCellDiagonalBorders(borders, candidate.value);
      continue;
    }
    if (
      candidate.rangeType !== 'range' ||
      !isSpreadsheetNativeBorderType(candidate.borderType) ||
      !Array.isArray(candidate.range)
    ) {
      continue;
    }
    const line = resolvedRangeBorderLine(candidate);
    for (const rangeCandidate of candidate.range) {
      const range = parseSpreadsheetCellRange(rangeCandidate);
      if (!range || !spreadsheetCellRangeContains(range, row, column)) continue;
      applyResolvedRangeBorder(
        borders,
        range,
        row,
        column,
        candidate.borderType,
        line,
      );
    }
  }
  return borders;
}

export function spreadsheetRenderableDiagonalBorders(
  sheet: WorkSpreadsheetSheet,
): ReadonlyMap<string, SpreadsheetRenderableDiagonalBorders> {
  const borders = new Map<
    string,
    {
      border: SpreadsheetRenderableDiagonalBorders;
      column: number;
      row: number;
    }
  >();
  const source = Array.isArray(sheet.config?.borderInfo)
    ? (sheet.config.borderInfo as unknown[])
    : [];
  for (const candidate of source) {
    if (!isRecord(candidate)) continue;
    if (candidate.rangeType === 'cell') {
      if (!isRecord(candidate.value)) continue;
      const row = finiteSpreadsheetBorderIndex(candidate.value.row_index);
      const column = finiteSpreadsheetBorderIndex(candidate.value.col_index);
      if (row === null || column === null) continue;
      const key = `${row}_${column}`;
      const diagonal = spreadsheetDiagonalBorderFromCellValue(candidate.value);
      if (diagonal === undefined) continue;
      if (!diagonal) {
        borders.delete(key);
        continue;
      }
      const line = resolvedCellBorderLine(diagonal.line);
      if (!line) {
        borders.delete(key);
        continue;
      }
      borders.set(key, {
        border: {
          ...(diagonal.down ? { down: line } : {}),
          ...(diagonal.up ? { up: line } : {}),
        },
        column,
        row,
      });
      continue;
    }
    if (
      candidate.rangeType !== 'range' ||
      (candidate.borderType !== 'border-none' &&
        candidate.borderType !== 'border-slash') ||
      !Array.isArray(candidate.range)
    ) {
      continue;
    }
    const line = resolvedRangeBorderLine(candidate);
    for (const rangeCandidate of candidate.range) {
      const range = parseSpreadsheetCellRange(rangeCandidate);
      if (!range) continue;
      for (const [key, entry] of borders) {
        if (!spreadsheetCellRangeContains(range, entry.row, entry.column)) {
          continue;
        }
        if (candidate.borderType === 'border-none' || !entry.border.up) {
          borders.delete(key);
          continue;
        }
        borders.set(key, {
          ...entry,
          border: { up: line ?? entry.border.up },
        });
      }
    }
  }
  return new Map(
    [...borders].map(([key, entry]) => [key, entry.border] as const),
  );
}

export function normalizeSpreadsheetBorderFormat(
  format: SpreadsheetCellBorderFormat,
): SpreadsheetCellBorderFormat | null {
  if (
    !spreadsheetCellBorderTargets.includes(format.target) ||
    !spreadsheetCellBorderStyles.includes(format.style)
  ) {
    return null;
  }
  const color = normalizeSpreadsheetBorderColor(format.color);
  return color ? { ...format, color } : null;
}

function normalizeSpreadsheetBorderColor(value: string): string | null {
  const trimmed = value.trim().toLowerCase();
  if (/^#[0-9a-f]{6}$/.test(trimmed)) return trimmed;
  if (!/^#[0-9a-f]{3}$/.test(trimmed)) return null;
  return `#${[...trimmed.slice(1)].map((character) => character.repeat(2)).join('')}`;
}

function spreadsheetBorderInfoIsWritable(
  sheet: WorkSpreadsheetSheet | undefined,
): boolean {
  return Boolean(
    sheet &&
      (sheet.config?.borderInfo === undefined ||
        Array.isArray(sheet.config.borderInfo)),
  );
}

function compactSpreadsheetBorderInfo(
  formats: unknown[],
  range: SpreadsheetCellRange,
  appliedType: SpreadsheetNativeBorderType,
): unknown[] {
  const supersedesAll =
    appliedType === 'border-all' || appliedType === 'border-none';
  return formats.flatMap((format) => {
    if (!isRecord(format)) return [format];
    if (format.rangeType === 'cell') {
      if (!isRecord(format.value)) return [format];
      const row = finiteSpreadsheetBorderIndex(format.value.row_index);
      const column = finiteSpreadsheetBorderIndex(format.value.col_index);
      const selected =
        row !== null &&
        column !== null &&
        spreadsheetCellRangeContains(range, row, column);
      if (!selected) return [format];
      if (supersedesAll) return [];
      if (appliedType !== 'border-slash') return [format];
      const value = spreadsheetCellValueWithDiagonalBorder(format.value, null);
      return Object.keys(value).some(
        (key) => key !== 'row_index' && key !== 'col_index',
      )
        ? [{ ...format, value }]
        : [];
    }
    if (
      format.rangeType !== 'range' ||
      !isSpreadsheetNativeBorderType(format.borderType) ||
      !Array.isArray(format.range) ||
      (!supersedesAll && format.borderType !== appliedType)
    ) {
      return [format];
    }
    const remaining = format.range.flatMap((candidate) => {
      const parsed = parseSpreadsheetCellRange(candidate);
      if (!parsed) return [candidate];
      if (format.borderType === 'border-slash') {
        if (!spreadsheetCellRangesIntersect(parsed, range)) return [candidate];
        return spreadsheetCellRangeArea(parsed) === 1 ? [] : [candidate];
      }
      return subtractSpreadsheetCellRange(parsed, range);
    });
    return remaining.length ? [{ ...format, range: remaining }] : [];
  });
}

function spreadsheetNativeBorderRecords(
  range: SpreadsheetCellRange,
  borderType: SpreadsheetNativeBorderType,
  color: string,
  style: string,
): UnknownRecord[] {
  return [
    {
      rangeType: 'range',
      borderType,
      color,
      style,
      range: [range],
    },
  ];
}

function spreadsheetDiagonalBorderRecords(
  sheet: WorkSpreadsheetSheet,
  range: SpreadsheetCellRange,
  target: 'diagonalDown' | 'diagonalUp',
  color: string,
  style: string,
): UnknownRecord[] {
  const records: UnknownRecord[] = [];
  const existing = spreadsheetDiagonalBorderDirectionsInRange(sheet, range);
  for (let row = range.row[0]; row <= range.row[1]; row += 1) {
    for (let column = range.column[0]; column <= range.column[1]; column += 1) {
      const current = existing.get(`${row}_${column}`);
      const border: WorkSpreadsheetDiagonalBorder = {
        down: target === 'diagonalDown' || Boolean(current?.down),
        line: { color, style },
        up: target === 'diagonalUp' || Boolean(current?.up),
      };
      records.push({
        rangeType: 'cell',
        value: spreadsheetCellValueWithDiagonalBorder(
          { col_index: column, row_index: row },
          border,
        ),
      });
    }
  }
  return records;
}

function spreadsheetDiagonalBorderDirectionsInRange(
  sheet: WorkSpreadsheetSheet,
  selection: SpreadsheetCellRange,
): ReadonlyMap<string, SpreadsheetDiagonalBorderDirections> {
  const directions = new Map<string, SpreadsheetDiagonalBorderDirections>();
  const source = Array.isArray(sheet.config?.borderInfo)
    ? (sheet.config.borderInfo as unknown[])
    : [];
  for (const candidate of source) {
    if (!isRecord(candidate)) continue;
    if (candidate.rangeType === 'cell') {
      if (!isRecord(candidate.value)) continue;
      const row = finiteSpreadsheetBorderIndex(candidate.value.row_index);
      const column = finiteSpreadsheetBorderIndex(candidate.value.col_index);
      if (
        row === null ||
        column === null ||
        !spreadsheetCellRangeContains(selection, row, column)
      ) {
        continue;
      }
      const diagonal = spreadsheetDiagonalBorderFromCellValue(candidate.value);
      if (diagonal === undefined) continue;
      const line = diagonal && resolvedCellBorderLine(diagonal.line);
      setSpreadsheetDiagonalDirections(
        directions,
        row,
        column,
        Boolean(line && diagonal?.down),
        Boolean(line && diagonal?.up),
      );
      continue;
    }
    if (
      candidate.rangeType !== 'range' ||
      (candidate.borderType !== 'border-none' &&
        candidate.borderType !== 'border-slash') ||
      !Array.isArray(candidate.range)
    ) {
      continue;
    }
    const line = resolvedRangeBorderLine(candidate);
    for (const rangeCandidate of candidate.range) {
      const range = parseSpreadsheetCellRange(rangeCandidate);
      if (!range || !spreadsheetCellRangesIntersect(range, selection)) {
        continue;
      }
      const startRow = Math.max(range.row[0], selection.row[0]);
      const endRow = Math.min(range.row[1], selection.row[1]);
      const startColumn = Math.max(range.column[0], selection.column[0]);
      const endColumn = Math.min(range.column[1], selection.column[1]);
      for (let row = startRow; row <= endRow; row += 1) {
        for (let column = startColumn; column <= endColumn; column += 1) {
          if (candidate.borderType === 'border-none') {
            directions.delete(`${row}_${column}`);
            continue;
          }
          const current = directions.get(`${row}_${column}`);
          setSpreadsheetDiagonalDirections(
            directions,
            row,
            column,
            Boolean(line),
            Boolean(current?.up),
          );
        }
      }
    }
  }
  return directions;
}

function setSpreadsheetDiagonalDirections(
  directions: Map<string, SpreadsheetDiagonalBorderDirections>,
  row: number,
  column: number,
  down: boolean,
  up: boolean,
): void {
  const key = `${row}_${column}`;
  if (!down && !up) {
    directions.delete(key);
    return;
  }
  directions.set(key, { down, up });
}

function contentWithSpreadsheetBorderInfo(
  content: WorkSpreadsheetContent,
  sheetIndex: number,
  sheet: WorkSpreadsheetSheet,
  borderInfo: unknown[],
): WorkSpreadsheetContent {
  const nextSheet: WorkSpreadsheetSheet = {
    ...sheet,
    config: {
      ...sheet.config,
      borderInfo: borderInfo as NonNullable<
        WorkSpreadsheetSheet['config']
      >['borderInfo'],
    },
  };
  const sheets = [...content.sheets];
  sheets[sheetIndex] = nextSheet;
  return { ...content, sheets };
}

function applyResolvedCellBorderLine(
  borders: SpreadsheetResolvedCellBorders,
  side: keyof SpreadsheetResolvedCellBorders,
  value: unknown,
): void {
  if (value === null) {
    delete borders[side];
    return;
  }
  if (!isRecord(value)) return;
  const color =
    typeof value.color === 'string'
      ? normalizeSpreadsheetBorderColor(value.color)
      : null;
  const style = typeof value.style === 'string' ? value.style : null;
  if (color && style) borders[side] = { color, style };
}

function applyResolvedCellDiagonalBorders(
  borders: SpreadsheetResolvedCellBorders,
  value: UnknownRecord,
): void {
  const diagonal = spreadsheetDiagonalBorderFromCellValue(value);
  if (diagonal === undefined) return;
  delete borders.diagonalDown;
  delete borders.diagonalUp;
  if (!diagonal) return;
  const line = resolvedCellBorderLine(diagonal.line);
  if (!line) return;
  if (diagonal.down) borders.diagonalDown = line;
  if (diagonal.up) borders.diagonalUp = line;
}

function resolvedCellBorderLine(
  value: unknown,
): SpreadsheetResolvedCellBorderLine | null {
  if (!isRecord(value)) return null;
  const color =
    typeof value.color === 'string'
      ? normalizeSpreadsheetBorderColor(value.color)
      : null;
  const style = typeof value.style === 'string' ? value.style : null;
  return color && style ? { color, style } : null;
}

function resolvedRangeBorderLine(
  value: UnknownRecord,
): SpreadsheetResolvedCellBorderLine | null {
  const color =
    typeof value.color === 'string'
      ? normalizeSpreadsheetBorderColor(value.color)
      : null;
  const style = typeof value.style === 'string' ? value.style : null;
  return color && style ? { color, style } : null;
}

function applyResolvedRangeBorder(
  borders: SpreadsheetResolvedCellBorders,
  range: SpreadsheetCellRange,
  row: number,
  column: number,
  type: SpreadsheetNativeBorderType,
  line: SpreadsheetResolvedCellBorderLine | null,
): void {
  const onTop = row === range.row[0];
  const onBottom = row === range.row[1];
  const onLeft = column === range.column[0];
  const onRight = column === range.column[1];
  const set = (side: keyof SpreadsheetResolvedCellBorders) => {
    if (line) borders[side] = line;
    else delete borders[side];
  };
  switch (type) {
    case 'border-top':
      if (onTop) set('top');
      break;
    case 'border-bottom':
      if (onBottom) set('bottom');
      break;
    case 'border-left':
      if (onLeft) set('left');
      break;
    case 'border-right':
      if (onRight) set('right');
      break;
    case 'border-none':
      delete borders.top;
      delete borders.bottom;
      delete borders.left;
      delete borders.right;
      delete borders.diagonalDown;
      delete borders.diagonalUp;
      break;
    case 'border-all':
      set('top');
      set('bottom');
      set('left');
      set('right');
      break;
    case 'border-outside':
      if (onTop) set('top');
      if (onBottom) set('bottom');
      if (onLeft) set('left');
      if (onRight) set('right');
      break;
    case 'border-inside':
      if (!onTop) set('top');
      if (!onBottom) set('bottom');
      if (!onLeft) set('left');
      if (!onRight) set('right');
      break;
    case 'border-horizontal':
      if (!onTop) set('top');
      if (!onBottom) set('bottom');
      break;
    case 'border-vertical':
      if (!onLeft) set('left');
      if (!onRight) set('right');
      break;
    case 'border-slash':
      if (line) {
        borders.diagonalDown = line;
        if (borders.diagonalUp) borders.diagonalUp = line;
      } else {
        delete borders.diagonalDown;
      }
      break;
  }
}

function isSpreadsheetDiagonalBorderTarget(
  target: SpreadsheetCellBorderTarget,
): target is 'diagonalDown' | 'diagonalUp' {
  return target === 'diagonalDown' || target === 'diagonalUp';
}

function isSpreadsheetNativeBorderType(
  value: unknown,
): value is SpreadsheetNativeBorderType {
  return (
    typeof value === 'string' &&
    spreadsheetNativeBorderTypes.has(value as SpreadsheetNativeBorderType)
  );
}

function finiteSpreadsheetBorderIndex(value: unknown): number | null {
  return typeof value === 'number' && Number.isFinite(value) && value >= 0
    ? Math.trunc(value)
    : null;
}

function isRecord(value: unknown): value is UnknownRecord {
  return Boolean(value && typeof value === 'object' && !Array.isArray(value));
}
