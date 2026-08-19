import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetSheet,
} from '../work-types';
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
  'diagonal',
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

const spreadsheetNativeBorderTypeByTarget: Record<
  SpreadsheetCellBorderTarget,
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
  diagonal: 'border-slash',
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
      (format.target !== 'diagonal' ||
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
    (normalizedFormat.target === 'diagonal' &&
      spreadsheetCellRangeArea(normalizedRange) >
        MAX_SPREADSHEET_DIAGONAL_BORDER_CELLS)
  ) {
    return null;
  }

  const borderType =
    spreadsheetNativeBorderTypeByTarget[normalizedFormat.target];
  const source = Array.isArray(sheet.config?.borderInfo)
    ? (sheet.config.borderInfo as unknown[])
    : [];
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

function normalizeSpreadsheetBorderFormat(
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
      if (!supersedesAll || !isRecord(format.value)) return [format];
      const row = finiteSpreadsheetBorderIndex(format.value.row_index);
      const column = finiteSpreadsheetBorderIndex(format.value.col_index);
      return row !== null &&
        column !== null &&
        spreadsheetCellRangeContains(range, row, column)
        ? []
        : [format];
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
        return spreadsheetCellRangesIntersect(parsed, range) ? [] : [candidate];
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
  if (borderType !== 'border-slash') {
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

  const records: UnknownRecord[] = [];
  for (let row = range.row[0]; row <= range.row[1]; row += 1) {
    for (let column = range.column[0]; column <= range.column[1]; column += 1) {
      records.push({
        rangeType: 'range',
        borderType,
        color,
        style,
        range: [
          {
            row: [row, row],
            column: [column, column],
            row_focus: row,
            column_focus: column,
          },
        ],
      });
    }
  }
  return records;
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
