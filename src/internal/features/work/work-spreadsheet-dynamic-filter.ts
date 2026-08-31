import type { Cell } from '@fortune-sheet/core';
import type {
  WorkSpreadsheetDateSystem,
  WorkSpreadsheetDynamicFilter,
} from './work-types';

const MILLISECONDS_PER_DAY = 86_400_000;
const EXCEL_1900_CIVIL_EPOCH_UTC = Date.UTC(1899, 11, 31);
const EXCEL_1904_EPOCH_UTC = Date.UTC(1904, 0, 1);
const SERIAL_INTEGER_EPSILON = 1e-7;

const CALENDAR_MONTH_FILTERS = [
  'month-1',
  'month-2',
  'month-3',
  'month-4',
  'month-5',
  'month-6',
  'month-7',
  'month-8',
  'month-9',
  'month-10',
  'month-11',
  'month-12',
] as const satisfies readonly WorkSpreadsheetDynamicFilter[];

const CALENDAR_QUARTER_FILTERS = [
  'quarter-1',
  'quarter-2',
  'quarter-3',
  'quarter-4',
] as const satisfies readonly WorkSpreadsheetDynamicFilter[];

export interface WorkSpreadsheetDynamicFilterContext {
  dateSystem?: WorkSpreadsheetDateSystem;
  now: Date;
}

export function workSpreadsheetDynamicFilterMatcher(
  kind: WorkSpreadsheetDynamicFilter,
  cells: Iterable<Cell | null>,
  context: WorkSpreadsheetDynamicFilterContext,
): ((cell: Cell | null) => boolean) | null {
  if (kind === 'above-average' || kind === 'below-average') {
    return averageMatcher(kind, cells);
  }

  const month = CALENDAR_MONTH_FILTERS.indexOf(
    kind as (typeof CALENDAR_MONTH_FILTERS)[number],
  );
  if (month >= 0) {
    return (cell) =>
      spreadsheetDateParts(cell, context.dateSystem)?.month === month;
  }

  const quarter = CALENDAR_QUARTER_FILTERS.indexOf(
    kind as (typeof CALENDAR_QUARTER_FILTERS)[number],
  );
  if (quarter >= 0) {
    return (cell) =>
      spreadsheetDateParts(cell, context.dateSystem)?.quarter === quarter;
  }

  const bounds = relativeDateBounds(kind, context.now);
  if (!bounds) return null;
  return (cell) => {
    const day = spreadsheetDateCivilDay(cell, context.dateSystem);
    return day !== null && day >= bounds.start && day < bounds.end;
  };
}

export function workSpreadsheetCellIsDate(
  cell: Cell | null,
  dateSystem: WorkSpreadsheetDateSystem = '1900',
): boolean {
  return spreadsheetDateCivilDay(cell, dateSystem) !== null;
}

function averageMatcher(
  kind: 'above-average' | 'below-average',
  cells: Iterable<Cell | null>,
): (cell: Cell | null) => boolean {
  const average = finiteArithmeticMean(cells);
  if (average === null) return () => false;
  return kind === 'above-average'
    ? (cell) => {
        const value = spreadsheetNumericCellValue(cell);
        return value !== null && value > average;
      }
    : (cell) => {
        const value = spreadsheetNumericCellValue(cell);
        return value !== null && value < average;
      };
}

function finiteArithmeticMean(cells: Iterable<Cell | null>): number | null {
  let count = 0;
  let scale = 0;
  let scaledTotal = 0;
  for (const cell of cells) {
    const value = spreadsheetNumericCellValue(cell);
    if (value === null) continue;
    count += 1;
    const magnitude = Math.abs(value);
    if (magnitude > scale) {
      scaledTotal = scaledTotal * (scale / magnitude) + value / magnitude;
      scale = magnitude;
    } else if (scale > 0) {
      scaledTotal += value / scale;
    }
  }
  return count ? (scaledTotal / count) * scale : null;
}

function relativeDateBounds(
  kind: WorkSpreadsheetDynamicFilter,
  now: Date,
): { start: number; end: number } | null {
  const timestamp = now.getTime();
  const year = now.getFullYear();
  if (
    !Number.isFinite(timestamp) ||
    !Number.isSafeInteger(year) ||
    year < 1900 ||
    year > 9999
  ) {
    return null;
  }
  const month = now.getMonth();
  const today = civilDayFromDateParts(year, month, now.getDate());

  if (kind === 'yesterday') return dayBounds(today - 1);
  if (kind === 'today') return dayBounds(today);
  if (kind === 'tomorrow') return dayBounds(today + 1);

  const weekStart = today - now.getDay();
  if (kind === 'last-week') return rangeBounds(weekStart - 7, weekStart);
  if (kind === 'this-week') return rangeBounds(weekStart, weekStart + 7);
  if (kind === 'next-week') return rangeBounds(weekStart + 7, weekStart + 14);

  const monthStart = civilDayFromDateParts(year, month, 1);
  if (kind === 'last-month') {
    return rangeBounds(civilDayFromDateParts(year, month - 1, 1), monthStart);
  }
  if (kind === 'this-month') {
    return rangeBounds(monthStart, civilDayFromDateParts(year, month + 1, 1));
  }
  if (kind === 'next-month') {
    return rangeBounds(
      civilDayFromDateParts(year, month + 1, 1),
      civilDayFromDateParts(year, month + 2, 1),
    );
  }

  const quarterStartMonth = Math.floor(month / 3) * 3;
  const quarterStart = civilDayFromDateParts(year, quarterStartMonth, 1);
  if (kind === 'last-quarter') {
    return rangeBounds(
      civilDayFromDateParts(year, quarterStartMonth - 3, 1),
      quarterStart,
    );
  }
  if (kind === 'this-quarter') {
    return rangeBounds(
      quarterStart,
      civilDayFromDateParts(year, quarterStartMonth + 3, 1),
    );
  }
  if (kind === 'next-quarter') {
    return rangeBounds(
      civilDayFromDateParts(year, quarterStartMonth + 3, 1),
      civilDayFromDateParts(year, quarterStartMonth + 6, 1),
    );
  }

  const yearStart = civilDayFromDateParts(year, 0, 1);
  if (kind === 'last-year') {
    return rangeBounds(civilDayFromDateParts(year - 1, 0, 1), yearStart);
  }
  if (kind === 'this-year') {
    return rangeBounds(yearStart, civilDayFromDateParts(year + 1, 0, 1));
  }
  if (kind === 'next-year') {
    return rangeBounds(
      civilDayFromDateParts(year + 1, 0, 1),
      civilDayFromDateParts(year + 2, 0, 1),
    );
  }
  if (kind === 'year-to-date') return rangeBounds(yearStart, today + 1);
  return null;
}

function spreadsheetDateParts(
  cell: Cell | null,
  dateSystem: WorkSpreadsheetDateSystem = '1900',
): { month: number; quarter: number } | null {
  const day = spreadsheetDateCivilDay(cell, dateSystem);
  if (day === null) return null;
  const date = new Date(day * MILLISECONDS_PER_DAY);
  const timestamp = date.getTime();
  const year = date.getUTCFullYear();
  if (!Number.isFinite(timestamp) || year > 9999) return null;
  const month = date.getUTCMonth();
  return { month, quarter: Math.floor(month / 3) };
}

function spreadsheetDateCivilDay(
  cell: Cell | null,
  dateSystem: WorkSpreadsheetDateSystem = '1900',
): number | null {
  const value = spreadsheetCellRawValue(cell);
  if (value instanceof Date) {
    if (!Number.isFinite(value.getTime())) return null;
    const year = value.getFullYear();
    const minimumYear = dateSystem === '1904' ? 1904 : 1900;
    return year >= minimumYear && year <= 9999
      ? civilDayFromDateParts(year, value.getMonth(), value.getDate())
      : null;
  }
  if (
    typeof value !== 'number' ||
    !Number.isFinite(value) ||
    cell?.ct?.t !== 'd'
  ) {
    return null;
  }
  const serial = value;
  const rounded = Math.round(serial);
  const normalized =
    Math.abs(serial - rounded) <= SERIAL_INTEGER_EPSILON ? rounded : serial;
  const day = Math.floor(normalized);
  if (dateSystem === '1904') {
    return day >= 0
      ? (EXCEL_1904_EPOCH_UTC + day * MILLISECONDS_PER_DAY) /
          MILLISECONDS_PER_DAY
      : null;
  }
  if (day < 1 || day === 60) return null;
  const civilOffset = day > 60 ? day - 1 : day;
  return (
    (EXCEL_1900_CIVIL_EPOCH_UTC + civilOffset * MILLISECONDS_PER_DAY) /
    MILLISECONDS_PER_DAY
  );
}

function spreadsheetNumericCellValue(cell: Cell | null): number | null {
  const value = spreadsheetCellRawValue(cell);
  if (value instanceof Date || cell?.ct?.t === 'd') return null;
  return typeof value === 'number' && Number.isFinite(value) ? value : null;
}

function spreadsheetCellRawValue(cell: Cell | null): unknown {
  const candidate = cell as { m?: unknown; v?: unknown } | null;
  return candidate?.v ?? candidate?.m;
}

function civilDayFromDateParts(
  year: number,
  month: number,
  day: number,
): number {
  return Date.UTC(year, month, day) / MILLISECONDS_PER_DAY;
}

function dayBounds(day: number): { start: number; end: number } {
  return { start: day, end: day + 1 };
}

function rangeBounds(
  start: number,
  end: number,
): { start: number; end: number } {
  return { start, end };
}
