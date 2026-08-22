import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetSheet,
} from '../work-types';
import { spreadsheetCellValueWithDiagonalBorder } from '../work-spreadsheet-diagonal-borders';
import {
  MAX_SPREADSHEET_DIAGONAL_BORDER_CELLS,
  normalizeSpreadsheetBorderFormat,
  setSpreadsheetCellBorders,
  type SpreadsheetCellBorderFormat,
  spreadsheetNativeBorderStyle,
} from './spreadsheet-cell-border';
import {
  normalizeSpreadsheetCellRange,
  type SpreadsheetCellRangeInput,
  spreadsheetCellRangeArea,
} from './spreadsheet-cell-range';

export const MAX_SPREADSHEET_PER_CELL_BORDER_CELLS = 10_000;

type UnknownRecord = Record<string, unknown>;

export function setSpreadsheetCellBordersPerCell(
  content: WorkSpreadsheetContent,
  sheetId: string,
  range: SpreadsheetCellRangeInput,
  formats: readonly SpreadsheetCellBorderFormat[],
): WorkSpreadsheetContent | null {
  const normalizedRange = normalizeSpreadsheetCellRange(range);
  const normalizedFormats = formats.map(normalizeSpreadsheetBorderFormat);
  const area = normalizedRange
    ? spreadsheetCellRangeArea(normalizedRange)
    : Number.POSITIVE_INFINITY;
  if (
    !normalizedRange ||
    area > MAX_SPREADSHEET_PER_CELL_BORDER_CELLS ||
    normalizedFormats.some(
      (format) =>
        !format ||
        ![
          'top',
          'bottom',
          'left',
          'right',
          'all',
          'diagonalDown',
          'diagonalUp',
        ].includes(format.target),
    ) ||
    (area > MAX_SPREADSHEET_DIAGONAL_BORDER_CELLS &&
      normalizedFormats.some(
        (format) =>
          format?.target === 'diagonalDown' || format?.target === 'diagonalUp',
      ))
  ) {
    return null;
  }

  const next = setSpreadsheetCellBorders(content, sheetId, normalizedRange, {
    target: 'none',
    color: '#000000',
    style: 'thin',
  });
  if (!next || !normalizedFormats.length) return next;
  const sheetIndex = next.sheets.findIndex((sheet) => sheet.id === sheetId);
  const sheet = next.sheets[sheetIndex];
  if (!sheet) return null;

  const records: UnknownRecord[] = [];
  for (
    let row = normalizedRange.row[0];
    row <= normalizedRange.row[1];
    row += 1
  ) {
    for (
      let column = normalizedRange.column[0];
      column <= normalizedRange.column[1];
      column += 1
    ) {
      const value: UnknownRecord = { row_index: row, col_index: column };
      let diagonalDown = false;
      let diagonalUp = false;
      let diagonalLine: { color: string; style: string } | null = null;
      for (const format of normalizedFormats) {
        if (!format) continue;
        const line = {
          color: format.color,
          style: spreadsheetNativeBorderStyle(format.style),
        };
        if (format.target === 'all' || format.target === 'left') value.l = line;
        if (format.target === 'all' || format.target === 'right')
          value.r = line;
        if (format.target === 'all' || format.target === 'top') value.t = line;
        if (format.target === 'all' || format.target === 'bottom')
          value.b = line;
        if (format.target === 'diagonalDown') {
          diagonalDown = true;
          diagonalLine = line;
        }
        if (format.target === 'diagonalUp') {
          diagonalUp = true;
          diagonalLine = line;
        }
      }
      records.push({
        rangeType: 'cell',
        value:
          diagonalLine && (diagonalDown || diagonalUp)
            ? spreadsheetCellValueWithDiagonalBorder(value, {
                down: diagonalDown,
                line: diagonalLine,
                up: diagonalUp,
              })
            : value,
      });
    }
  }

  const borderInfo = [
    ...(Array.isArray(sheet.config?.borderInfo)
      ? (sheet.config.borderInfo as unknown[])
      : []),
    ...records,
  ];
  const sheets = [...next.sheets];
  sheets[sheetIndex] = {
    ...sheet,
    config: {
      ...sheet.config,
      borderInfo: borderInfo as NonNullable<
        WorkSpreadsheetSheet['config']
      >['borderInfo'],
    },
  };
  return { ...next, sheets };
}
