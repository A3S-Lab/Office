import type { Cell, Selection, Sheet } from '@fortune-sheet/core';
import {
  type XlsxNativeFillCellKey,
  xlsxNativeFillCellKeys,
} from '../work-xlsx-native-fill';
import {
  parseSpreadsheetCellRange,
  type SpreadsheetCellRange,
  spreadsheetCellRangeContains,
  spreadsheetCellRangesIntersect,
  subtractSpreadsheetCellRange,
} from './spreadsheet-cell-range';

export type SpreadsheetCellClearMode =
  | 'all'
  | 'formats'
  | 'contents'
  | 'comments'
  | 'hyperlinks';

type SpreadsheetCellClearRange = SpreadsheetCellRange;

type ClearableCell = Cell & { hi?: number } & Partial<
    Record<XlsxNativeFillCellKey, unknown>
  >;
type UnknownRecord = Record<string, unknown>;

const spreadsheetCellFormatProperties = [
  'bg',
  'bl',
  'it',
  'ff',
  'fs',
  'fc',
  'ht',
  'vt',
  'tb',
  'cl',
  'un',
  'tr',
  'rt',
  'ct',
  'lo',
  'hi',
  ...xlsxNativeFillCellKeys,
] as const satisfies readonly (keyof ClearableCell)[];

export function clearSpreadsheetSheetSelection(
  sheet: Sheet,
  selection: Pick<Selection, 'row' | 'column'>,
  mode: SpreadsheetCellClearMode,
): Sheet {
  const range = normalizeClearRange(selection);
  const data = sheet.data?.map((row, rowIndex) => {
    if (rowIndex < range.row[0] || rowIndex > range.row[1]) return row;
    return row.map((cell, columnIndex) =>
      clearSpreadsheetCellAt(cell, rowIndex, columnIndex, range, mode),
    );
  });
  const celldata = data
    ? undefined
    : sheet.celldata?.flatMap((entry) => {
        const cell = clearSpreadsheetCellAt(
          entry.v,
          entry.r,
          entry.c,
          range,
          mode,
        );
        return cell ? [{ ...entry, v: cell }] : [];
      });
  let next: Sheet = {
    ...sheet,
    ...(data ? { data } : {}),
    ...(celldata ? { celldata } : {}),
  };
  if (data && next.celldata) delete next.celldata;

  if (mode === 'formats' || mode === 'all') {
    next = clearSpreadsheetSheetFormats(next, range);
  }
  if (mode === 'hyperlinks' || mode === 'all') {
    next = clearSpreadsheetSheetHyperlinks(next, range);
  }
  return next;
}

function clearSpreadsheetCellAt(
  cell: Cell | null,
  row: number,
  column: number,
  range: SpreadsheetCellClearRange,
  mode: SpreadsheetCellClearMode,
): Cell | null {
  if (!cell || !clearRangeContains(range, row, column)) return cell;
  if (mode === 'all') {
    return cell.mc ? { mc: { ...cell.mc } } : null;
  }

  const next: ClearableCell = { ...cell };
  if (mode === 'contents') clearSpreadsheetCellContents(next);
  if (mode === 'formats') clearSpreadsheetCellFormats(next);
  if (mode === 'comments') delete next.ps;
  if (mode === 'hyperlinks') delete next.hl;
  removeUndefinedCellProperties(next);
  return Object.keys(next).length ? next : null;
}

function clearSpreadsheetCellContents(cell: ClearableCell): void {
  delete cell.v;
  delete cell.m;
  delete cell.f;
  delete cell.spl;
  delete cell.qp;
  if (cell.ct && 's' in cell.ct) {
    const cellType = { ...cell.ct };
    delete cellType.s;
    cell.ct = Object.keys(cellType).length ? cellType : undefined;
  }
}

function clearSpreadsheetCellFormats(cell: ClearableCell): void {
  const inlineText = spreadsheetInlineCellText(cell);
  for (const property of spreadsheetCellFormatProperties) {
    delete cell[property];
  }
  if (cell.v !== undefined) {
    cell.m = spreadsheetGeneralDisplayValue(cell.v);
  } else if (!cell.f && inlineText !== null) {
    cell.v = inlineText;
    cell.m = inlineText;
  }
}

function spreadsheetInlineCellText(cell: Cell): string | null {
  const segments = cell.ct?.s;
  if (!Array.isArray(segments)) return null;
  return segments
    .map((segment) => {
      if (!segment || typeof segment !== 'object') return '';
      const value = (segment as UnknownRecord).v;
      return value === null || value === undefined ? '' : String(value);
    })
    .join('');
}

function spreadsheetGeneralDisplayValue(value: Cell['v']): string {
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  return value === undefined ? '' : String(value);
}

function removeUndefinedCellProperties(cell: ClearableCell): void {
  const record = cell as UnknownRecord;
  for (const key of Object.keys(record)) {
    if (record[key] === undefined) delete record[key];
  }
}

function clearSpreadsheetSheetFormats(
  sheet: Sheet,
  range: SpreadsheetCellClearRange,
): Sheet {
  const config = sheet.config ? { ...sheet.config } : undefined;
  if (config?.borderInfo) {
    const borderInfo = clearSpreadsheetBorderFormats(
      config.borderInfo as unknown[],
      range,
    );
    if (borderInfo.length) {
      config.borderInfo = borderInfo as NonNullable<
        Sheet['config']
      >['borderInfo'];
    } else {
      delete config.borderInfo;
    }
  }

  const conditionalFormats = clearSpreadsheetConditionalFormats(
    sheet.luckysheet_conditionformat_save,
    range,
  );
  const alternateFormats = clearSpreadsheetAlternateFormats(
    sheet.luckysheet_alternateformat_save,
    range,
  );
  const next: Sheet = {
    ...sheet,
    ...(config ? { config } : {}),
    ...(conditionalFormats.length
      ? { luckysheet_conditionformat_save: conditionalFormats }
      : {}),
    ...(alternateFormats.length
      ? { luckysheet_alternateformat_save: alternateFormats }
      : {}),
  };
  if (!conditionalFormats.length) delete next.luckysheet_conditionformat_save;
  if (!alternateFormats.length) delete next.luckysheet_alternateformat_save;
  return next;
}

function clearSpreadsheetBorderFormats(
  formats: unknown[],
  range: SpreadsheetCellClearRange,
): unknown[] {
  return formats.flatMap((format) => {
    if (!isRecord(format)) return [format];
    if (format.rangeType === 'cell') {
      const value = isRecord(format.value) ? format.value : null;
      const row = finiteIndex(value?.row_index);
      const column = finiteIndex(value?.col_index);
      return row !== null &&
        column !== null &&
        spreadsheetCellRangeContains(range, row, column)
        ? []
        : [format];
    }
    if (format.rangeType !== 'range' || !Array.isArray(format.range)) {
      return [format];
    }
    const remaining = format.range.flatMap((candidate) => {
      const parsed = parseSpreadsheetCellRange(candidate);
      if (!parsed) return [candidate];
      if (format.borderType === 'border-slash') {
        return spreadsheetCellRangesIntersect(parsed, range) ? [] : [parsed];
      }
      return subtractSpreadsheetCellRange(parsed, range);
    });
    return remaining.length ? [{ ...format, range: remaining }] : [];
  });
}

function clearSpreadsheetConditionalFormats(
  formats: Sheet['luckysheet_conditionformat_save'],
  range: SpreadsheetCellClearRange,
): unknown[] {
  if (!Array.isArray(formats)) return [];
  return formats.flatMap((format) => {
    if (!isRecord(format) || !Array.isArray(format.cellrange)) return [format];
    const remaining = format.cellrange.flatMap((candidate) => {
      const parsed = parseSpreadsheetCellRange(candidate);
      return parsed ? subtractSpreadsheetCellRange(parsed, range) : [candidate];
    });
    return remaining.length ? [{ ...format, cellrange: remaining }] : [];
  });
}

function clearSpreadsheetAlternateFormats(
  formats: Sheet['luckysheet_alternateformat_save'],
  range: SpreadsheetCellClearRange,
): unknown[] {
  if (!Array.isArray(formats)) return [];
  return formats.flatMap((format) => {
    if (!isRecord(format)) return [format];
    const source = parseSpreadsheetCellRange(format.cellrange);
    if (!source) return [format];
    return subtractSpreadsheetCellRange(source, range).map((cellrange) => ({
      ...format,
      cellrange,
    }));
  });
}

function clearSpreadsheetSheetHyperlinks(
  sheet: Sheet,
  range: SpreadsheetCellClearRange,
): Sheet {
  if (!sheet.hyperlink) return sheet;
  const hyperlink = Object.fromEntries(
    Object.entries(sheet.hyperlink).filter(([key]) => {
      const position = spreadsheetHyperlinkPosition(key);
      return (
        !position ||
        !spreadsheetCellRangeContains(range, position.row, position.column)
      );
    }),
  );
  const next: Sheet = { ...sheet, hyperlink };
  if (!Object.keys(hyperlink).length) delete next.hyperlink;
  return next;
}

function spreadsheetHyperlinkPosition(
  value: string,
): { row: number; column: number } | null {
  const match = /^(\d+)_(\d+)$/.exec(value);
  if (!match) return null;
  return { row: Number(match[1]), column: Number(match[2]) };
}

function clearRangeContains(
  range: SpreadsheetCellClearRange,
  row: number,
  column: number,
): boolean {
  return (
    row >= range.row[0] &&
    row <= range.row[1] &&
    column >= range.column[0] &&
    column <= range.column[1]
  );
}

function normalizeClearRange(
  range: Pick<Selection, 'row' | 'column'>,
): SpreadsheetCellClearRange {
  return {
    row: normalizeClearAxis(range.row),
    column: normalizeClearAxis(range.column),
  };
}

function normalizeClearAxis(value: number[]): [number, number] {
  const first = finiteIndex(value[0]) ?? 0;
  const second = finiteIndex(value[1]) ?? first;
  return [Math.min(first, second), Math.max(first, second)];
}

function finiteIndex(value: unknown): number | null {
  return typeof value === 'number' && Number.isFinite(value) && value >= 0
    ? Math.trunc(value)
    : null;
}

function isRecord(value: unknown): value is UnknownRecord {
  return Boolean(value && typeof value === 'object' && !Array.isArray(value));
}
