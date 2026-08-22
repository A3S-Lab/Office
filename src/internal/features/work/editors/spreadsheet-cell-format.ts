import type { Cell } from '@fortune-sheet/core';
import { cloneSparseMatrix } from '../spreadsheet-sparse';
import {
  normalizeSheetProtectionAuthority,
  type SpreadsheetCellProtectionRange,
} from '../work-spreadsheet-protection';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetSheet,
} from '../work-types';
import {
  isSpreadsheetUnderlineStyle,
  spreadsheetUnderlineCellValue,
  type SpreadsheetUnderlineStyle,
} from '../work-spreadsheet-underline';
import {
  isSpreadsheetTextOrientationId,
  spreadsheetTextOrientationCellStyle,
  spreadsheetTextOrientationFromAngle,
  spreadsheetTextOrientationFromChoice,
  type SpreadsheetTextOrientationId,
} from '../work-spreadsheet-text-orientation';
import {
  MAX_SPREADSHEET_DIAGONAL_BORDER_CELLS,
  normalizeSpreadsheetBorderFormat,
  type SpreadsheetCellBorderFormat,
} from './spreadsheet-cell-border';
import { setSpreadsheetCellBordersPerCell } from './spreadsheet-cell-border-per-cell';
import {
  normalizeSpreadsheetCellRange,
  type SpreadsheetCellRange,
  spreadsheetCellRangeArea,
  spreadsheetCellRangeContains,
  subtractSpreadsheetCellRange,
} from './spreadsheet-cell-range';
import { spreadsheetNumberFormatValue } from './spreadsheet-number-format';
import { patchSpreadsheetRichTextFontRuns } from '../work-xlsx-rich-text';
import {
  normalizeSpreadsheetCellFillFormat,
  type SpreadsheetCellFillFormat,
  withSpreadsheetCellFillFormat,
} from './spreadsheet-cell-fill-format';

export const MAX_SPREADSHEET_CELL_FORMAT_CELLS = 10_000;

export type SpreadsheetHorizontalAlignment =
  | 'general'
  | 'left'
  | 'center'
  | 'right';
export type SpreadsheetVerticalAlignment = 'top' | 'middle' | 'bottom';

export interface SpreadsheetCellFormatPatch {
  numberFormat?: string;
  horizontalAlignment?: SpreadsheetHorizontalAlignment;
  verticalAlignment?: SpreadsheetVerticalAlignment;
  wrapText?: boolean;
  rotation?: number;
  textOrientation?: SpreadsheetTextOrientationId;
  fontFamily?: string;
  fontSize?: number;
  fontColor?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: SpreadsheetUnderlineStyle;
  strike?: boolean;
  fill?: SpreadsheetCellFillFormat;
  borders?: readonly SpreadsheetCellBorderFormat[];
  locked?: boolean;
  hidden?: boolean;
}

export interface SpreadsheetCellFormatRequest {
  sheetId: string;
  range: SpreadsheetCellRange;
  patch: SpreadsheetCellFormatPatch;
}

type ProtectionCell = Cell & { hi?: number };
type SpreadsheetCellDataEntry = NonNullable<
  WorkSpreadsheetSheet['celldata']
>[number];

const patchKeys = new Set<keyof SpreadsheetCellFormatPatch>([
  'numberFormat',
  'horizontalAlignment',
  'verticalAlignment',
  'wrapText',
  'rotation',
  'textOrientation',
  'fontFamily',
  'fontSize',
  'fontColor',
  'bold',
  'italic',
  'underline',
  'strike',
  'fill',
  'borders',
  'locked',
  'hidden',
]);

export function canApplySpreadsheetCellFormat(
  content: WorkSpreadsheetContent,
  request: SpreadsheetCellFormatRequest,
): boolean {
  const range = normalizeSpreadsheetCellRange(request.range);
  const sheet = content.sheets.find(
    (candidate) => candidate.id === request.sheetId,
  );
  const keys = Object.keys(
    request.patch,
  ) as (keyof SpreadsheetCellFormatPatch)[];
  if (
    !sheet ||
    !range ||
    !keys.length ||
    keys.some((key) => !patchKeys.has(key)) ||
    spreadsheetCellRangeArea(range) > MAX_SPREADSHEET_CELL_FORMAT_CELLS
  ) {
    return false;
  }
  const patch = request.patch;
  if (
    (patch.numberFormat !== undefined &&
      !validTrimmedString(patch.numberFormat, 255)) ||
    (patch.horizontalAlignment !== undefined &&
      !['general', 'left', 'center', 'right'].includes(
        patch.horizontalAlignment,
      )) ||
    (patch.verticalAlignment !== undefined &&
      !['top', 'middle', 'bottom'].includes(patch.verticalAlignment)) ||
    (patch.wrapText !== undefined && typeof patch.wrapText !== 'boolean') ||
    (patch.rotation !== undefined &&
      (!Number.isInteger(patch.rotation) ||
        patch.rotation < -90 ||
        patch.rotation > 90)) ||
    (patch.textOrientation !== undefined &&
      !isSpreadsheetTextOrientationId(patch.textOrientation)) ||
    (patch.rotation !== undefined && patch.textOrientation !== undefined) ||
    (patch.fontFamily !== undefined &&
      !validTrimmedString(patch.fontFamily, 128)) ||
    (patch.fontSize !== undefined &&
      (!Number.isFinite(patch.fontSize) ||
        patch.fontSize < 1 ||
        patch.fontSize > 409)) ||
    (patch.fontColor !== undefined && !normalizeColor(patch.fontColor)) ||
    (patch.fill !== undefined &&
      !normalizeSpreadsheetCellFillFormat(patch.fill)) ||
    !validOptionalBoolean(patch.bold) ||
    !validOptionalBoolean(patch.italic) ||
    (patch.underline !== undefined &&
      !isSpreadsheetUnderlineStyle(patch.underline)) ||
    !validOptionalBoolean(patch.strike) ||
    !validOptionalBoolean(patch.locked) ||
    !validOptionalBoolean(patch.hidden)
  ) {
    return false;
  }
  if (patch.borders !== undefined) {
    const normalized = patch.borders.map(normalizeSpreadsheetBorderFormat);
    if (
      normalized.some(
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
      (normalized.some(
        (format) =>
          format?.target === 'diagonalDown' || format?.target === 'diagonalUp',
      ) &&
        spreadsheetCellRangeArea(range) > MAX_SPREADSHEET_DIAGONAL_BORDER_CELLS)
    ) {
      return false;
    }
  }
  return true;
}

export function applySpreadsheetCellFormat(
  content: WorkSpreadsheetContent,
  request: SpreadsheetCellFormatRequest,
): WorkSpreadsheetContent | null {
  if (!canApplySpreadsheetCellFormat(content, request)) return null;
  const range = normalizeSpreadsheetCellRange(request.range);
  const sheetIndex = content.sheets.findIndex(
    (sheet) => sheet.id === request.sheetId,
  );
  const sheet = content.sheets[sheetIndex];
  if (!range || !sheet || sheetIndex < 0) return null;

  const protectionRanges = formatProtectionRanges(sheet, range, request.patch);
  let nextSheet = formatSheetCells(sheet, range, request.patch);
  if (protectionRanges) {
    const authority = normalizeSheetProtectionAuthority(
      nextSheet.config?.authority,
    );
    authority.cellProtectionRanges = protectionRanges;
    nextSheet = {
      ...nextSheet,
      config: { ...nextSheet.config, authority },
    };
  }

  const sheets = [...content.sheets];
  sheets[sheetIndex] = nextSheet;
  let next: WorkSpreadsheetContent = { ...content, sheets };
  if (request.patch.borders !== undefined) {
    const borders = request.patch.borders.flatMap((format) => {
      const normalized = normalizeSpreadsheetBorderFormat(format);
      return normalized ? [normalized] : [];
    });
    const withBorders = setSpreadsheetCellBordersPerCell(
      next,
      request.sheetId,
      range,
      borders,
    );
    if (!withBorders) return null;
    next = withBorders;
  }
  return next;
}

export function spreadsheetCellProtectionAt(
  sheet: WorkSpreadsheetSheet,
  row: number,
  column: number,
  cell: Cell | null | undefined = sheetCellAt(sheet, row, column),
): { locked: boolean; hidden: boolean } {
  const direct = cell as ProtectionCell | null | undefined;
  let compact: SpreadsheetCellProtectionRange | undefined;
  const ranges = normalizeSheetProtectionAuthority(
    sheet.config?.authority,
  ).cellProtectionRanges;
  for (let index = ranges.length - 1; index >= 0; index -= 1) {
    const candidate = ranges[index];
    if (spreadsheetCellRangeContains(candidate.range, row, column)) {
      compact = candidate;
      break;
    }
  }
  return {
    locked:
      direct?.lo === undefined ? (compact?.locked ?? true) : direct.lo !== 0,
    hidden:
      direct?.hi === undefined ? (compact?.hidden ?? false) : direct.hi === 1,
  };
}

function formatSheetCells(
  sheet: WorkSpreadsheetSheet,
  range: SpreadsheetCellRange,
  patch: SpreadsheetCellFormatPatch,
): WorkSpreadsheetSheet {
  if (!patchTouchesCells(patch)) return sheet;
  if (sheet.data !== undefined) return formatMatrixCells(sheet, range, patch);
  return formatSparseCells(sheet, range, patch);
}

function formatMatrixCells(
  sheet: WorkSpreadsheetSheet,
  range: SpreadsheetCellRange,
  patch: SpreadsheetCellFormatPatch,
): WorkSpreadsheetSheet {
  const data = cloneSparseMatrix(sheet.data);
  let materialized = false;
  for (let row = range.row[0]; row <= range.row[1]; row += 1) {
    for (let column = range.column[0]; column <= range.column[1]; column += 1) {
      const source = sheet.data?.[row]?.[column] ?? null;
      const next = formatCell(source, patch);
      if (next) {
        const values = data[row] ?? [];
        data[row] = values;
        values[column] = next;
        materialized = true;
      } else if (source) {
        const values = data[row];
        if (values) delete values[column];
      }
    }
  }
  return {
    ...sheet,
    ...(materialized
      ? {
          row: Math.max(sheet.row ?? 0, range.row[1] + 1),
          column: Math.max(sheet.column ?? 0, range.column[1] + 1),
        }
      : {}),
    data,
  };
}

function formatSparseCells(
  sheet: WorkSpreadsheetSheet,
  range: SpreadsheetCellRange,
  patch: SpreadsheetCellFormatPatch,
): WorkSpreadsheetSheet {
  const byCoordinate = new Map<string, SpreadsheetCellDataEntry>();
  for (const entry of sheet.celldata ?? []) {
    byCoordinate.set(cellKey(entry.r, entry.c), entry);
  }
  let materialized = false;
  for (let row = range.row[0]; row <= range.row[1]; row += 1) {
    for (let column = range.column[0]; column <= range.column[1]; column += 1) {
      const key = cellKey(row, column);
      const entry = byCoordinate.get(key);
      const next = formatCell(entry?.v ?? null, patch);
      if (next) {
        byCoordinate.set(key, { ...(entry ?? { r: row, c: column }), v: next });
        materialized = true;
      } else {
        byCoordinate.delete(key);
      }
    }
  }
  const celldata = [...byCoordinate.values()].sort(
    (left, right) => left.r - right.r || left.c - right.c,
  );
  return {
    ...sheet,
    ...(materialized
      ? {
          row: Math.max(sheet.row ?? 0, range.row[1] + 1),
          column: Math.max(sheet.column ?? 0, range.column[1] + 1),
        }
      : {}),
    celldata,
  };
}

function formatCell(
  source: Cell | null | undefined,
  patch: SpreadsheetCellFormatPatch,
): Cell | null {
  let next = { ...(source ?? {}) } as ProtectionCell;
  const writable = next as unknown as Record<string, unknown>;
  if (patch.numberFormat !== undefined) {
    const format = spreadsheetNumberFormatValue(
      patch.numberFormat.trim(),
      source,
    );
    next.ct = {
      ...source?.ct,
      ...format,
      ...(source?.ct?.t === 'inlineStr' ? { t: 'inlineStr' } : {}),
    };
    delete next.m;
  }
  if (patch.horizontalAlignment !== undefined) {
    if (patch.horizontalAlignment === 'general') delete writable.ht;
    else
      writable.ht = { left: '1', center: '0', right: '2' }[
        patch.horizontalAlignment
      ];
  }
  if (patch.verticalAlignment !== undefined) {
    next.vt = { top: 1, middle: 0, bottom: 2 }[patch.verticalAlignment];
  }
  if (patch.wrapText !== undefined) next.tb = patch.wrapText ? '2' : '1';
  if (patch.rotation !== undefined || patch.textOrientation !== undefined) {
    const orientation =
      patch.textOrientation !== undefined
        ? spreadsheetTextOrientationFromChoice(patch.textOrientation)
        : spreadsheetTextOrientationFromAngle(patch.rotation);
    const style = spreadsheetTextOrientationCellStyle(orientation);
    delete next.rt;
    delete next.tr;
    if (style?.rt !== undefined) next.rt = style.rt;
    if (style?.tr !== undefined) next.tr = style.tr;
  }
  if (patch.fontFamily !== undefined) next.ff = patch.fontFamily.trim();
  if (patch.fontSize !== undefined) next.fs = patch.fontSize;
  if (patch.fontColor !== undefined)
    next.fc = normalizeColor(patch.fontColor) ?? patch.fontColor;
  if (patch.bold !== undefined) next.bl = patch.bold ? 1 : 0;
  if (patch.italic !== undefined) next.it = patch.italic ? 1 : 0;
  if (patch.underline !== undefined) {
    next.un = spreadsheetUnderlineCellValue(patch.underline);
  }
  if (patch.strike !== undefined) next.cl = patch.strike ? 1 : 0;
  if (patch.fill !== undefined) {
    const filled = withSpreadsheetCellFillFormat(next, patch.fill);
    if (filled) next = filled as ProtectionCell;
  }
  if (patch.locked !== undefined || patch.hidden !== undefined) {
    delete next.lo;
    delete next.hi;
  }
  const withRichText = patchSpreadsheetRichTextFontRuns(next, patch);
  return Object.keys(withRichText).length ? withRichText : null;
}

function formatProtectionRanges(
  sheet: WorkSpreadsheetSheet,
  range: SpreadsheetCellRange,
  patch: SpreadsheetCellFormatPatch,
): SpreadsheetCellProtectionRange[] | null {
  if (patch.locked === undefined && patch.hidden === undefined) return null;
  const authority = normalizeSheetProtectionAuthority(sheet.config?.authority);
  const preserved = authority.cellProtectionRanges.flatMap((item) =>
    subtractSpreadsheetCellRange(item.range, range).map((remaining) => ({
      ...item,
      range: remaining,
    })),
  );
  return [...preserved, ...compactTargetProtection(sheet, range, patch)];
}

function compactTargetProtection(
  sheet: WorkSpreadsheetSheet,
  range: SpreadsheetCellRange,
  patch: SpreadsheetCellFormatPatch,
): SpreadsheetCellProtectionRange[] {
  const completed: SpreadsheetCellProtectionRange[] = [];
  let active = new Map<string, SpreadsheetCellProtectionRange>();
  for (let row = range.row[0]; row <= range.row[1]; row += 1) {
    const runs: SpreadsheetCellProtectionRange[] = [];
    let current: SpreadsheetCellProtectionRange | null = null;
    for (let column = range.column[0]; column <= range.column[1]; column += 1) {
      const source = spreadsheetCellProtectionAt(sheet, row, column);
      const state = {
        locked: patch.locked ?? source.locked,
        hidden: patch.hidden ?? source.hidden,
      };
      if (
        current &&
        current.locked === state.locked &&
        current.hidden === state.hidden
      ) {
        current.range.column[1] = column;
      } else {
        if (current) runs.push(current);
        current = {
          range: { row: [row, row], column: [column, column] },
          ...state,
        };
      }
    }
    if (current) runs.push(current);

    const nextActive = new Map<string, SpreadsheetCellProtectionRange>();
    for (const run of runs) {
      const signature = `${run.locked}:${run.hidden}:${run.range.column[0]}:${run.range.column[1]}`;
      const previous = active.get(signature);
      if (previous && previous.range.row[1] === row - 1) {
        previous.range.row[1] = row;
        nextActive.set(signature, previous);
      } else {
        nextActive.set(signature, run);
      }
    }
    for (const [signature, item] of active) {
      if (!nextActive.has(signature)) completed.push(item);
    }
    active = nextActive;
  }
  completed.push(...active.values());
  return completed;
}

function patchTouchesCells(patch: SpreadsheetCellFormatPatch): boolean {
  return Object.keys(patch).some((key) => key !== 'borders');
}

function sheetCellAt(
  sheet: WorkSpreadsheetSheet,
  row: number,
  column: number,
): Cell | null | undefined {
  return (
    sheet.data?.[row]?.[column] ??
    sheet.celldata?.find((entry) => entry.r === row && entry.c === column)?.v
  );
}

function validTrimmedString(value: unknown, maximum: number): value is string {
  return (
    typeof value === 'string' &&
    value.trim().length > 0 &&
    value.trim().length <= maximum
  );
}

function validOptionalBoolean(value: unknown): boolean {
  return value === undefined || typeof value === 'boolean';
}

function normalizeColor(value: string): string | null {
  const color = value.trim().toLowerCase();
  if (/^#[0-9a-f]{6}$/.test(color)) return color;
  if (!/^#[0-9a-f]{3}$/.test(color)) return null;
  return `#${[...color.slice(1)]
    .map((character) => character.repeat(2))
    .join('')}`;
}

function cellKey(row: number, column: number): string {
  return `${row}:${column}`;
}
