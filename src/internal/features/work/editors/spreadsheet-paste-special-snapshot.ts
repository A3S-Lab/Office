import type { Cell } from '@fortune-sheet/core';
import { normalizeSheetProtectionAuthority } from '../work-spreadsheet-protection';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetDataValidationItem,
  WorkSpreadsheetSheet,
} from '../work-types';
import { spreadsheetCellBordersAt } from './spreadsheet-cell-border';
import {
  normalizeSpreadsheetCellRange,
  type SpreadsheetCellRange,
  type SpreadsheetCellRangeInput,
  spreadsheetCellRangeArea,
  spreadsheetCellRangeContains,
  spreadsheetCellRangesIntersect,
} from './spreadsheet-cell-range';
import {
  MAX_SPREADSHEET_PASTE_SPECIAL_CELLS,
  spreadsheetPasteContentOptions,
  type SpreadsheetClipboardCell,
  type SpreadsheetClipboardMerge,
  type SpreadsheetClipboardSnapshot,
  type SpreadsheetPasteContent,
} from './spreadsheet-paste-special-types';

type UnknownRecord = Record<string, unknown>;

export function captureSpreadsheetClipboardSnapshot(
  content: WorkSpreadsheetContent,
  sheetId: string,
  selection: SpreadsheetCellRangeInput,
  plainText: string,
): SpreadsheetClipboardSnapshot | null {
  const range = normalizeSpreadsheetCellRange(selection);
  const sheet = content.sheets.find((candidate) => candidate.id === sheetId);
  if (
    !range ||
    !sheet ||
    spreadsheetCellRangeArea(range) > MAX_SPREADSHEET_PASTE_SPECIAL_CELLS
  ) {
    return null;
  }

  const merges = captureSpreadsheetMerges(sheet, range);
  if (!merges) return null;
  const cellAt = createSpreadsheetCellReader(sheet);
  const validationAt = createSpreadsheetValidationReader(sheet);
  const protectionAt = createSpreadsheetExplicitProtectionReader(sheet);
  const cells: SpreadsheetClipboardCell[][] = [];
  let containsUnsupportedFormulaState = false;

  for (let row = range.row[0]; row <= range.row[1]; row += 1) {
    const values: SpreadsheetClipboardCell[] = [];
    for (let column = range.column[0]; column <= range.column[1]; column += 1) {
      const source = cellAt(row, column);
      if (
        source?.spl !== undefined ||
        source?.qp !== undefined ||
        source?.f?.includes('[')
      ) {
        containsUnsupportedFormulaState = true;
      }
      values.push({
        cell: source ? structuredClone(source) : null,
        borders: structuredClone(spreadsheetCellBordersAt(sheet, row, column)),
        validation: cloneOptional(validationAt(row, column)),
        protection: protectionAt(row, column),
        hyperlink: cloneOptional(sheet.hyperlink?.[`${row}_${column}`]),
      });
    }
    cells.push(values);
  }

  const columnWidths = Array.from(
    { length: range.column[1] - range.column[0] + 1 },
    (_value, offset) =>
      positiveNumber(sheet.config?.columnlen?.[range.column[0] + offset]) ??
      positiveNumber(sheet.defaultColWidth) ??
      96,
  );
  return {
    version: 1,
    kind: 'rich',
    plainText: normalizeSpreadsheetClipboardText(plainText),
    sourceSheetId: sheetId,
    sourceRange: range,
    rowCount: range.row[1] - range.row[0] + 1,
    columnCount: range.column[1] - range.column[0] + 1,
    cells,
    columnWidths,
    merges,
    containsUnsupportedFormulaState,
  };
}

export function createSpreadsheetTextClipboardSnapshot(
  plainText: string,
): SpreadsheetClipboardSnapshot | null {
  const normalized = normalizeSpreadsheetClipboardText(plainText);
  if (!normalized) return null;
  const rows = normalized.split('\n').map((row) => row.split('\t'));
  const columnCount = Math.max(...rows.map((row) => row.length));
  const rowCount = rows.length;
  if (
    !columnCount ||
    rowCount * columnCount > MAX_SPREADSHEET_PASTE_SPECIAL_CELLS
  ) {
    return null;
  }
  const cells = rows.map((row) =>
    Array.from({ length: columnCount }, (_value, column) => {
      const value = row[column] ?? '';
      const cell = value
        ? value.startsWith('=')
          ? { v: value, m: value, f: value }
          : { v: value, m: value }
        : null;
      return { cell, borders: {} } satisfies SpreadsheetClipboardCell;
    }),
  );
  const containsUnsupportedFormulaState = cells.some((row) =>
    row.some(({ cell }) => Boolean(cell?.f?.includes('['))),
  );
  return {
    version: 1,
    kind: 'text',
    plainText: normalized,
    sourceRange: {
      row: [0, rowCount - 1],
      column: [0, columnCount - 1],
    },
    rowCount,
    columnCount,
    cells,
    merges: [],
    containsUnsupportedFormulaState,
  };
}

export function spreadsheetPasteSpecialModeAvailable(
  snapshot: SpreadsheetClipboardSnapshot,
  content: SpreadsheetPasteContent,
): boolean {
  const option = spreadsheetPasteContentOptions.find(
    (candidate) => candidate.value === content,
  );
  return Boolean(option && (!option.richOnly || snapshot.kind === 'rich'));
}

export function validSpreadsheetClipboardSnapshot(
  snapshot: SpreadsheetClipboardSnapshot,
): boolean {
  const range = normalizeSpreadsheetCellRange(snapshot.sourceRange);
  return Boolean(
    snapshot.version === 1 &&
      (snapshot.kind === 'rich' || snapshot.kind === 'text') &&
      typeof snapshot.plainText === 'string' &&
      range &&
      snapshot.rowCount > 0 &&
      snapshot.columnCount > 0 &&
      range.row[1] - range.row[0] + 1 === snapshot.rowCount &&
      range.column[1] - range.column[0] + 1 === snapshot.columnCount &&
      snapshot.rowCount * snapshot.columnCount <=
        MAX_SPREADSHEET_PASTE_SPECIAL_CELLS &&
      snapshot.cells.length === snapshot.rowCount &&
      snapshot.cells.every((row) => row.length === snapshot.columnCount) &&
      Array.isArray(snapshot.merges),
  );
}

export function createSpreadsheetCellReader(
  sheet: WorkSpreadsheetSheet,
): (row: number, column: number) => Cell | null | undefined {
  let sparse: Map<string, Cell | null> | null = null;
  return (row, column) => {
    if (sheet.data !== undefined) return sheet.data[row]?.[column];
    sparse ??= new Map(
      (sheet.celldata ?? []).map((entry) => [`${entry.r}:${entry.c}`, entry.v]),
    );
    return sparse.get(`${row}:${column}`);
  };
}

export function normalizeSpreadsheetClipboardText(value: string): string {
  return value.replace(/\r\n?/g, '\n').replace(/\n$/, '');
}

function createSpreadsheetValidationReader(
  sheet: WorkSpreadsheetSheet,
): (
  row: number,
  column: number,
) => WorkSpreadsheetDataValidationItem | undefined {
  return (row, column) => {
    const direct = sheet.dataVerification?.[`${row}_${column}`];
    if (isRecord(direct)) {
      return direct as unknown as WorkSpreadsheetDataValidationItem;
    }
    for (
      let index = (sheet.dataValidationRanges?.length ?? 0) - 1;
      index >= 0;
      index -= 1
    ) {
      const candidate = sheet.dataValidationRanges?.[index];
      if (
        candidate?.ranges.some((range) =>
          spreadsheetCellRangeContains(range, row, column),
        )
      ) {
        return candidate.item;
      }
    }
    return undefined;
  };
}

function createSpreadsheetExplicitProtectionReader(
  sheet: WorkSpreadsheetSheet,
): (
  row: number,
  column: number,
) => { locked: boolean; hidden: boolean } | undefined {
  const authority = normalizeSheetProtectionAuthority(sheet.config?.authority);
  const cellAt = createSpreadsheetCellReader(sheet);
  return (row, column) => {
    const cell = cellAt(row, column) as (Cell & { hi?: number }) | undefined;
    if (cell?.lo !== undefined || cell?.hi !== undefined) {
      return { locked: cell.lo !== 0, hidden: cell.hi === 1 };
    }
    for (
      let index = authority.cellProtectionRanges.length - 1;
      index >= 0;
      index -= 1
    ) {
      const candidate = authority.cellProtectionRanges[index];
      if (spreadsheetCellRangeContains(candidate.range, row, column)) {
        return { locked: candidate.locked, hidden: candidate.hidden };
      }
    }
    return undefined;
  };
}

function captureSpreadsheetMerges(
  sheet: WorkSpreadsheetSheet,
  range: SpreadsheetCellRange,
): SpreadsheetClipboardMerge[] | null {
  const merges: SpreadsheetClipboardMerge[] = [];
  for (const value of Object.values(sheet.config?.merge ?? {})) {
    if (!isNativeMerge(value)) continue;
    const mergeRange: SpreadsheetCellRange = {
      row: [value.r, value.r + value.rs - 1],
      column: [value.c, value.c + value.cs - 1],
    };
    if (!spreadsheetCellRangesIntersect(range, mergeRange)) continue;
    if (
      !spreadsheetCellRangeContains(
        range,
        mergeRange.row[0],
        mergeRange.column[0],
      ) ||
      !spreadsheetCellRangeContains(
        range,
        mergeRange.row[1],
        mergeRange.column[1],
      )
    ) {
      return null;
    }
    merges.push({
      row: value.r - range.row[0],
      column: value.c - range.column[0],
      rowSpan: value.rs,
      columnSpan: value.cs,
    });
  }
  return merges.sort(
    (left, right) => left.row - right.row || left.column - right.column,
  );
}

function isNativeMerge(
  value: unknown,
): value is { r: number; c: number; rs: number; cs: number } {
  if (!isRecord(value)) return false;
  return (
    [value.r, value.c, value.rs, value.cs].every(Number.isSafeInteger) &&
    Number(value.r) >= 0 &&
    Number(value.c) >= 0 &&
    Number(value.rs) > 0 &&
    Number(value.cs) > 0
  );
}

function positiveNumber(value: unknown): number | null {
  return typeof value === 'number' && Number.isFinite(value) && value > 0
    ? value
    : null;
}

function cloneOptional<T>(value: T | undefined): T | undefined {
  return value === undefined ? undefined : structuredClone(value);
}

function isRecord(value: unknown): value is UnknownRecord {
  return Boolean(value && typeof value === 'object' && !Array.isArray(value));
}
