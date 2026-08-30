import type { Selection } from '@fortune-sheet/core';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetSheet,
} from '../work-types';
import {
  spreadsheetCurrentRegion,
  spreadsheetRowHasContent,
  spreadsheetSheetCellReader,
} from './spreadsheet-current-region';
import { finiteSpreadsheetSelection } from './spreadsheet-editor-support';

export interface SpreadsheetAutoFilterRange {
  row: [number, number];
  column: [number, number];
}

export function spreadsheetAutoFilterRange(
  sheet: WorkSpreadsheetSheet,
  selection: Selection,
): SpreadsheetAutoFilterRange | null {
  const range = spreadsheetSelectionOrCurrentRegion(sheet, selection);
  if (!range || !validSpreadsheetAutoFilterRange(sheet, range)) return null;
  return range;
}

export function spreadsheetSelectionOrCurrentRegion(
  sheet: WorkSpreadsheetSheet,
  selection: Selection,
): SpreadsheetAutoFilterRange | null {
  if (sheet.isPivotTable || sheet.pivotTables?.length) return null;
  const normalized = normalizeAutoFilterSelection(selection);
  return normalized.row[0] !== normalized.row[1]
    ? normalized
    : spreadsheetCurrentRegion(sheet, normalized);
}

export function spreadsheetAutoFilterHeaderColumn(
  sheet: WorkSpreadsheetSheet | undefined,
  selection: Selection,
): number | null {
  const range = normalizedFilterSelection(sheet?.filter_select);
  if (!sheet || !range) return null;
  const row = finiteIndex(selection.row_focus, selection.row[0]);
  const column = finiteIndex(selection.column_focus, selection.column[0]);
  if (
    row !== range.row[0] ||
    column < range.column[0] ||
    column > range.column[1]
  ) {
    return null;
  }
  return column;
}

export function toggleSpreadsheetAutoFilter(
  content: WorkSpreadsheetContent,
  sheetId: string,
  selection: Selection,
): WorkSpreadsheetContent | null {
  const sheetIndex = content.sheets.findIndex((sheet) => sheet.id === sheetId);
  const sheet = content.sheets[sheetIndex];
  if (!sheet) return null;
  const nextSheet = sheet.filter_select
    ? spreadsheetSheetWithoutAutoFilter(sheet)
    : spreadsheetSheetWithAutoFilter(sheet, selection);
  if (!nextSheet) return null;
  const sheets = [...content.sheets];
  sheets[sheetIndex] = {
    ...nextSheet,
    luckysheet_select_save: [finiteSpreadsheetSelection(selection)],
  };
  return { ...content, sheets };
}

function spreadsheetSheetWithAutoFilter(
  sheet: WorkSpreadsheetSheet,
  selection: Selection,
): WorkSpreadsheetSheet | null {
  const range = spreadsheetAutoFilterRange(sheet, selection);
  if (!range) return null;
  return {
    ...sheet,
    filter: {},
    filter_select: {
      row: [...range.row],
      column: [...range.column],
    },
  };
}

function spreadsheetSheetWithoutAutoFilter(
  sheet: WorkSpreadsheetSheet,
): WorkSpreadsheetSheet {
  const hiddenByFilter = spreadsheetFilterHiddenRows(sheet.filter);
  const rowhidden = { ...sheet.config?.rowhidden };
  for (const row of hiddenByFilter) delete rowhidden[row];
  const config = sheet.config
    ? {
        ...sheet.config,
        rowhidden,
      }
    : undefined;
  const { filter: _filter, filter_select: _selection, ...remaining } = sheet;
  return { ...remaining, config };
}

function spreadsheetFilterHiddenRows(
  filter: WorkSpreadsheetSheet['filter'],
): Set<string> {
  const rows = new Set<string>();
  for (const value of Object.values(filter ?? {})) {
    if (!value || typeof value !== 'object') continue;
    const rowhidden = (value as { rowhidden?: unknown }).rowhidden;
    if (!rowhidden || typeof rowhidden !== 'object') continue;
    for (const row of Object.keys(rowhidden)) rows.add(row);
  }
  return rows;
}

function validSpreadsheetAutoFilterRange(
  sheet: WorkSpreadsheetSheet,
  range: SpreadsheetAutoFilterRange,
): boolean {
  if (range.row[1] <= range.row[0]) return false;
  if (spreadsheetAutoFilterIntersectsTable(sheet, range)) return false;
  const cells = spreadsheetSheetCellReader(sheet);
  if (
    !spreadsheetRowHasContent(
      cells,
      range.row[0],
      range.column[0],
      range.column[1],
    )
  ) {
    return false;
  }
  let dataFound = false;
  for (let row = range.row[0] + 1; row <= range.row[1]; row += 1) {
    if (
      spreadsheetRowHasContent(cells, row, range.column[0], range.column[1])
    ) {
      dataFound = true;
      break;
    }
  }
  return dataFound && !spreadsheetRangeIntersectsMerge(sheet, range);
}

function spreadsheetAutoFilterIntersectsTable(
  sheet: WorkSpreadsheetSheet,
  range: SpreadsheetAutoFilterRange,
): boolean {
  return (sheet.tables ?? []).some(
    (table) =>
      table.range.row[0] <= range.row[1] &&
      table.range.row[1] >= range.row[0] &&
      table.range.column[0] <= range.column[1] &&
      table.range.column[1] >= range.column[0],
  );
}

function spreadsheetRangeIntersectsMerge(
  sheet: WorkSpreadsheetSheet,
  range: SpreadsheetAutoFilterRange,
): boolean {
  return Object.values(sheet.config?.merge ?? {}).some((merge) => {
    const endRow = merge.r + Math.max(merge.rs, 1) - 1;
    const endColumn = merge.c + Math.max(merge.cs, 1) - 1;
    return (
      merge.r <= range.row[1] &&
      endRow >= range.row[0] &&
      merge.c <= range.column[1] &&
      endColumn >= range.column[0]
    );
  });
}

function normalizeAutoFilterSelection(
  selection: Selection,
): SpreadsheetAutoFilterRange {
  const startRow = finiteIndex(selection.row[0], 0);
  const endRow = finiteIndex(selection.row[1], startRow);
  const startColumn = finiteIndex(selection.column[0], 0);
  const endColumn = finiteIndex(selection.column[1], startColumn);
  return {
    row: [Math.min(startRow, endRow), Math.max(startRow, endRow)],
    column: [
      Math.min(startColumn, endColumn),
      Math.max(startColumn, endColumn),
    ],
  };
}

function normalizedFilterSelection(
  selection: WorkSpreadsheetSheet['filter_select'],
): SpreadsheetAutoFilterRange | null {
  if (!selection?.row?.length || !selection.column?.length) return null;
  return normalizeAutoFilterSelection(selection as Selection);
}

function finiteIndex(value: unknown, fallback: number): number {
  return typeof value === 'number' && Number.isFinite(value)
    ? Math.max(0, Math.floor(value))
    : fallback;
}
