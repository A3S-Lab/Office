import type { Op } from '@fortune-sheet/core';
import {
  normalizedWorkSpreadsheetAutoFilterRange,
  workSpreadsheetAutoFilterCriteriaEntries,
  workSpreadsheetAutoFilterManuallyHiddenRows,
  workSpreadsheetAutoFilterOwnedRows,
  workSpreadsheetFilterHiddenRows,
  workSpreadsheetSheetWithReappliedAutoFilterCriteria,
} from '../work-spreadsheet-auto-filter';
import type { WorkSpreadsheetDynamicFilterContext } from '../work-spreadsheet-dynamic-filter';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetSheet,
  WorkSpreadsheetTable,
} from '../work-types';
import type { SpreadsheetCellRange } from './spreadsheet-cell-range';
import { spreadsheetCellOperationCoordinates } from './spreadsheet-operation-projection';
import { MAX_SPREADSHEET_SORT_CELLS } from './spreadsheet-sort';

export function reconcileSpreadsheetFiltersAfterFortune(
  sheets: WorkSpreadsheetContent['sheets'],
  sourceSheets: WorkSpreadsheetContent['sheets'],
  operations: readonly Op[] = [],
  context: WorkSpreadsheetDynamicFilterContext = { now: new Date() },
): WorkSpreadsheetContent['sheets'] {
  const coordinates = spreadsheetCellOperationCoordinates(operations);
  return sheets.map((sheet, index) => {
    const source = sourceSheet(sheet, sourceSheets, index);
    if (
      !source ||
      !spreadsheetSheetHasCriteria(source) ||
      !spreadsheetFilterReconciliationIsBounded(source) ||
      !spreadsheetFilterRangesWereEdited(source, operations, coordinates)
    ) {
      return sheet;
    }

    const previousTableOwnedRows = spreadsheetTableFilterOwnedRows(
      source,
      context,
    );
    const manuallyHiddenRows =
      workSpreadsheetAutoFilterManuallyHiddenRows(source);
    for (const row of previousTableOwnedRows) manuallyHiddenRows.delete(row);

    let next = sheet;
    if (
      normalizedWorkSpreadsheetAutoFilterRange(sheet.filter_select) &&
      workSpreadsheetAutoFilterCriteriaEntries(source).length
    ) {
      next = workSpreadsheetSheetWithReappliedAutoFilterCriteria(
        sheet,
        source,
        manuallyHiddenRows,
        context,
      );
    }

    const hiddenRows = new Set(manuallyHiddenRows);
    for (const row of workSpreadsheetAutoFilterOwnedRows(next)) {
      hiddenRows.add(row);
    }
    for (const row of spreadsheetTableFilterOwnedRows(next, context)) {
      hiddenRows.add(row);
    }
    const rowhidden = Object.fromEntries(
      [...hiddenRows]
        .sort(compareSpreadsheetRowKeys)
        .map((row) => [row, 0 as const]),
    );
    return {
      ...next,
      config: { ...(next.config ?? {}), rowhidden },
    };
  });
}

export function reconcileSpreadsheetFiltersAfterSort(
  sheet: WorkSpreadsheetSheet,
  source: WorkSpreadsheetSheet,
  range: SpreadsheetCellRange,
  sourceIndexes: readonly number[],
  context: WorkSpreadsheetDynamicFilterContext = { now: new Date() },
): WorkSpreadsheetSheet | null {
  if (
    !spreadsheetFilterReconciliationIsBounded(source) ||
    !spreadsheetSortSourceIndexesAreValid(range, sourceIndexes)
  ) {
    return null;
  }
  const reconciled = reconcileSpreadsheetFiltersAfterFortune(
    [sheet],
    [source],
    [],
    context,
  )[0];
  if (!reconciled) return null;
  let filter = reconciled.filter;
  const autoFilter = normalizedWorkSpreadsheetAutoFilterRange(
    source.filter_select,
  );
  if (
    autoFilter &&
    autoFilter.row[0] === range.row[0] &&
    autoFilter.row[1] === range.row[1] &&
    autoFilter.column[0] === range.column[0] &&
    autoFilter.column[1] === range.column[1]
  ) {
    filter = spreadsheetFilterWithRemappedOpaqueEntries(
      reconciled,
      source,
      range,
      sourceIndexes,
    );
  }
  const manuallyHiddenRows =
    workSpreadsheetAutoFilterManuallyHiddenRows(source);
  for (const row of spreadsheetTableFilterOwnedRows(source, context)) {
    manuallyHiddenRows.delete(row);
  }
  const withFilter = { ...reconciled, filter };
  const hiddenRows = new Set(manuallyHiddenRows);
  for (const row of workSpreadsheetAutoFilterOwnedRows(withFilter)) {
    hiddenRows.add(row);
  }
  for (const row of spreadsheetTableFilterOwnedRows(withFilter, context)) {
    hiddenRows.add(row);
  }
  const rowhidden = Object.fromEntries(
    [...hiddenRows]
      .sort(compareSpreadsheetRowKeys)
      .map((row) => [row, 0 as const]),
  );
  return {
    ...sheet,
    filter,
    config: { ...(sheet.config ?? {}), rowhidden },
  };
}

export function spreadsheetFilterReconciliationIsBounded(
  sheet: WorkSpreadsheetSheet,
): boolean {
  let cellVisits = 0;
  const autoFilter = normalizedWorkSpreadsheetAutoFilterRange(
    sheet.filter_select,
  );
  if (autoFilter) {
    const filterCount = Object.keys(sheet.filter ?? {}).length;
    cellVisits +=
      (autoFilter.row[1] - autoFilter.row[0]) *
      Math.max(
        filterCount,
        workSpreadsheetAutoFilterCriteriaEntries(sheet).length,
      );
  }
  for (const table of sheet.tables ?? []) {
    const range = spreadsheetTableFilterRange(table);
    if (!range) continue;
    cellVisits += (range.row[1] - range.row[0]) * table.filters.length;
    if (
      !Number.isSafeInteger(cellVisits) ||
      cellVisits > MAX_SPREADSHEET_SORT_CELLS
    ) {
      return false;
    }
  }
  return (
    Number.isSafeInteger(cellVisits) && cellVisits <= MAX_SPREADSHEET_SORT_CELLS
  );
}

function spreadsheetFilterWithRemappedOpaqueEntries(
  reconciled: WorkSpreadsheetSheet,
  source: WorkSpreadsheetSheet,
  range: SpreadsheetCellRange,
  sourceIndexes: readonly number[],
): WorkSpreadsheetSheet['filter'] {
  const next = { ...(reconciled.filter ?? {}) };
  const typedColumns = new Set(
    workSpreadsheetAutoFilterCriteriaEntries(source).map((entry) =>
      String(entry.column),
    ),
  );
  for (const [column, entry] of Object.entries(source.filter ?? {})) {
    if (typedColumns.has(column) || !isRecord(entry)) continue;
    next[column] = {
      ...entry,
      rowhidden: spreadsheetRemappedFilterRows(
        entry.rowhidden,
        range,
        sourceIndexes,
      ),
    };
  }
  return next;
}

function spreadsheetRemappedFilterRows(
  value: unknown,
  range: SpreadsheetCellRange,
  sourceIndexes: readonly number[],
): Record<string, 0> {
  const sourceRows = isRecord(value) ? value : {};
  const rowhidden: Record<string, 0> = {};
  const dataStart = range.row[0] + 1;
  sourceIndexes.forEach((sourceIndex, targetIndex) => {
    if (Object.hasOwn(sourceRows, String(dataStart + sourceIndex))) {
      rowhidden[String(dataStart + targetIndex)] = 0;
    }
  });
  return rowhidden;
}

function spreadsheetSortSourceIndexesAreValid(
  range: SpreadsheetCellRange,
  sourceIndexes: readonly number[],
): boolean {
  const count = range.row[1] - range.row[0];
  return (
    sourceIndexes.length === count &&
    sourceIndexes.every(
      (index) => Number.isSafeInteger(index) && index >= 0 && index < count,
    ) &&
    new Set(sourceIndexes).size === count
  );
}

function spreadsheetSheetHasCriteria(sheet: WorkSpreadsheetSheet): boolean {
  return Boolean(
    workSpreadsheetAutoFilterCriteriaEntries(sheet).length ||
      sheet.tables?.some((table) => table.filters.length),
  );
}

function spreadsheetFilterRangesWereEdited(
  sheet: WorkSpreadsheetSheet,
  operations: readonly Op[],
  coordinates: ReturnType<typeof spreadsheetCellOperationCoordinates>,
): boolean {
  if (!operations.length) return true;
  if (!sheet.id) return false;
  const sheetCoordinates = coordinates?.get(sheet.id);
  if (sheetCoordinates) {
    const ranges = spreadsheetFilterRanges(sheet);
    return [...sheetCoordinates.values()].some(({ column, row }) =>
      ranges.some(
        (range) =>
          row >= range.row[0] &&
          row <= range.row[1] &&
          column >= range.column[0] &&
          column <= range.column[1],
      ),
    );
  }
  return operations.some((operation) => operation.id === sheet.id);
}

function spreadsheetFilterRanges(sheet: WorkSpreadsheetSheet) {
  const ranges: Array<{
    column: [number, number];
    row: [number, number];
  }> = [];
  const autoFilter = normalizedWorkSpreadsheetAutoFilterRange(
    sheet.filter_select,
  );
  if (autoFilter && workSpreadsheetAutoFilterCriteriaEntries(sheet).length) {
    ranges.push(autoFilter);
  }
  for (const table of sheet.tables ?? []) {
    if (!table.filters.length) continue;
    const range = spreadsheetTableFilterRange(table);
    if (range) ranges.push(range);
  }
  return ranges;
}

function spreadsheetTableFilterOwnedRows(
  sheet: WorkSpreadsheetSheet,
  context: WorkSpreadsheetDynamicFilterContext,
): Set<string> {
  const hidden = new Set<string>();
  for (const table of sheet.tables ?? []) {
    const range = spreadsheetTableFilterRange(table);
    if (!range) continue;
    for (const filter of table.filters) {
      const column = range.column[0] + filter.column;
      const rowhidden = workSpreadsheetFilterHiddenRows(
        sheet,
        range,
        column,
        filter.criteria,
        context,
      );
      for (const row of Object.keys(rowhidden ?? {})) hidden.add(row);
    }
  }
  return hidden;
}

function spreadsheetTableFilterRange(table: WorkSpreadsheetTable) {
  if (!table.headerRow || !table.filters.length) return null;
  const rowStart = finiteIndex(table.range.row[0]);
  const rowEnd = finiteIndex(table.range.row[1]);
  const columnStart = finiteIndex(table.range.column[0]);
  const columnEnd = finiteIndex(table.range.column[1]);
  if (
    rowStart === null ||
    rowEnd === null ||
    columnStart === null ||
    columnEnd === null
  ) {
    return null;
  }
  const end = Math.max(rowStart, rowEnd) - Number(table.totalsRow);
  const start = Math.min(rowStart, rowEnd);
  if (end <= start) return null;
  return {
    row: [start, end] as [number, number],
    column: [
      Math.min(columnStart, columnEnd),
      Math.max(columnStart, columnEnd),
    ] as [number, number],
  };
}

function sourceSheet(
  sheet: WorkSpreadsheetSheet,
  sources: WorkSpreadsheetContent['sheets'],
  index: number,
): WorkSpreadsheetSheet | undefined {
  return (
    sources.find((candidate) => candidate.id === sheet.id) ?? sources[index]
  );
}

function finiteIndex(value: unknown): number | null {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0
    ? value
    : null;
}

function compareSpreadsheetRowKeys(left: string, right: string): number {
  return Number(left) - Number(right);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value && typeof value === 'object' && !Array.isArray(value));
}
