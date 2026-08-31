import type { Cell } from '@fortune-sheet/core';
import { parseSpreadsheetCellRanges } from '../work-spreadsheet-ranges';
import type { WorkSpreadsheetSheet } from '../work-types';
import {
  createOfficeEditorExtension,
  type OfficeEditorExtension,
} from './office-editor-extension';
import { canMutateSpreadsheetCellRange } from './spreadsheet-cell-mutation-guard';
import {
  normalizeSpreadsheetCellRange,
  type SpreadsheetCellRange,
  spreadsheetCellRangesEqual,
  spreadsheetCellRangesIntersect,
} from './spreadsheet-cell-range';
import type {
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import {
  canEditSpreadsheetSelection,
  spreadsheetLiveCommandRange,
  spreadsheetLiveCommandSelection,
} from './spreadsheet-command-selection';
import {
  createSpreadsheetSortRangePlan,
  MAX_SPREADSHEET_SORT_CELLS,
  type SpreadsheetSortDirection,
  type SpreadsheetSortIntent,
  type SpreadsheetSortNormalizedRequest,
  type SpreadsheetSortOpenRequest,
  type SpreadsheetSortOwnedRange,
  type SpreadsheetSortRequest,
  spreadsheetSortOwnedRangeForExactRange,
  spreadsheetSortRowsFromSheet,
  spreadsheetSortRowsMatchRange,
  validateSpreadsheetSortRequest,
} from './spreadsheet-sort';
import { createSpreadsheetSortAppearanceRows } from './spreadsheet-sort-appearance';
import {
  reconcileSpreadsheetFiltersAfterSort,
  spreadsheetFilterReconciliationIsBounded,
} from './spreadsheet-filter-reconciliation';
import { sortSpreadsheetMatrix } from './spreadsheet-sort-matrix';

export function createSpreadsheetSortExtension(): OfficeEditorExtension<
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands
> {
  return createOfficeEditorExtension<
    SpreadsheetCommandContext,
    SpreadsheetEditorCommands
  >({
    name: 'spreadsheetSort',
    addCommands: () => ({
      applyCustomSort: {
        canExecute: canApplySpreadsheetSort,
        execute: applySpreadsheetSort,
      },
      openCustomSort: {
        canExecute: canOpenSpreadsheetSort,
        execute: openSpreadsheetSort,
      },
      sortSelectedCells: {
        canExecute: canQuickSortSpreadsheet,
        execute: quickSortSpreadsheet,
      },
    }),
  });
}

function canOpenSpreadsheetSort(context: SpreadsheetCommandContext): boolean {
  return Boolean(
    context.sort.canOpen &&
      createSpreadsheetSortOpenRequest(context, { type: 'custom' }),
  );
}

function openSpreadsheetSort(context: SpreadsheetCommandContext): boolean {
  if (!context.sort.canOpen) return false;
  const request = createSpreadsheetSortOpenRequest(context, {
    type: 'custom',
  });
  return Boolean(request && context.sort.open(request));
}

function canQuickSortSpreadsheet(
  context: SpreadsheetCommandContext,
  direction: SpreadsheetSortDirection,
): boolean {
  if (direction !== 'ascending' && direction !== 'descending') return false;
  const request = createSpreadsheetSortOpenRequest(context, {
    type: 'quick',
    direction,
  });
  return Boolean(request && (!request.expanded || context.sort.canOpen));
}

function quickSortSpreadsheet(
  context: SpreadsheetCommandContext,
  direction: SpreadsheetSortDirection,
): boolean {
  const request = createSpreadsheetSortOpenRequest(context, {
    type: 'quick',
    direction,
  });
  if (!request) return false;
  if (request.expanded) {
    return context.sort.canOpen && context.sort.open(request);
  }
  if (!request.selected.available) return false;
  return applySpreadsheetSort(context, {
    sheetId: context.targetSheetId,
    range: request.selected.range,
    orientation: 'top-to-bottom',
    hasHeader: request.selected.scope?.hasHeader ?? false,
    ...(request.selected.scope ? { scope: request.selected.scope } : {}),
    keys: [{ index: request.activeColumn, direction }],
  });
}

function createSpreadsheetSortOpenRequest(
  context: SpreadsheetCommandContext,
  intent: SpreadsheetSortIntent,
): SpreadsheetSortOpenRequest | null {
  if (!canEditSpreadsheetSelection(context)) return null;
  const sheet = context.content.sheets.find(
    (candidate) => candidate.id === context.targetSheetId,
  );
  const selectedRange = normalizeSpreadsheetCellRange(
    spreadsheetLiveCommandRange(context),
  );
  if (!sheet || !selectedRange) return null;
  const plan = createSpreadsheetSortRangePlan(sheet, selectedRange);
  if (!plan) return null;
  const selected = {
    range: plan.selectedRange,
    available: spreadsheetSortRangeCanApply(
      context,
      sheet,
      plan.selectedRange,
      intent,
    ),
    ...spreadsheetSortCandidateScope(sheet, plan.selectedRange),
  };
  const expanded = plan.expandedRange
    ? {
        range: plan.expandedRange,
        available: spreadsheetSortRangeCanApply(
          context,
          sheet,
          plan.expandedRange,
          intent,
        ),
        ...spreadsheetSortCandidateScope(sheet, plan.expandedRange),
      }
    : undefined;
  if (!selected.available && !expanded?.available) return null;
  const selection = spreadsheetLiveCommandSelection(context);
  const activeColumn = selection?.column_focus ?? selectedRange.column[0];
  const activeRow = selection?.row_focus ?? selectedRange.row[0];
  if (
    activeColumn < selectedRange.column[0] ||
    activeColumn > selectedRange.column[1] ||
    activeRow < selectedRange.row[0] ||
    activeRow > selectedRange.row[1]
  ) {
    return null;
  }
  return {
    sheetId: context.targetSheetId,
    activeColumn,
    activeRow,
    intent,
    selected,
    ...(expanded ? { expanded } : {}),
  };
}

function spreadsheetSortRangeCanApply(
  context: SpreadsheetCommandContext,
  sheet: WorkSpreadsheetSheet,
  range: SpreadsheetCellRange,
  intent: SpreadsheetSortIntent,
): boolean {
  return (
    spreadsheetSortRangeCanRun(range, intent) &&
    canMutateSpreadsheetCellRange(sheet, range) &&
    !spreadsheetSortRangeHasStructuralConflict(context, range)
  );
}

function canApplySpreadsheetSort(
  context: SpreadsheetCommandContext,
  request: SpreadsheetSortRequest,
): boolean {
  if (!canEditSpreadsheetSelection(context)) return false;
  const validation = validateSpreadsheetSortRequest(request);
  if (!validation.ok || request.sheetId !== context.targetSheetId) return false;
  const sheet = context.content.sheets.find(
    (candidate) => candidate.id === context.targetSheetId,
  );
  if (!sheet) return false;
  const owned = spreadsheetSortOwnedRangeForExactRange(
    sheet,
    validation.request.range,
  );
  if (
    owned &&
    spreadsheetSortRequiresFilterReconciliation(sheet, owned) &&
    !spreadsheetFilterReconciliationIsBounded(sheet)
  ) {
    return false;
  }
  if (!spreadsheetSortRequestMatchesOwnedRange(validation.request, owned)) {
    return false;
  }
  if (
    spreadsheetSortRangeHasStructuralConflict(
      context,
      validation.request.range,
      owned,
    )
  ) {
    return false;
  }
  return (
    spreadsheetCellRangesEqual(
      spreadsheetLiveCommandRange(context),
      validation.request.range,
    ) ||
    context.sort.canApply(
      validation.request,
      spreadsheetLiveCommandRange(context),
    )
  );
}

function applySpreadsheetSort(
  context: SpreadsheetCommandContext,
  request: SpreadsheetSortRequest,
): boolean {
  if (!context.workbook || !canApplySpreadsheetSort(context, request)) {
    return false;
  }
  const validation = validateSpreadsheetSortRequest(request);
  if (!validation.ok) return false;
  const sheet = context.content.sheets.find(
    (candidate) => candidate.id === request.sheetId,
  );
  if (
    !sheet ||
    !canMutateSpreadsheetCellRange(sheet, validation.request.range)
  ) {
    return false;
  }
  const owned = spreadsheetSortOwnedRangeForExactRange(
    sheet,
    validation.request.range,
  );
  let liveRows: (Cell | null)[][] | null = null;
  try {
    liveRows = context.workbook.getCellsByRange(validation.request.range, {
      id: request.sheetId,
    });
  } catch {
    // Filtered native views may not expose every hidden row through this API.
  }
  const rows =
    liveRows &&
    spreadsheetSortRowsMatchRange(liveRows, validation.request.range)
      ? liveRows
      : spreadsheetSortRowsFromSheet(sheet, validation.request.range);
  if (!rows) return false;
  try {
    const appearances = createSpreadsheetSortAppearanceRows(
      sheet,
      validation.request.range,
      rows,
    );
    const result = sortSpreadsheetMatrix(rows, validation.request, appearances);
    if (!result.ok) return false;
    if (owned && spreadsheetSortRequiresFilterReconciliation(sheet, owned)) {
      const sorted = spreadsheetSheetWithSortedRange(
        sheet,
        validation.request.range,
        result.rows,
      );
      const reconciled = reconcileSpreadsheetFiltersAfterSort(
        sorted,
        sheet,
        validation.request.range,
        result.sourceIndexes,
        {
          dateSystem: context.content.dateSystem,
          now: new Date(),
        },
      );
      if (!reconciled) return false;
      context.onChange({
        ...context.content,
        sheets: context.content.sheets.map((candidate) =>
          candidate.id === sheet.id ? reconciled : candidate,
        ),
      });
      return true;
    }
    context.workbook.setCellValuesByRange(
      result.rows,
      validation.request.range,
      { id: request.sheetId },
    );
    return true;
  } catch {
    return false;
  }
}

function spreadsheetSheetWithSortedRange(
  sheet: WorkSpreadsheetSheet,
  range: SpreadsheetCellRange,
  rows: readonly (readonly (Cell | null)[])[],
): WorkSpreadsheetSheet {
  const data = sheet.data?.slice();
  const celldata = sheet.celldata?.filter(
    (entry) =>
      entry.r < range.row[0] ||
      entry.r > range.row[1] ||
      entry.c < range.column[0] ||
      entry.c > range.column[1],
  );
  rows.forEach((source, rowOffset) => {
    const row = range.row[0] + rowOffset;
    const target = data
      ? (sheet.data?.[row]?.slice() ?? [])
      : ([] as (Cell | null)[]);
    source.forEach((cell, columnOffset) => {
      const column = range.column[0] + columnOffset;
      if (cell) target[column] = cell;
      else delete target[column];
      if (cell && celldata) celldata.push({ r: row, c: column, v: cell });
    });
    if (data) data[row] = target;
  });
  if (data) {
    celldata?.sort((left, right) => left.r - right.r || left.c - right.c);
    return { ...sheet, data, ...(celldata ? { celldata } : {}) };
  }
  if (celldata) {
    celldata.sort((left, right) => left.r - right.r || left.c - right.c);
    return { ...sheet, celldata };
  }
  const generated: NonNullable<WorkSpreadsheetSheet['data']> = [];
  rows.forEach((source, rowOffset) => {
    const target: (Cell | null)[] = [];
    source.forEach((cell, columnOffset) => {
      if (cell) target[range.column[0] + columnOffset] = cell;
    });
    generated[range.row[0] + rowOffset] = target;
  });
  return { ...sheet, data: generated };
}

function spreadsheetSortRequiresFilterReconciliation(
  sheet: WorkSpreadsheetSheet,
  owned: SpreadsheetSortOwnedRange,
): boolean {
  if (owned.scope.kind === 'auto-filter') {
    return Object.keys(sheet.filter ?? {}).length > 0;
  }
  const tableId = owned.scope.tableId;
  const table = sheet.tables?.find((candidate) => candidate.id === tableId);
  return Boolean(table?.headerRow && table.filters.length);
}

function spreadsheetSortRangeCanRun(
  range: { column: readonly number[]; row: readonly number[] },
  intent: SpreadsheetSortIntent,
): boolean {
  const rowStart = Math.min(range.row[0] ?? 0, range.row[1] ?? 0);
  const rowEnd = Math.max(range.row[0] ?? 0, range.row[1] ?? 0);
  const columnStart = Math.min(range.column[0] ?? 0, range.column[1] ?? 0);
  const columnEnd = Math.max(range.column[0] ?? 0, range.column[1] ?? 0);
  const rows = rowEnd - rowStart + 1;
  const columns = columnEnd - columnStart + 1;
  const area = (rowEnd - rowStart + 1) * (columnEnd - columnStart + 1);
  return (
    (rows >= 2 || (intent.type === 'custom' && columns >= 2)) &&
    Number.isSafeInteger(area) &&
    area <= MAX_SPREADSHEET_SORT_CELLS
  );
}

function spreadsheetSortRangeHasStructuralConflict(
  context: SpreadsheetCommandContext,
  range: SpreadsheetCellRange,
  exactOwnedRange?: SpreadsheetSortOwnedRange | null,
): boolean {
  const sheet = context.content.sheets.find(
    (candidate) => candidate.id === context.targetSheetId,
  );
  if (!sheet) return true;
  const ownedRange =
    exactOwnedRange === undefined
      ? spreadsheetSortOwnedRangeForExactRange(sheet, range)
      : exactOwnedRange;
  if (
    (sheet.tables ?? []).some(
      (table) =>
        spreadsheetCellRangesIntersect(table.range, range) &&
        !(
          ownedRange?.scope.kind === 'table' &&
          ownedRange.scope.tableId === table.id
        ),
    )
  ) {
    return true;
  }
  const autoFilter = normalizeSpreadsheetCellRange(
    sheet.filter_select ?? { row: [], column: [] },
  );
  return Boolean(
    (autoFilter &&
      spreadsheetCellRangesIntersect(autoFilter, range) &&
      ownedRange?.scope.kind !== 'auto-filter') ||
      spreadsheetSortRangeIntersectsHyperlink(sheet, range) ||
      spreadsheetSortRangeIntersectsFormulaMetadata(sheet, range) ||
      spreadsheetSortRangeIntersectsBorderMetadata(sheet, range),
  );
}

function spreadsheetSortCandidateScope(
  sheet: WorkSpreadsheetSheet,
  range: SpreadsheetCellRange,
): { scope?: SpreadsheetSortOwnedRange['scope'] } {
  const owned = spreadsheetSortOwnedRangeForExactRange(sheet, range);
  return owned ? { scope: owned.scope } : {};
}

function spreadsheetSortRequestMatchesOwnedRange(
  request: SpreadsheetSortNormalizedRequest,
  owned: SpreadsheetSortOwnedRange | null,
): boolean {
  if (!owned) return request.scope === undefined;
  return Boolean(
    request.scope &&
      spreadsheetSortOwnedScopesEqual(request.scope, owned.scope) &&
      request.orientation === 'top-to-bottom' &&
      request.hasHeader === owned.scope.hasHeader,
  );
}

function spreadsheetSortOwnedScopesEqual(
  left: SpreadsheetSortOwnedRange['scope'],
  right: SpreadsheetSortOwnedRange['scope'],
): boolean {
  return (
    left.kind === right.kind &&
    left.hasHeader === right.hasHeader &&
    (left.kind !== 'table' ||
      (right.kind === 'table' && left.tableId === right.tableId))
  );
}

function spreadsheetSortRangeIntersectsHyperlink(
  sheet: WorkSpreadsheetSheet,
  range: SpreadsheetCellRange,
): boolean {
  for (const key of Object.keys(sheet.hyperlink ?? {})) {
    const match = /^(\d+)_(\d+)$/.exec(key);
    if (!match) return true;
    const point: SpreadsheetCellRange = {
      row: [Number(match[1]), Number(match[1])],
      column: [Number(match[2]), Number(match[2])],
    };
    if (spreadsheetCellRangesIntersect(point, range)) return true;
  }
  return false;
}

function spreadsheetSortRangeIntersectsFormulaMetadata(
  sheet: WorkSpreadsheetSheet,
  range: SpreadsheetCellRange,
): boolean {
  const references = [
    ...Object.keys(sheet.formulaMetadata?.sourceFormulas ?? {}),
    ...(sheet.formulaMetadata?.ranges ?? []).map((item) => item.reference),
  ];
  for (const reference of references) {
    const parsed = parseSpreadsheetCellRanges(reference);
    if (!parsed) return true;
    if (
      parsed.some((candidate) =>
        spreadsheetCellRangesIntersect(candidate, range),
      )
    ) {
      return true;
    }
  }
  return false;
}

function spreadsheetSortRangeIntersectsBorderMetadata(
  sheet: WorkSpreadsheetSheet,
  range: SpreadsheetCellRange,
): boolean {
  const records = sheet.config?.borderInfo;
  if (!Array.isArray(records)) return false;
  for (const record of records as unknown[]) {
    if (!isRecord(record)) return true;
    if (record.rangeType === 'cell') {
      if (!isRecord(record.value)) return true;
      const row = finiteSpreadsheetSortIndex(record.value.row_index);
      const column = finiteSpreadsheetSortIndex(record.value.col_index);
      if (row === null || column === null) return true;
      if (
        spreadsheetCellRangesIntersect(
          { row: [row, row], column: [column, column] },
          range,
        )
      ) {
        return true;
      }
      continue;
    }
    if (record.rangeType !== 'range' || !Array.isArray(record.range)) {
      return true;
    }
    for (const candidate of record.range) {
      const normalized = isRecord(candidate)
        ? normalizeSpreadsheetCellRange({
            row: Array.isArray(candidate.row) ? candidate.row : [],
            column: Array.isArray(candidate.column) ? candidate.column : [],
          })
        : null;
      if (!normalized) return true;
      if (spreadsheetCellRangesIntersect(normalized, range)) return true;
    }
  }
  return false;
}

function finiteSpreadsheetSortIndex(value: unknown): number | null {
  return typeof value === 'number' && Number.isSafeInteger(value) && value >= 0
    ? value
    : null;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value && typeof value === 'object' && !Array.isArray(value));
}
