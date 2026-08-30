import {
  createOfficeEditorExtension,
  type OfficeEditorExtension,
} from './office-editor-extension';
import { parseSpreadsheetCellRanges } from '../work-spreadsheet-ranges';
import type { WorkSpreadsheetSheet } from '../work-types';
import { canMutateSpreadsheetCellRange } from './spreadsheet-cell-mutation-guard';
import {
  canEditSpreadsheetSelection,
  spreadsheetLiveCommandRange,
  spreadsheetLiveCommandSelection,
} from './spreadsheet-command-selection';
import type {
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import {
  normalizeSpreadsheetCellRange,
  spreadsheetCellRangesIntersect,
  type SpreadsheetCellRange,
} from './spreadsheet-cell-range';
import {
  MAX_SPREADSHEET_SORT_CELLS,
  sortSpreadsheetRows,
  type SpreadsheetSortDirection,
  type SpreadsheetSortRequest,
  validateSpreadsheetSortRequest,
} from './spreadsheet-sort';

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
  if (!canEditSpreadsheetSelection(context) || !context.sort.canOpen) {
    return false;
  }
  const range = spreadsheetLiveCommandRange(context);
  const normalized = normalizeSpreadsheetCellRange(range);
  return Boolean(
    normalized &&
      spreadsheetSortRangeCanRun(normalized, false) &&
      !spreadsheetSortRangeHasStructuralConflict(context, normalized),
  );
}

function openSpreadsheetSort(context: SpreadsheetCommandContext): boolean {
  if (!canOpenSpreadsheetSort(context)) return false;
  const range = spreadsheetLiveCommandRange(context);
  const activeColumn =
    spreadsheetLiveCommandSelection(context)?.column_focus ?? range.column[0];
  return context.sort.open({
    sheetId: context.targetSheetId,
    range: {
      row: [range.row[0] ?? 0, range.row[1] ?? range.row[0] ?? 0],
      column: [range.column[0] ?? 0, range.column[1] ?? range.column[0] ?? 0],
    },
    activeColumn,
  });
}

function canQuickSortSpreadsheet(
  context: SpreadsheetCommandContext,
  direction: SpreadsheetSortDirection,
): boolean {
  if (
    (direction !== 'ascending' && direction !== 'descending') ||
    !canEditSpreadsheetSelection(context)
  ) {
    return false;
  }
  const range = normalizeSpreadsheetCellRange(
    spreadsheetLiveCommandRange(context),
  );
  return Boolean(
    range &&
      spreadsheetSortRangeCanRun(range, false) &&
      !spreadsheetSortRangeHasStructuralConflict(context, range),
  );
}

function quickSortSpreadsheet(
  context: SpreadsheetCommandContext,
  direction: SpreadsheetSortDirection,
): boolean {
  if (!canQuickSortSpreadsheet(context, direction)) return false;
  const range = spreadsheetLiveCommandRange(context);
  return applySpreadsheetSort(context, {
    sheetId: context.targetSheetId,
    range: {
      row: [range.row[0] ?? 0, range.row[1] ?? range.row[0] ?? 0],
      column: [range.column[0] ?? 0, range.column[1] ?? range.column[0] ?? 0],
    },
    hasHeader: false,
    keys: [{ column: range.column[0] ?? 0, direction }],
  });
}

function canApplySpreadsheetSort(
  context: SpreadsheetCommandContext,
  request: SpreadsheetSortRequest,
): boolean {
  if (!canEditSpreadsheetSelection(context)) return false;
  const validation = validateSpreadsheetSortRequest(request);
  if (!validation.ok || request.sheetId !== context.targetSheetId) return false;
  if (
    spreadsheetSortRangeHasStructuralConflict(context, validation.request.range)
  ) {
    return false;
  }
  return sameSpreadsheetSortRange(
    spreadsheetLiveCommandRange(context),
    validation.request.range,
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
  if (!canMutateSpreadsheetCellRange(sheet, validation.request.range)) {
    return false;
  }
  try {
    const rows = context.workbook.getCellsByRange(validation.request.range, {
      id: request.sheetId,
    });
    const result = sortSpreadsheetRows(rows, validation.request);
    if (!result.ok) return false;
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

function spreadsheetSortRangeCanRun(
  range: { column: readonly number[]; row: readonly number[] },
  hasHeader: boolean,
): boolean {
  const rowStart = Math.min(range.row[0] ?? 0, range.row[1] ?? 0);
  const rowEnd = Math.max(range.row[0] ?? 0, range.row[1] ?? 0);
  const columnStart = Math.min(range.column[0] ?? 0, range.column[1] ?? 0);
  const columnEnd = Math.max(range.column[0] ?? 0, range.column[1] ?? 0);
  const rows = rowEnd - rowStart + 1 - (hasHeader ? 1 : 0);
  const area = (rowEnd - rowStart + 1) * (columnEnd - columnStart + 1);
  return (
    rows >= 2 &&
    Number.isSafeInteger(area) &&
    area <= MAX_SPREADSHEET_SORT_CELLS
  );
}

function spreadsheetSortRangeHasStructuralConflict(
  context: SpreadsheetCommandContext,
  range: SpreadsheetCellRange,
): boolean {
  const sheet = context.content.sheets.find(
    (candidate) => candidate.id === context.targetSheetId,
  );
  if (!sheet) return true;
  if (
    (sheet.tables ?? []).some((table) =>
      spreadsheetCellRangesIntersect(table.range, range),
    )
  ) {
    return true;
  }
  const autoFilter = normalizeSpreadsheetCellRange(
    sheet.filter_select ?? { row: [], column: [] },
  );
  return Boolean(
    (autoFilter && spreadsheetCellRangesIntersect(autoFilter, range)) ||
      spreadsheetSortRangeIntersectsHyperlink(sheet, range) ||
      spreadsheetSortRangeIntersectsFormulaMetadata(sheet, range) ||
      spreadsheetSortRangeIntersectsBorderMetadata(sheet, range),
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

function sameSpreadsheetSortRange(
  left: { column: readonly number[]; row: readonly number[] },
  right: { column: readonly number[]; row: readonly number[] },
): boolean {
  return (
    Math.min(left.row[0] ?? 0, left.row[1] ?? 0) === right.row[0] &&
    Math.max(left.row[0] ?? 0, left.row[1] ?? 0) === right.row[1] &&
    Math.min(left.column[0] ?? 0, left.column[1] ?? 0) === right.column[0] &&
    Math.max(left.column[0] ?? 0, left.column[1] ?? 0) === right.column[1]
  );
}
