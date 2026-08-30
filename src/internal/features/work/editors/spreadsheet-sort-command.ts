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
  type SpreadsheetSortOpenRequest,
  type SpreadsheetSortRequest,
  sortSpreadsheetRows,
  validateSpreadsheetSortRequest,
} from './spreadsheet-sort';
import { createSpreadsheetSortAppearanceRows } from './spreadsheet-sort-appearance';

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
    hasHeader: false,
    keys: [{ column: request.activeColumn, direction }],
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
    available: spreadsheetSortRangeCanApply(context, sheet, plan.selectedRange),
  };
  const expanded = plan.expandedRange
    ? {
        range: plan.expandedRange,
        available: spreadsheetSortRangeCanApply(
          context,
          sheet,
          plan.expandedRange,
        ),
      }
    : undefined;
  if (!selected.available && !expanded?.available) return null;
  const selection = spreadsheetLiveCommandSelection(context);
  const activeColumn = selection?.column_focus ?? selectedRange.column[0];
  if (
    activeColumn < selectedRange.column[0] ||
    activeColumn > selectedRange.column[1]
  ) {
    return null;
  }
  return {
    sheetId: context.targetSheetId,
    activeColumn,
    intent,
    selected,
    ...(expanded ? { expanded } : {}),
  };
}

function spreadsheetSortRangeCanApply(
  context: SpreadsheetCommandContext,
  sheet: WorkSpreadsheetSheet,
  range: SpreadsheetCellRange,
): boolean {
  return (
    spreadsheetSortRangeCanRun(range, false) &&
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
  if (
    spreadsheetSortRangeHasStructuralConflict(context, validation.request.range)
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
  try {
    const rows = context.workbook.getCellsByRange(validation.request.range, {
      id: request.sheetId,
    });
    const appearances = createSpreadsheetSortAppearanceRows(
      sheet,
      validation.request.range,
      rows,
    );
    const result = sortSpreadsheetRows(rows, validation.request, appearances);
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
