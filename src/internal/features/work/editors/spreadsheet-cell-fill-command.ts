import type { Sheet } from '@fortune-sheet/core';
import {
  canApplySpreadsheetCellFill,
  planSpreadsheetCellFill,
  type SpreadsheetCellFillDirection,
  type SpreadsheetCellFillPlan,
} from './spreadsheet-cell-fill';
import type {
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import {
  createOfficeEditorExtension,
  type OfficeEditorExtension,
} from './office-editor-extension';
import { spreadsheetLiveCommandSelections } from './spreadsheet-command-selection';

export function createSpreadsheetCellFillExtension(): OfficeEditorExtension<
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands
> {
  return createOfficeEditorExtension<
    SpreadsheetCommandContext,
    SpreadsheetEditorCommands
  >({
    name: 'spreadsheetCellFill',
    addCommands: () => ({
      fillSelectedCells: {
        canExecute: (context, direction) =>
          spreadsheetSelectedCellFillPlan(context, direction) !== null,
        execute: fillSelectedCells,
      },
    }),
  });
}

function fillSelectedCells(
  context: SpreadsheetCommandContext,
  direction: SpreadsheetCellFillDirection,
): boolean {
  const plan = spreadsheetSelectedCellFillPlan(context, direction);
  if (!plan || !context.workbook) return false;
  let rollbackSparseRows: (() => void) | null = null;
  try {
    const sheet = context.workbook.getSheet({ id: context.targetSheetId });
    rollbackSparseRows = materializeSpreadsheetCellFillRows(sheet, plan);
    context.workbook.autoFillCell(
      plan.copyRange,
      plan.applyRange,
      plan.direction,
    );
    return true;
  } catch {
    rollbackSparseRows?.();
    return false;
  }
}

function materializeSpreadsheetCellFillRows(
  sheet: Sheet,
  plan: SpreadsheetCellFillPlan,
): () => void {
  const hadDataProperty = Object.hasOwn(sheet, 'data');
  const originalData = sheet.data;
  const data = originalData ?? [];
  const originalLength = data.length;
  const createdRows: number[] = [];
  const firstRow = Math.min(plan.copyRange.row[0], plan.applyRange.row[0]);
  const lastRow = Math.max(plan.copyRange.row[1], plan.applyRange.row[1]);

  if (!originalData) sheet.data = data;
  for (let row = firstRow; row <= lastRow; row += 1) {
    if (Array.isArray(data[row])) continue;
    data[row] = [];
    createdRows.push(row);
  }

  return () => {
    for (const row of createdRows) delete data[row];
    data.length = originalLength;
    if (originalData) return;
    if (hadDataProperty) {
      sheet.data = undefined;
    } else {
      delete sheet.data;
    }
  };
}

function spreadsheetSelectedCellFillPlan(
  context: SpreadsheetCommandContext,
  direction: SpreadsheetCellFillDirection,
): SpreadsheetCellFillPlan | null {
  if (
    !context.editable ||
    !context.workbook ||
    !context.targetSheetId ||
    context.targetSheetId !== context.activeSheetId
  ) {
    return null;
  }
  const selections = spreadsheetLiveCommandSelections(context);
  if (selections?.length !== 1) return null;
  const selection = selections[0];
  if (!selection) return null;
  const plan = planSpreadsheetCellFill(selection, direction);
  if (!plan) return null;
  const sheet = context.content.sheets.find(
    (candidate) => candidate.id === context.targetSheetId,
  );
  return canApplySpreadsheetCellFill(sheet, plan) ? plan : null;
}
