import type {
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import {
  planSpreadsheetAutoSum,
  type SpreadsheetAutoSumFunction,
  type SpreadsheetAutoSumPlan,
} from './spreadsheet-auto-sum';
import { canMutateSpreadsheetCellRanges } from './spreadsheet-cell-mutation-guard';
import { isSpreadsheetNativeTextUndoTarget } from './spreadsheet-editor-support';
import { isSpreadsheetGridKeyboardTarget } from './spreadsheet-keyboard-navigation';
import {
  createOfficeEditorExtension,
  type OfficeEditorExtension,
} from './office-editor-extension';
import { isOfficeShortcutBlocked } from './office-shortcuts';

export function createSpreadsheetAutoSumExtension(): OfficeEditorExtension<
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands
> {
  return createOfficeEditorExtension<
    SpreadsheetCommandContext,
    SpreadsheetEditorCommands
  >({
    name: 'spreadsheetAutoSum',
    addCommands: () => ({
      applyAutoSum: {
        canExecute: (context, functionName) =>
          spreadsheetSelectedAutoSumPlan(context, functionName) !== null,
        execute: applySpreadsheetAutoSum,
      },
    }),
    addKeyboardShortcuts: () => ({
      'Alt-equal': ({ can, commands }, event) =>
        runSpreadsheetAutoSumShortcut(
          event,
          can.applyAutoSum,
          commands.applyAutoSum,
        ),
    }),
  });
}

function applySpreadsheetAutoSum(
  context: SpreadsheetCommandContext,
  functionName: SpreadsheetAutoSumFunction,
): boolean {
  const plan = spreadsheetSelectedAutoSumPlan(context, functionName);
  if (!plan || !context.workbook) return false;
  try {
    context.workbook.batchCallApis([
      ...plan.writes.map((write) => ({
        name: 'setCellValuesByRange',
        args: [write.values, write.range, { id: context.targetSheetId }],
      })),
      {
        name: 'setSelection',
        args: [[plan.selection], { id: context.targetSheetId }],
      },
    ]);
  } catch {
    return false;
  }
  try {
    context.formulaBar?.setValue(plan.formulaBarValue);
  } catch {
    // The workbook mutation succeeded; a detached formula bar must not invert it.
  }
  return true;
}

function spreadsheetSelectedAutoSumPlan(
  context: SpreadsheetCommandContext,
  functionName: SpreadsheetAutoSumFunction,
): SpreadsheetAutoSumPlan | null {
  if (
    !context.editable ||
    !context.workbook ||
    !context.targetSheetId ||
    context.targetSheetId !== context.activeSheetId
  ) {
    return null;
  }
  const selections = context.workbook.getSelection();
  if (selections?.length !== 1) return null;
  const sheet = context.content.sheets.find(
    (candidate) => candidate.id === context.targetSheetId,
  );
  const plan = planSpreadsheetAutoSum(sheet, selections[0], functionName);
  if (
    !plan ||
    !canMutateSpreadsheetCellRanges(
      sheet,
      plan.writes.map((write) => write.range),
    )
  ) {
    return null;
  }
  return plan;
}

function runSpreadsheetAutoSumShortcut(
  event: KeyboardEvent,
  canExecute: (functionName: SpreadsheetAutoSumFunction) => boolean,
  execute: (functionName: SpreadsheetAutoSumFunction) => boolean,
): boolean {
  if (
    event.repeat ||
    event.isComposing ||
    isOfficeShortcutBlocked(event.target) ||
    isSpreadsheetNativeTextUndoTarget(event.target) ||
    !isSpreadsheetGridKeyboardTarget(event.target) ||
    !canExecute('sum')
  ) {
    return false;
  }
  return execute('sum');
}
