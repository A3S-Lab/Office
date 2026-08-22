import {
  createOfficeEditorExtension,
  type OfficeEditorExtension,
} from './office-editor-extension';
import { isOfficeShortcutBlocked } from './office-shortcuts';
import { spreadsheetCommandCatalog } from './spreadsheet-command-catalog';
import type {
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import {
  planSpreadsheetCopyFromAbove,
  type SpreadsheetCopyFromAboveKind,
  type SpreadsheetCopyFromAbovePlan,
} from './spreadsheet-copy-from-above';
import {
  finiteSpreadsheetSelection,
  isSpreadsheetNativeTextUndoTarget,
} from './spreadsheet-editor-support';
import { isSpreadsheetGridKeyboardTarget } from './spreadsheet-keyboard-navigation';

export function createSpreadsheetCopyFromAboveExtension(): OfficeEditorExtension<
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands
> {
  return createOfficeEditorExtension<
    SpreadsheetCommandContext,
    SpreadsheetEditorCommands
  >({
    name: 'spreadsheetCopyFromAbove',
    addCommands: () => ({
      copyCellFromAbove: {
        canExecute: (context, kind) =>
          spreadsheetSelectedCopyFromAbovePlan(context, kind) !== null,
        execute: copySpreadsheetCellFromAbove,
      },
    }),
    addKeyboardShortcuts: () => ({
      [spreadsheetCommandCatalog.copyFormulaFromAbove.shortcut.editor[0]]: (
        { can, commands },
        event,
      ) =>
        runSpreadsheetCopyFromAboveShortcut(
          event,
          'formula',
          can.copyCellFromAbove,
          commands.copyCellFromAbove,
        ),
      [spreadsheetCommandCatalog.copyValueFromAbove.shortcut.editor[0]]: (
        { can, commands },
        event,
      ) =>
        runSpreadsheetCopyFromAboveShortcut(
          event,
          'value',
          can.copyCellFromAbove,
          commands.copyCellFromAbove,
        ),
    }),
  });
}

function copySpreadsheetCellFromAbove(
  context: SpreadsheetCommandContext,
  kind: SpreadsheetCopyFromAboveKind,
): boolean {
  const plan = spreadsheetSelectedCopyFromAbovePlan(context, kind);
  if (!plan || !context.workbook || !context.targetSheetId) return false;

  try {
    context.workbook.batchCallApis([
      {
        name: 'setCellValuesByRange',
        args: [[[plan.value]], plan.targetRange, { id: context.targetSheetId }],
      },
    ]);
  } catch {
    return false;
  }

  try {
    context.formulaBar?.setValue(plan.formulaBarValue);
  } catch {
    // The native mutation succeeded; a detached formula bar is best effort.
  }
  return true;
}

function spreadsheetSelectedCopyFromAbovePlan(
  context: SpreadsheetCommandContext,
  kind: SpreadsheetCopyFromAboveKind,
): SpreadsheetCopyFromAbovePlan | null {
  const workbook = context.workbook;
  const sheetId = context.targetSheetId;
  if (
    !context.editable ||
    !workbook ||
    !sheetId ||
    sheetId !== context.activeSheetId
  ) {
    return null;
  }

  try {
    const selection = finiteSpreadsheetSelection(
      workbook.getSelection()?.at(-1) ?? context.fallbackRange,
    );
    const row = selection.row_focus ?? selection.row[0] ?? 0;
    const column = selection.column_focus ?? selection.column[0] ?? 0;
    const sheet = context.content.sheets.find(
      (candidate) => candidate.id === sheetId,
    );
    return planSpreadsheetCopyFromAbove(sheet, { row, column }, kind);
  } catch {
    return null;
  }
}

function runSpreadsheetCopyFromAboveShortcut(
  event: KeyboardEvent,
  kind: SpreadsheetCopyFromAboveKind,
  canExecute: SpreadsheetEditorCommands['copyCellFromAbove'],
  execute: SpreadsheetEditorCommands['copyCellFromAbove'],
): boolean {
  if (
    event.repeat ||
    event.isComposing ||
    isOfficeShortcutBlocked(event.target) ||
    isSpreadsheetNativeTextUndoTarget(event.target) ||
    !isSpreadsheetGridKeyboardTarget(event.target) ||
    !canExecute(kind)
  ) {
    return false;
  }
  return execute(kind);
}
