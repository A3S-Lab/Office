import type { WorkSpreadsheetContent } from '../work-types';
import {
  createOfficeEditorExtension,
  type OfficeEditorExtension,
} from './office-editor-extension';
import type {
  SpreadsheetCommandContext,
  SpreadsheetCommandRange,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import {
  moveSpreadsheetKeyboardSelection,
  scopeSpreadsheetKeyboardSelection,
  spreadsheetSelectionContainsFocus,
} from './spreadsheet-keyboard-navigation';
import {
  rememberSpreadsheetCommandSelection,
  spreadsheetLiveCommandSelection,
} from './spreadsheet-command-selection';

interface SpreadsheetSelectionNavigationStorage {
  focus: { column: number; row: number } | null;
}

export function createSpreadsheetSelectionNavigationExtension(): OfficeEditorExtension<
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands
> {
  return createOfficeEditorExtension<
    SpreadsheetCommandContext,
    SpreadsheetEditorCommands,
    SpreadsheetSelectionNavigationStorage
  >({
    name: 'spreadsheetSelectionNavigation',
    addStorage: () => ({ focus: null }),
    addCommands: ({ storage }) => ({
      moveSelection: {
        canExecute: canNavigateSpreadsheetSelection,
        execute: (context, move, extend) =>
          updateSpreadsheetSelection(context, storage, (sheet, selection) =>
            moveSpreadsheetKeyboardSelection(sheet, selection, move, extend),
          ),
      },
      selectCellRange: {
        canExecute: canNavigateSpreadsheetSelection,
        execute: (context, scope) =>
          updateSpreadsheetSelection(context, storage, (sheet, selection) =>
            scopeSpreadsheetKeyboardSelection(sheet, selection, scope),
          ),
      },
    }),
  });
}

function canNavigateSpreadsheetSelection(
  context: SpreadsheetCommandContext,
): boolean {
  return Boolean(
    context.workbook &&
      context.targetSheetId &&
      context.content.sheets.some(
        (sheet) => sheet.id === context.targetSheetId,
      ),
  );
}

function updateSpreadsheetSelection(
  context: SpreadsheetCommandContext,
  storage: SpreadsheetSelectionNavigationStorage,
  update: (
    sheet: WorkSpreadsheetContent['sheets'][number],
    selection: SpreadsheetCommandRange,
  ) => SpreadsheetCommandRange,
): boolean {
  if (!context.workbook || !context.targetSheetId) return false;
  const sheet = context.content.sheets.find(
    (candidate) => candidate.id === context.targetSheetId,
  );
  if (!sheet) return false;
  const liveSelection =
    spreadsheetLiveCommandSelection(context) ?? context.fallbackRange;
  const rememberedFocus = storage.focus;
  const selection =
    rememberedFocus &&
    spreadsheetSelectionContainsFocus(liveSelection, rememberedFocus)
      ? {
          ...liveSelection,
          row_focus: rememberedFocus.row,
          column_focus: rememberedFocus.column,
        }
      : liveSelection;
  const next = update(sheet, selection);
  try {
    context.workbook.setSelection([next], { id: context.targetSheetId });
    rememberSpreadsheetCommandSelection(context, next);
    storage.focus = {
      row: next.row_focus ?? next.row[1] ?? next.row[0] ?? 0,
      column: next.column_focus ?? next.column[1] ?? next.column[0] ?? 0,
    };
    return true;
  } catch {
    return false;
  }
}
