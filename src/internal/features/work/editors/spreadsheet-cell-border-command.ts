import { isOfficeShortcutBlocked } from './office-shortcuts';
import {
  canSetSpreadsheetCellBorders,
  type SpreadsheetCellBorderFormat,
  setSpreadsheetCellBorders,
} from './spreadsheet-cell-border';
import { spreadsheetCommandCatalog } from './spreadsheet-command-catalog';
import type {
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import {
  isSpreadsheetNativeTextUndoTarget,
  spreadsheetContentWithSelection,
  spreadsheetSingleRange,
} from './spreadsheet-editor-support';
import {
  createOfficeEditorExtension,
  type OfficeEditorExtension,
} from './office-editor-extension';

export function createSpreadsheetCellBorderExtension(): OfficeEditorExtension<
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands
> {
  const outsideShortcuts =
    spreadsheetCommandCatalog.borderOutside.shortcut.editor;
  const noneShortcuts = spreadsheetCommandCatalog.borderNone.shortcut.editor;
  return createOfficeEditorExtension<
    SpreadsheetCommandContext,
    SpreadsheetEditorCommands
  >({
    name: 'spreadsheetCellBorders',
    addCommands: () => ({
      setSelectedCellBorders: {
        canExecute: canSetSelectedCellBorders,
        execute: setSelectedCellBorders,
      },
    }),
    addKeyboardShortcuts: () => ({
      [outsideShortcuts[0]]: ({ can, commands }, event) =>
        runSpreadsheetBorderShortcut(
          event,
          can.setSelectedCellBorders,
          commands.setSelectedCellBorders,
          { target: 'outside', color: '#000000', style: 'thin' },
        ),
      [outsideShortcuts[1]]: ({ can, commands }, event) =>
        runSpreadsheetBorderShortcut(
          event,
          can.setSelectedCellBorders,
          commands.setSelectedCellBorders,
          { target: 'outside', color: '#000000', style: 'thin' },
        ),
      [noneShortcuts[0]]: ({ can, commands }, event) =>
        runSpreadsheetBorderShortcut(
          event,
          can.setSelectedCellBorders,
          commands.setSelectedCellBorders,
          { target: 'none', color: '#000000', style: 'thin' },
        ),
      [noneShortcuts[1]]: ({ can, commands }, event) =>
        runSpreadsheetBorderShortcut(
          event,
          can.setSelectedCellBorders,
          commands.setSelectedCellBorders,
          { target: 'none', color: '#000000', style: 'thin' },
        ),
    }),
  });
}

function canSetSelectedCellBorders(
  context: SpreadsheetCommandContext,
  format: SpreadsheetCellBorderFormat,
): boolean {
  return Boolean(
    context.editable &&
      context.workbook &&
      context.targetSheetId &&
      canSetSpreadsheetCellBorders(
        context.content,
        context.targetSheetId,
        liveSpreadsheetBorderRange(context),
        format,
      ),
  );
}

function setSelectedCellBorders(
  context: SpreadsheetCommandContext,
  format: SpreadsheetCellBorderFormat,
): boolean {
  if (!context.editable || !context.workbook || !context.targetSheetId) {
    return false;
  }
  const selection =
    context.workbook.getSelection()?.at(-1) ?? context.fallbackRange;
  const range = spreadsheetSingleRange(selection);
  if (
    !canSetSpreadsheetCellBorders(
      context.content,
      context.targetSheetId,
      range,
      format,
    )
  ) {
    return false;
  }
  const next = setSpreadsheetCellBorders(
    context.content,
    context.targetSheetId,
    range,
    format,
  );
  if (!next) return false;
  context.onChange(
    spreadsheetContentWithSelection(next, context.targetSheetId, selection),
  );
  return true;
}

function liveSpreadsheetBorderRange(context: SpreadsheetCommandContext) {
  return spreadsheetSingleRange(
    context.workbook?.getSelection()?.at(-1) ?? context.fallbackRange,
  );
}

function runSpreadsheetBorderShortcut(
  event: KeyboardEvent,
  canExecute: SpreadsheetEditorCommands['setSelectedCellBorders'],
  execute: SpreadsheetEditorCommands['setSelectedCellBorders'],
  format: SpreadsheetCellBorderFormat,
): boolean {
  if (
    event.repeat ||
    isOfficeShortcutBlocked(event.target) ||
    isSpreadsheetNativeTextUndoTarget(event.target) ||
    !canExecute(format)
  ) {
    return false;
  }
  return execute(format);
}
