import {
  createOfficeEditorExtension,
  type OfficeEditorExtension,
} from './office-editor-extension';
import { isOfficeShortcutBlocked } from './office-shortcuts';
import { spreadsheetCommandCatalog } from './spreadsheet-command-catalog';
import type {
  SpreadsheetCommandContext,
  SpreadsheetEditorCanCommands,
  SpreadsheetEditorCommands,
  SpreadsheetStructureAxis,
} from './spreadsheet-command-controller';
import { isSpreadsheetNativeTextUndoTarget } from './spreadsheet-editor-support';
import { isSpreadsheetGridKeyboardTarget } from './spreadsheet-keyboard-navigation';

export function createSpreadsheetVisibilityShortcutExtension(): OfficeEditorExtension<
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands
> {
  return createOfficeEditorExtension<
    SpreadsheetCommandContext,
    SpreadsheetEditorCommands
  >({
    name: 'spreadsheetStructureVisibilityShortcuts',
    addKeyboardShortcuts: () => ({
      [spreadsheetCommandCatalog.hideRows.shortcut.editor[0]]: (
        { can, commands },
        event,
      ) => runVisibilityShortcut(event, can, commands, 'row', true),
      [spreadsheetCommandCatalog.hideColumns.shortcut.editor[0]]: (
        { can, commands },
        event,
      ) => runVisibilityShortcut(event, can, commands, 'column', true),
      [spreadsheetCommandCatalog.unhideRows.shortcut.editor[0]]: (
        { can, commands },
        event,
      ) => runVisibilityShortcut(event, can, commands, 'row', false),
      [spreadsheetCommandCatalog.unhideColumns.shortcut.editor[0]]: (
        { can, commands },
        event,
      ) => runVisibilityShortcut(event, can, commands, 'column', false),
    }),
  });
}

function runVisibilityShortcut(
  event: KeyboardEvent,
  can: SpreadsheetEditorCanCommands,
  commands: SpreadsheetEditorCommands,
  axis: SpreadsheetStructureAxis,
  hidden: boolean,
): boolean {
  if (
    event.repeat ||
    event.isComposing ||
    isOfficeShortcutBlocked(event.target) ||
    isSpreadsheetNativeTextUndoTarget(event.target) ||
    !isSpreadsheetGridKeyboardTarget(event.target) ||
    !can.setSelectedStructureHidden(axis, hidden)
  ) {
    return false;
  }
  return commands.setSelectedStructureHidden(axis, hidden);
}
