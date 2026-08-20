import { spreadsheetCommandCatalog } from './spreadsheet-command-catalog';
import type {
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import { isSpreadsheetNativeTextUndoTarget } from './spreadsheet-editor-support';
import { isSpreadsheetGridKeyboardTarget } from './spreadsheet-keyboard-navigation';
import {
  createOfficeEditorExtension,
  type OfficeEditorExtension,
} from './office-editor-extension';
import { isOfficeShortcutBlocked } from './office-shortcuts';

type SpreadsheetNavigationCommand = 'find' | 'go-to';

export function createSpreadsheetNavigationExtension(): OfficeEditorExtension<
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands
> {
  const findShortcut = spreadsheetCommandCatalog.find.shortcut.editor[0];
  const [goToShortcut, goToFunctionKey] =
    spreadsheetCommandCatalog.goTo.shortcut.editor;
  return createOfficeEditorExtension<
    SpreadsheetCommandContext,
    SpreadsheetEditorCommands
  >({
    name: 'spreadsheetNavigation',
    addCommands: () => ({
      openFind: {
        canExecute: (context) =>
          context.editable && context.navigation.canOpenFind,
        execute: (context) =>
          context.editable &&
          context.navigation.canOpenFind &&
          context.navigation.openFind(),
      },
      openGoTo: {
        canExecute: (context) =>
          context.editable && context.navigation.canOpenGoTo,
        execute: (context) =>
          context.editable &&
          context.navigation.canOpenGoTo &&
          context.navigation.openGoTo(),
      },
    }),
    addKeyboardShortcuts: () => ({
      [findShortcut]: ({ can, commands }, event) =>
        runSpreadsheetNavigationShortcut(
          event,
          'find',
          can.openFind,
          commands.openFind,
        ),
      [goToShortcut]: ({ can, commands }, event) =>
        runSpreadsheetNavigationShortcut(
          event,
          'go-to',
          can.openGoTo,
          commands.openGoTo,
        ),
      [goToFunctionKey]: ({ can, commands }, event) =>
        runSpreadsheetNavigationShortcut(
          event,
          'go-to',
          can.openGoTo,
          commands.openGoTo,
        ),
    }),
  });
}

function runSpreadsheetNavigationShortcut(
  event: KeyboardEvent,
  command: SpreadsheetNavigationCommand,
  canExecute: () => boolean,
  execute: () => boolean,
): boolean {
  const findSurface =
    command === 'find' &&
    event.target instanceof Element &&
    Boolean(event.target.closest('.work-spreadsheet-find-bar'));
  if (
    event.repeat ||
    event.isComposing ||
    isOfficeShortcutBlocked(event.target) ||
    (!findSurface && isSpreadsheetNativeTextUndoTarget(event.target)) ||
    (!findSurface && !isSpreadsheetGridKeyboardTarget(event.target)) ||
    !canExecute()
  ) {
    return false;
  }
  return execute();
}
