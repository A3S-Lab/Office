import { isOfficeShortcutBlocked } from './office-shortcuts';
import type {
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import { spreadsheetCommandCatalog } from './spreadsheet-command-catalog';
import {
  isSpreadsheetNativeTextUndoTarget,
  spreadsheetSingleRange,
} from './spreadsheet-editor-support';
import {
  applySpreadsheetHyperlink,
  canRemoveSpreadsheetHyperlink,
  createSpreadsheetHyperlinkDialogSource,
  removeSpreadsheetHyperlink,
  type SpreadsheetHyperlinkCell,
  type SpreadsheetHyperlinkRequest,
  validateSpreadsheetHyperlinkRequest,
} from './spreadsheet-hyperlink';
import { isSpreadsheetGridKeyboardTarget } from './spreadsheet-keyboard-navigation';
import {
  createOfficeEditorExtension,
  type OfficeEditorExtension,
} from './office-editor-extension';

export function createSpreadsheetHyperlinkExtension(): OfficeEditorExtension<
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands
> {
  return createOfficeEditorExtension<
    SpreadsheetCommandContext,
    SpreadsheetEditorCommands
  >({
    name: 'spreadsheetHyperlink',
    addCommands: () => ({
      openHyperlink: {
        canExecute: canOpenSpreadsheetHyperlink,
        execute: openSpreadsheetHyperlink,
      },
      applyHyperlink: {
        canExecute: (context, request) =>
          context.editable &&
          validateSpreadsheetHyperlinkRequest(context.content, request).ok,
        execute: applySpreadsheetHyperlinkCommand,
      },
      removeHyperlink: {
        canExecute: (context, target) =>
          context.editable &&
          canRemoveSpreadsheetHyperlink(context.content, target),
        execute: removeSpreadsheetHyperlinkCommand,
      },
    }),
    addKeyboardShortcuts: () => ({
      [spreadsheetCommandCatalog.hyperlink.shortcut.editor[0]]: (
        { can, commands },
        event,
      ) =>
        runSpreadsheetHyperlinkShortcut(
          event,
          can.openHyperlink,
          commands.openHyperlink,
        ),
    }),
  });
}

function canOpenSpreadsheetHyperlink(
  context: SpreadsheetCommandContext,
): boolean {
  const target = liveSpreadsheetHyperlinkCell(context);
  return Boolean(
    context.editable &&
      context.hyperlink.canOpen &&
      context.workbook &&
      target &&
      target.sheetId === context.activeSheetId &&
      createSpreadsheetHyperlinkDialogSource(context.content, target),
  );
}

function openSpreadsheetHyperlink(context: SpreadsheetCommandContext): boolean {
  const target = liveSpreadsheetHyperlinkCell(context);
  return Boolean(
    target &&
      canOpenSpreadsheetHyperlink(context) &&
      context.hyperlink.open(target),
  );
}

function applySpreadsheetHyperlinkCommand(
  context: SpreadsheetCommandContext,
  request: SpreadsheetHyperlinkRequest,
): boolean {
  if (!context.editable) return false;
  const next = applySpreadsheetHyperlink(context.content, request);
  if (!next) return false;
  context.onChange(next);
  return true;
}

function removeSpreadsheetHyperlinkCommand(
  context: SpreadsheetCommandContext,
  target: SpreadsheetHyperlinkCell,
): boolean {
  if (!context.editable) return false;
  const next = removeSpreadsheetHyperlink(context.content, target);
  if (!next) return false;
  context.onChange(next);
  return true;
}

function liveSpreadsheetHyperlinkCell(
  context: SpreadsheetCommandContext,
): SpreadsheetHyperlinkCell | null {
  if (!context.workbook || !context.targetSheetId) return null;
  const selection =
    context.workbook.getSelection()?.at(-1) ?? context.fallbackRange;
  const range = spreadsheetSingleRange(selection);
  return {
    sheetId: context.targetSheetId,
    row: clampSpreadsheetHyperlinkFocus(selection.row_focus, range.row),
    column: clampSpreadsheetHyperlinkFocus(
      selection.column_focus,
      range.column,
    ),
  };
}

function clampSpreadsheetHyperlinkFocus(
  value: unknown,
  range: readonly number[],
): number {
  const minimum = range[0] ?? 0;
  const maximum = range[1] ?? minimum;
  const focus =
    typeof value === 'number' && Number.isFinite(value)
      ? Math.trunc(value)
      : minimum;
  return Math.min(maximum, Math.max(minimum, focus));
}

function runSpreadsheetHyperlinkShortcut(
  event: KeyboardEvent,
  canExecute: SpreadsheetEditorCommands['openHyperlink'],
  execute: SpreadsheetEditorCommands['openHyperlink'],
): boolean {
  if (
    event.repeat ||
    event.isComposing ||
    isOfficeShortcutBlocked(event.target) ||
    isSpreadsheetNativeTextUndoTarget(event.target) ||
    !isSpreadsheetGridKeyboardTarget(event.target) ||
    !canExecute()
  ) {
    return false;
  }
  return execute();
}
