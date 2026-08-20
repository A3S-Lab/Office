import { isOfficeShortcutBlocked } from './office-shortcuts';
import { isSpreadsheetNativeTextUndoTarget } from './spreadsheet-editor-support';

export interface SpreadsheetClipboardShortcutOptions {
  canExecute(): boolean;
  execute(): boolean;
}

export function runSpreadsheetClipboardShortcut(
  event: KeyboardEvent,
  canExecute: SpreadsheetClipboardShortcutOptions['canExecute'],
  execute: SpreadsheetClipboardShortcutOptions['execute'],
): boolean {
  if (
    event.defaultPrevented ||
    event.repeat ||
    event.altKey ||
    event.shiftKey ||
    !(event.metaKey || event.ctrlKey) ||
    isOfficeShortcutBlocked(event.target) ||
    isSpreadsheetNativeTextUndoTarget(event.target) ||
    !canExecute() ||
    !execute()
  ) {
    return false;
  }

  event.preventDefault();
  event.stopPropagation();
  return true;
}

export function runSpreadsheetPasteSpecialShortcut(
  event: KeyboardEvent,
  canExecute: SpreadsheetClipboardShortcutOptions['canExecute'],
  execute: SpreadsheetClipboardShortcutOptions['execute'],
): boolean {
  if (
    event.defaultPrevented ||
    event.repeat ||
    !event.altKey ||
    event.shiftKey ||
    !(event.metaKey || event.ctrlKey) ||
    isOfficeShortcutBlocked(event.target) ||
    isSpreadsheetNativeTextUndoTarget(event.target) ||
    !canExecute() ||
    !execute()
  ) {
    return false;
  }

  event.preventDefault();
  event.stopPropagation();
  return true;
}
