import {
  parseSpreadsheetClipboardText,
  type SpreadsheetClipboardPort,
} from './spreadsheet-context-menu';
import { isSpreadsheetNativeTextUndoTarget } from './spreadsheet-editor-support';
import { isOfficeShortcutBlocked } from './office-shortcuts';

export interface SpreadsheetClipboardShortcutOptions {
  clipboard: SpreadsheetClipboardPort;
  clearSelectedCells(): boolean;
  pasteCells(values: readonly (readonly unknown[])[]): boolean;
  readSelectionText(): string | null;
  restoreFocus(): void;
}

export function runSpreadsheetClipboardShortcut(
  event: KeyboardEvent,
  options: SpreadsheetClipboardShortcutOptions,
): boolean {
  const key = event.key.toLocaleLowerCase();
  if (
    event.defaultPrevented ||
    event.repeat ||
    event.altKey ||
    event.shiftKey ||
    !(event.metaKey || event.ctrlKey) ||
    !clipboardShortcutKeys.has(key) ||
    isOfficeShortcutBlocked(event.target) ||
    isSpreadsheetNativeTextUndoTarget(event.target)
  ) {
    return false;
  }

  const selectionText = key === 'v' ? null : options.readSelectionText();
  if (key !== 'v' && selectionText === null) return false;
  event.preventDefault();
  event.stopPropagation();

  if (key === 'v') {
    void options.clipboard
      .readText()
      .then((value) => {
        const cells = parseSpreadsheetClipboardText(value);
        if (cells.length) options.pasteCells(cells);
      })
      .catch(() => undefined)
      .finally(options.restoreFocus);
    return true;
  }

  void options.clipboard.writeText(selectionText ?? '').catch(() => undefined);
  if (key === 'x') options.clearSelectedCells();
  options.restoreFocus();
  return true;
}

const clipboardShortcutKeys = new Set(['c', 'v', 'x']);
