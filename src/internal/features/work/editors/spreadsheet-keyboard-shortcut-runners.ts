import type { Cell } from '@fortune-sheet/core';
import { spreadsheetUnderlineStyle } from '../work-spreadsheet-underline';
import { isOfficeShortcutBlocked } from './office-shortcuts';
import type { SpreadsheetCellFillDirection } from './spreadsheet-cell-fill';
import type {
  SpreadsheetCommandContext,
  SpreadsheetEditorCanCommands,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import { isSpreadsheetNativeTextUndoTarget } from './spreadsheet-editor-support';
import { isSpreadsheetGridKeyboardTarget } from './spreadsheet-keyboard-navigation';
import {
  type SpreadsheetNumberFormatChoice,
  spreadsheetNumberFormatCode,
  spreadsheetNumberFormatValue,
} from './spreadsheet-number-format';
import {
  adjacentSpreadsheetSheetId,
  type SpreadsheetSheetMoveDirection,
} from './spreadsheet-sheet-model';

export function runSpreadsheetHistoryShortcut(
  event: KeyboardEvent,
  canExecute: () => boolean,
  execute: () => boolean,
): boolean {
  if (
    event.repeat ||
    isOfficeShortcutBlocked(event.target) ||
    isSpreadsheetNativeTextUndoTarget(event.target) ||
    !canExecute()
  ) {
    return false;
  }
  return execute();
}

export function runSpreadsheetCellFormatShortcut(
  event: KeyboardEvent,
  context: SpreadsheetCommandContext,
  canExecute: SpreadsheetEditorCanCommands['setCellFormat'],
  execute: SpreadsheetEditorCommands['setCellFormat'],
  attribute: Extract<keyof Cell, 'bl' | 'cl' | 'it' | 'un'>,
): boolean {
  if (
    event.repeat ||
    isOfficeShortcutBlocked(event.target) ||
    isSpreadsheetNativeTextUndoTarget(event.target)
  ) {
    return false;
  }
  const active =
    attribute === 'un'
      ? spreadsheetUnderlineStyle(context.toolbarCell?.un) !== 'none'
      : Number(context.toolbarCell?.[attribute]) === 1;
  const value = active ? 0 : 1;
  return canExecute(attribute, value) && execute(attribute, value);
}

export function runSpreadsheetNumberFormatShortcut(
  event: KeyboardEvent,
  context: SpreadsheetCommandContext,
  canExecute: SpreadsheetEditorCanCommands['setCellFormat'],
  execute: SpreadsheetEditorCommands['setCellFormat'],
  preset: SpreadsheetNumberFormatChoice,
): boolean {
  if (
    event.repeat ||
    isOfficeShortcutBlocked(event.target) ||
    isSpreadsheetNativeTextUndoTarget(event.target)
  ) {
    return false;
  }
  const value = spreadsheetNumberFormatValue(
    spreadsheetNumberFormatCode(preset),
    context.toolbarCell,
  );
  return canExecute('ct', value) && execute('ct', value);
}

export function runSpreadsheetMergeShortcut(
  event: KeyboardEvent,
  canExecute: SpreadsheetEditorCanCommands['mergeSelectedCells'],
  execute: SpreadsheetEditorCommands['mergeSelectedCells'],
): boolean {
  if (
    event.repeat ||
    isOfficeShortcutBlocked(event.target) ||
    isSpreadsheetNativeTextUndoTarget(event.target) ||
    !canExecute('merge-and-center')
  ) {
    return false;
  }
  return execute('merge-and-center');
}

export function runSpreadsheetClearShortcut(
  event: KeyboardEvent,
  canExecute: SpreadsheetEditorCanCommands['clearSelectedCells'],
  execute: SpreadsheetEditorCommands['clearSelectedCells'],
): boolean {
  if (
    event.repeat ||
    isOfficeShortcutBlocked(event.target) ||
    isSpreadsheetNativeTextUndoTarget(event.target) ||
    !isSpreadsheetGridKeyboardTarget(event.target) ||
    !canExecute()
  ) {
    return false;
  }
  return execute();
}

export function runSpreadsheetFillShortcut(
  event: KeyboardEvent,
  canExecute: SpreadsheetEditorCanCommands['fillSelectedCells'],
  execute: SpreadsheetEditorCommands['fillSelectedCells'],
  direction: SpreadsheetCellFillDirection,
): boolean {
  if (
    event.isComposing ||
    isOfficeShortcutBlocked(event.target) ||
    isSpreadsheetNativeTextUndoTarget(event.target) ||
    !isSpreadsheetGridKeyboardTarget(event.target)
  ) {
    return false;
  }
  if (!event.repeat && canExecute(direction)) execute(direction);
  return true;
}

export function runSpreadsheetAutoFilterShortcut(
  event: KeyboardEvent,
  canExecute: () => boolean,
  execute: () => boolean,
): boolean {
  if (
    event.repeat ||
    isOfficeShortcutBlocked(event.target) ||
    isSpreadsheetNativeTextUndoTarget(event.target) ||
    isSpreadsheetAutoFilterTextTarget(event.target) ||
    !isSpreadsheetGridKeyboardTarget(event.target) ||
    !canExecute()
  ) {
    return false;
  }
  return execute();
}

export function runSpreadsheetFormatPainterEscape(
  event: KeyboardEvent,
  canExecute: SpreadsheetEditorCanCommands['cancelFormatPainter'],
  execute: SpreadsheetEditorCommands['cancelFormatPainter'],
): boolean {
  if (
    event.repeat ||
    isOfficeShortcutBlocked(event.target) ||
    isSpreadsheetNativeTextUndoTarget(event.target) ||
    !canExecute()
  ) {
    return false;
  }
  return execute();
}

export function runSpreadsheetSheetNavigationShortcut(
  event: KeyboardEvent,
  context: SpreadsheetCommandContext,
  activateSheet: SpreadsheetEditorCommands['activateSheet'],
  direction: SpreadsheetSheetMoveDirection,
): boolean {
  if (
    event.repeat ||
    isOfficeShortcutBlocked(event.target) ||
    isSpreadsheetNativeTextUndoTarget(event.target)
  ) {
    return false;
  }
  const target = adjacentSpreadsheetSheetId(
    context.content,
    context.activeSheetId,
    direction,
  );
  return Boolean(
    target && target !== context.activeSheetId && activateSheet(target),
  );
}

export function runSpreadsheetAddSheetShortcut(
  event: KeyboardEvent,
  canExecute: SpreadsheetEditorCanCommands['addSheet'],
  execute: SpreadsheetEditorCommands['addSheet'],
): boolean {
  if (
    event.repeat ||
    isOfficeShortcutBlocked(event.target) ||
    isSpreadsheetNativeTextUndoTarget(event.target) ||
    !canExecute()
  ) {
    return false;
  }
  return execute();
}

export function runSpreadsheetRecalculationShortcut(
  event: KeyboardEvent,
  canExecute: SpreadsheetEditorCanCommands['recalculateFormula'],
  execute: SpreadsheetEditorCommands['recalculateFormula'],
): boolean {
  if (
    event.repeat ||
    isOfficeShortcutBlocked(event.target) ||
    isSpreadsheetNativeTextUndoTarget(event.target) ||
    !canExecute('workbook')
  ) {
    return false;
  }
  return execute('workbook');
}

function isSpreadsheetAutoFilterTextTarget(
  target: EventTarget | null,
): boolean {
  return (
    target instanceof HTMLInputElement ||
    target instanceof HTMLTextAreaElement ||
    target instanceof HTMLSelectElement ||
    (target instanceof HTMLElement &&
      (target.isContentEditable ||
        Boolean(target.closest('[contenteditable="true"]'))))
  );
}
