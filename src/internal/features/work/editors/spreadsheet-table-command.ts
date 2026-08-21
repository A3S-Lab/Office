import { isOfficeShortcutBlocked } from './office-shortcuts';
import type {
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import { spreadsheetCommandCatalog } from './spreadsheet-command-catalog';
import {
  finiteSpreadsheetSelection,
  isSpreadsheetNativeTextUndoTarget,
} from './spreadsheet-editor-support';
import { isSpreadsheetGridKeyboardTarget } from './spreadsheet-keyboard-navigation';
import {
  applySpreadsheetTable,
  convertSpreadsheetTableToRange,
  type SpreadsheetTableDesignPatch,
  type SpreadsheetTableRequest,
  type SpreadsheetTableTarget,
  spreadsheetTableAtCell,
  updateSpreadsheetTable,
  validateSpreadsheetTableRequest,
} from './spreadsheet-table';
import { canMaterializeSpreadsheetTableAppearance } from './spreadsheet-table-conversion';
import {
  createOfficeEditorExtension,
  type OfficeEditorExtension,
} from './office-editor-extension';

export function createSpreadsheetTableExtension(): OfficeEditorExtension<
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands
> {
  return createOfficeEditorExtension<
    SpreadsheetCommandContext,
    SpreadsheetEditorCommands
  >({
    name: 'spreadsheetTables',
    addCommands: () => ({
      openTable: {
        canExecute: canOpenSpreadsheetTable,
        execute: openSpreadsheetTable,
      },
      applyTable: {
        canExecute: (context, request) =>
          context.editable &&
          validateSpreadsheetTableRequest(context.content, request).ok,
        execute: applySpreadsheetTableCommand,
      },
      updateTable: {
        canExecute: canUpdateSpreadsheetTable,
        execute: updateSpreadsheetTableCommand,
      },
      convertTableToRange: {
        canExecute: canConvertSpreadsheetTableToRange,
        execute: convertSpreadsheetTableToRangeCommand,
      },
    }),
    addKeyboardShortcuts: () => ({
      [spreadsheetCommandCatalog.table.shortcut.editor[0]]: (
        { can, commands },
        event,
      ) =>
        runSpreadsheetTableShortcut(event, can.openTable, commands.openTable),
    }),
  });
}

function canOpenSpreadsheetTable(context: SpreadsheetCommandContext): boolean {
  const target = liveSpreadsheetTableTarget(context);
  const sheet = context.content.sheets.find(
    (candidate) => candidate.id === target?.sheetId,
  );
  const focus = target?.selection;
  const row = focus?.row_focus ?? focus?.row[0] ?? 0;
  const column = focus?.column_focus ?? focus?.column[0] ?? 0;
  return Boolean(
    context.editable &&
      context.table.canOpen &&
      context.workbook &&
      target &&
      target.sheetId === context.activeSheetId &&
      sheet &&
      !sheet.isPivotTable &&
      !sheet.pivotTables?.length &&
      !spreadsheetTableAtCell(sheet, row, column),
  );
}

function openSpreadsheetTable(context: SpreadsheetCommandContext): boolean {
  const target = liveSpreadsheetTableTarget(context);
  return Boolean(
    target && canOpenSpreadsheetTable(context) && context.table.open(target),
  );
}

function applySpreadsheetTableCommand(
  context: SpreadsheetCommandContext,
  request: SpreadsheetTableRequest,
): boolean {
  if (!context.editable) return false;
  const next = applySpreadsheetTable(context.content, request);
  if (!next) return false;
  context.onChange(next);
  return true;
}

function canUpdateSpreadsheetTable(
  context: SpreadsheetCommandContext,
  sheetId: string,
  tableId: string,
  patch: SpreadsheetTableDesignPatch,
): boolean {
  return Boolean(
    context.editable &&
      updateSpreadsheetTable(context.content, sheetId, tableId, patch),
  );
}

function updateSpreadsheetTableCommand(
  context: SpreadsheetCommandContext,
  sheetId: string,
  tableId: string,
  patch: SpreadsheetTableDesignPatch,
): boolean {
  if (!context.editable) return false;
  const next = updateSpreadsheetTable(context.content, sheetId, tableId, patch);
  if (!next) return false;
  context.onChange(next);
  return true;
}

function canConvertSpreadsheetTableToRange(
  context: SpreadsheetCommandContext,
  sheetId: string,
  tableId: string,
): boolean {
  return Boolean(
    context.editable &&
      context.content.sheets
        .find((sheet) => sheet.id === sheetId)
        ?.tables?.some(
          (table) =>
            table.id === tableId &&
            canMaterializeSpreadsheetTableAppearance(table),
        ),
  );
}

function convertSpreadsheetTableToRangeCommand(
  context: SpreadsheetCommandContext,
  sheetId: string,
  tableId: string,
): boolean {
  if (!context.editable) return false;
  const next = convertSpreadsheetTableToRange(
    context.content,
    sheetId,
    tableId,
  );
  if (!next) return false;
  context.onChange(next);
  return true;
}

function liveSpreadsheetTableTarget(
  context: SpreadsheetCommandContext,
): SpreadsheetTableTarget | null {
  if (!context.workbook || !context.targetSheetId) return null;
  const selection =
    context.workbook.getSelection()?.at(-1) ?? context.fallbackRange;
  return {
    sheetId: context.targetSheetId,
    selection: finiteSpreadsheetSelection(selection),
  };
}

function runSpreadsheetTableShortcut(
  event: KeyboardEvent,
  canExecute: SpreadsheetEditorCommands['openTable'],
  execute: SpreadsheetEditorCommands['openTable'],
): boolean {
  if (
    event.repeat ||
    event.isComposing ||
    isOfficeShortcutBlocked(event.target) ||
    isSpreadsheetNativeTextUndoTarget(event.target) ||
    !isSpreadsheetGridKeyboardTarget(event.target)
  ) {
    return false;
  }
  if (!event.repeat && canExecute()) execute();
  return true;
}
