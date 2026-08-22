import { isOfficeShortcutBlocked } from './office-shortcuts';
import {
  applySpreadsheetCellFormat,
  canApplySpreadsheetCellFormat,
  MAX_SPREADSHEET_CELL_FORMAT_CELLS,
  type SpreadsheetCellFormatRequest,
} from './spreadsheet-cell-format';
import {
  normalizeSpreadsheetCellRange,
  spreadsheetCellRangeArea,
} from './spreadsheet-cell-range';
import { spreadsheetCommandCatalog } from './spreadsheet-command-catalog';
import type {
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands,
  SpreadsheetFormatCellsOpenRequest,
} from './spreadsheet-command-controller';
import {
  normalizeSpreadsheetFormatCellsOpenIntent,
  type SpreadsheetFormatCellsOpenIntent,
} from './spreadsheet-format-cells-intent';
import {
  isSpreadsheetNativeTextUndoTarget,
  spreadsheetSingleRange,
} from './spreadsheet-editor-support';
import {
  createOfficeEditorExtension,
  type OfficeEditorExtension,
} from './office-editor-extension';

export function createSpreadsheetCellFormatExtension(): OfficeEditorExtension<
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands
> {
  return createOfficeEditorExtension<
    SpreadsheetCommandContext,
    SpreadsheetEditorCommands
  >({
    name: 'spreadsheetCellFormat',
    addCommands: () => ({
      openFormatCells: {
        canExecute: canOpenSpreadsheetFormatCells,
        execute: openSpreadsheetFormatCells,
      },
      applyCellFormat: {
        canExecute: (context, request) =>
          context.editable &&
          canApplySpreadsheetCellFormat(context.content, request),
        execute: applySpreadsheetFormatCells,
      },
    }),
    addKeyboardShortcuts: () => ({
      [spreadsheetCommandCatalog.formatCells.shortcut.editor[0]]: (
        { can, commands },
        event,
      ) =>
        runSpreadsheetFormatCellsShortcut(
          event,
          undefined,
          can.openFormatCells,
          commands.openFormatCells,
        ),
      [spreadsheetCommandCatalog.formatCellsFont.shortcut.editor[0]]: (
        { can, commands },
        event,
      ) =>
        runSpreadsheetFormatCellsShortcut(
          event,
          { tab: 'font', focus: 'fontFamily' },
          can.openFormatCells,
          commands.openFormatCells,
        ),
      [spreadsheetCommandCatalog.formatCellsFontSize.shortcut.editor[0]]: (
        { can, commands },
        event,
      ) =>
        runSpreadsheetFormatCellsShortcut(
          event,
          { tab: 'font', focus: 'fontSize' },
          can.openFormatCells,
          commands.openFormatCells,
        ),
    }),
  });
}

function canOpenSpreadsheetFormatCells(
  context: SpreadsheetCommandContext,
  intent?: SpreadsheetFormatCellsOpenIntent,
): boolean {
  if (!normalizeSpreadsheetFormatCellsOpenIntent(intent)) return false;
  if (
    !context.editable ||
    !context.formatCells.canOpen ||
    !context.workbook ||
    !context.targetSheetId ||
    context.targetSheetId !== context.activeSheetId
  ) {
    return false;
  }
  const range = liveSpreadsheetFormatCellsRange(context);
  return Boolean(
    range &&
      spreadsheetCellRangeArea(range) <= MAX_SPREADSHEET_CELL_FORMAT_CELLS,
  );
}

function openSpreadsheetFormatCells(
  context: SpreadsheetCommandContext,
  intent?: SpreadsheetFormatCellsOpenIntent,
): boolean {
  const normalizedIntent = normalizeSpreadsheetFormatCellsOpenIntent(intent);
  if (
    !normalizedIntent ||
    !canOpenSpreadsheetFormatCells(context, normalizedIntent) ||
    !context.workbook
  ) {
    return false;
  }
  const selection =
    context.workbook.getSelection()?.at(-1) ?? context.fallbackRange;
  const range = normalizeSpreadsheetCellRange(
    spreadsheetSingleRange(selection),
  );
  if (!range) return false;
  const request: SpreadsheetFormatCellsOpenRequest = {
    sheetId: context.targetSheetId,
    range,
    activeCell: {
      row: clampFocus(selection.row_focus, range.row),
      column: clampFocus(selection.column_focus, range.column),
    },
    cells: [],
    intent: normalizedIntent,
  };
  try {
    request.cells = context.workbook.getCellsByRange(range, {
      id: context.targetSheetId,
    });
    return context.formatCells.open(request);
  } catch {
    return false;
  }
}

function applySpreadsheetFormatCells(
  context: SpreadsheetCommandContext,
  request: SpreadsheetCellFormatRequest,
): boolean {
  if (!context.editable) return false;
  const next = applySpreadsheetCellFormat(context.content, request);
  if (!next) return false;
  context.onChange(next);
  return true;
}

function liveSpreadsheetFormatCellsRange(context: SpreadsheetCommandContext) {
  return normalizeSpreadsheetCellRange(
    spreadsheetSingleRange(
      context.workbook?.getSelection()?.at(-1) ?? context.fallbackRange,
    ),
  );
}

function runSpreadsheetFormatCellsShortcut(
  event: KeyboardEvent,
  intent: SpreadsheetFormatCellsOpenIntent | undefined,
  canExecute: SpreadsheetEditorCommands['openFormatCells'],
  execute: SpreadsheetEditorCommands['openFormatCells'],
): boolean {
  if (
    event.repeat ||
    event.isComposing ||
    isOfficeShortcutBlocked(event.target) ||
    isSpreadsheetNativeTextUndoTarget(event.target) ||
    !canExecute(intent)
  ) {
    return false;
  }
  return execute(intent);
}

function clampFocus(value: unknown, range: readonly number[]): number {
  const minimum = range[0] ?? 0;
  const maximum = range[1] ?? minimum;
  const focus =
    typeof value === 'number' && Number.isFinite(value)
      ? Math.trunc(value)
      : minimum;
  return Math.min(maximum, Math.max(minimum, focus));
}
