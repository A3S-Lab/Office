import {
  canSetSpreadsheetCellBorders,
  type SpreadsheetCellBorderFormat,
  setSpreadsheetCellBorders,
} from './spreadsheet-cell-border';
import type {
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import { spreadsheetSingleRange } from './spreadsheet-editor-support';
import {
  createOfficeEditorExtension,
  type OfficeEditorExtension,
} from './office-editor-extension';

export function createSpreadsheetCellBorderExtension(): OfficeEditorExtension<
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands
> {
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
  if (!canSetSelectedCellBorders(context, format)) return false;
  const next = setSpreadsheetCellBorders(
    context.content,
    context.targetSheetId,
    liveSpreadsheetBorderRange(context),
    format,
  );
  if (!next) return false;
  context.onChange(next);
  return true;
}

function liveSpreadsheetBorderRange(context: SpreadsheetCommandContext) {
  return spreadsheetSingleRange(
    context.workbook?.getSelection()?.at(-1) ?? context.fallbackRange,
  );
}
