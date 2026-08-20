import {
  applySpreadsheetCellStyle,
  canApplySpreadsheetCellStyle,
  type SpreadsheetCellStyleChoice,
} from './spreadsheet-cell-style';
import type {
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import { spreadsheetSingleRange } from './spreadsheet-editor-support';
import {
  createOfficeEditorExtension,
  type OfficeEditorExtension,
} from './office-editor-extension';

export function createSpreadsheetCellStyleExtension(): OfficeEditorExtension<
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands
> {
  return createOfficeEditorExtension<
    SpreadsheetCommandContext,
    SpreadsheetEditorCommands
  >({
    name: 'spreadsheetCellStyles',
    addCommands: () => ({
      applyCellStyle: {
        canExecute: canApplySelectedCellStyle,
        execute: applySelectedCellStyle,
      },
    }),
  });
}

function canApplySelectedCellStyle(
  context: SpreadsheetCommandContext,
  preset: SpreadsheetCellStyleChoice,
): boolean {
  return Boolean(
    context.editable &&
      context.workbook &&
      context.targetSheetId &&
      context.targetSheetId === context.activeSheetId &&
      canApplySpreadsheetCellStyle(
        context.content,
        context.targetSheetId,
        liveSpreadsheetCellStyleRange(context),
        preset,
      ),
  );
}

function applySelectedCellStyle(
  context: SpreadsheetCommandContext,
  preset: SpreadsheetCellStyleChoice,
): boolean {
  if (!canApplySelectedCellStyle(context, preset)) return false;
  const next = applySpreadsheetCellStyle(
    context.content,
    context.targetSheetId,
    liveSpreadsheetCellStyleRange(context),
    preset,
  );
  if (!next) return false;
  context.onChange(next);
  return true;
}

function liveSpreadsheetCellStyleRange(context: SpreadsheetCommandContext) {
  return spreadsheetSingleRange(
    context.workbook?.getSelection()?.at(-1) ?? context.fallbackRange,
  );
}
