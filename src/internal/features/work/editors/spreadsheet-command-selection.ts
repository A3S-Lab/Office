import type {
  SpreadsheetCommandContext,
  SpreadsheetCommandRange,
} from './spreadsheet-command-controller';
import { spreadsheetSingleRange } from './spreadsheet-editor-support';

export function canEditSpreadsheetSelection(
  context: SpreadsheetCommandContext,
): boolean {
  return Boolean(context.editable && context.workbook && context.targetSheetId);
}

export function spreadsheetLiveCommandRange(
  context: SpreadsheetCommandContext,
): SpreadsheetCommandRange {
  const selection = context.workbook?.getSelection()?.at(-1);
  return spreadsheetSingleRange(selection ?? context.fallbackRange);
}
