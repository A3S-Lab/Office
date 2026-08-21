import type { Cell } from '@fortune-sheet/core';
import {
  isSpreadsheetTextOrientationId,
  spreadsheetTextOrientationCellStyle,
  spreadsheetTextOrientationFromChoice,
  type SpreadsheetTextOrientationId,
} from '../work-spreadsheet-text-orientation';
import {
  createOfficeEditorExtension,
  type OfficeEditorExtension,
} from './office-editor-extension';
import {
  normalizeSpreadsheetCellRange,
  spreadsheetCellRangeArea,
} from './spreadsheet-cell-range';
import type {
  SpreadsheetCommandContext,
  SpreadsheetCommandRange,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import { spreadsheetSingleRange } from './spreadsheet-editor-support';

export const MAX_SPREADSHEET_TEXT_ORIENTATION_CELLS = 10_000;

export interface SpreadsheetTextOrientationApiCall {
  name: 'setCellFormatByRange';
  args: [
    'rt' | 'tr',
    Cell['rt'] | Cell['tr'] | undefined,
    SpreadsheetCommandRange,
    { id: string },
  ];
}

export function createSpreadsheetTextOrientationExtension(): OfficeEditorExtension<
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands
> {
  return createOfficeEditorExtension<
    SpreadsheetCommandContext,
    SpreadsheetEditorCommands
  >({
    name: 'spreadsheetTextOrientation',
    addCommands: () => ({
      setTextOrientation: {
        canExecute: canSetSpreadsheetTextOrientation,
        execute: setSpreadsheetTextOrientation,
      },
    }),
  });
}

export function spreadsheetTextOrientationApiCalls(
  range: SpreadsheetCommandRange,
  sheetId: string,
  choice: SpreadsheetTextOrientationId,
): SpreadsheetTextOrientationApiCall[] {
  if (!sheetId || !isSpreadsheetTextOrientationId(choice)) return [];
  const style = spreadsheetTextOrientationCellStyle(
    spreadsheetTextOrientationFromChoice(choice),
  );
  const target = spreadsheetSingleRange(range);
  const options = { id: sheetId };
  return style?.rt !== undefined
    ? [
        {
          name: 'setCellFormatByRange',
          args: ['tr', undefined, target, options],
        },
        {
          name: 'setCellFormatByRange',
          args: ['rt', style.rt, target, options],
        },
      ]
    : [
        {
          name: 'setCellFormatByRange',
          args: ['rt', undefined, target, options],
        },
        {
          name: 'setCellFormatByRange',
          args: ['tr', style?.tr, target, options],
        },
      ];
}

function canSetSpreadsheetTextOrientation(
  context: SpreadsheetCommandContext,
  choice: SpreadsheetTextOrientationId,
): boolean {
  if (
    !context.editable ||
    !context.workbook ||
    !context.targetSheetId ||
    !context.content.sheets.some(
      (sheet) => sheet.id === context.targetSheetId,
    ) ||
    !isSpreadsheetTextOrientationId(choice)
  ) {
    return false;
  }
  const range = normalizeSpreadsheetCellRange(
    spreadsheetSingleRange(
      context.workbook.getSelection()?.at(-1) ?? context.fallbackRange,
    ),
  );
  return Boolean(
    range &&
      spreadsheetCellRangeArea(range) <= MAX_SPREADSHEET_TEXT_ORIENTATION_CELLS,
  );
}

function setSpreadsheetTextOrientation(
  context: SpreadsheetCommandContext,
  choice: SpreadsheetTextOrientationId,
): boolean {
  if (!context.workbook || !canSetSpreadsheetTextOrientation(context, choice)) {
    return false;
  }
  const range = spreadsheetSingleRange(
    context.workbook.getSelection()?.at(-1) ?? context.fallbackRange,
  );
  const calls = spreadsheetTextOrientationApiCalls(
    range,
    context.targetSheetId,
    choice,
  );
  if (!calls.length) return false;
  try {
    context.workbook.batchCallApis(calls);
    return true;
  } catch {
    return false;
  }
}
