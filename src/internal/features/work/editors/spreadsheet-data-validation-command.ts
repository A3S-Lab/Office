import type {
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import {
  applySpreadsheetDataValidation,
  canRemoveSpreadsheetDataValidation,
  createSpreadsheetDataValidationDialogSource,
  removeSpreadsheetDataValidation,
  type SpreadsheetDataValidationRequest,
  type SpreadsheetDataValidationTarget,
  validateSpreadsheetDataValidationRequest,
} from './spreadsheet-data-validation';
import { normalizeSpreadsheetCellRange } from './spreadsheet-cell-range';
import {
  createOfficeEditorExtension,
  type OfficeEditorExtension,
} from './office-editor-extension';

export function createSpreadsheetDataValidationExtension(): OfficeEditorExtension<
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands
> {
  return createOfficeEditorExtension<
    SpreadsheetCommandContext,
    SpreadsheetEditorCommands
  >({
    name: 'spreadsheetDataValidation',
    addCommands: () => ({
      openDataValidation: {
        canExecute: canOpenSpreadsheetDataValidation,
        execute: openSpreadsheetDataValidation,
      },
      applyDataValidation: {
        canExecute: (context, request) =>
          context.editable &&
          validateSpreadsheetDataValidationRequest(context.content, request).ok,
        execute: applySpreadsheetDataValidationCommand,
      },
      removeDataValidation: {
        canExecute: (context, target) =>
          context.editable &&
          canRemoveSpreadsheetDataValidation(context.content, target),
        execute: removeSpreadsheetDataValidationCommand,
      },
    }),
  });
}

function canOpenSpreadsheetDataValidation(
  context: SpreadsheetCommandContext,
): boolean {
  const target = liveSpreadsheetDataValidationTarget(context);
  return Boolean(
    context.editable &&
      context.dataValidation.canOpen &&
      target &&
      target.sheetId === context.activeSheetId &&
      createSpreadsheetDataValidationDialogSource(context.content, target),
  );
}

function openSpreadsheetDataValidation(
  context: SpreadsheetCommandContext,
): boolean {
  const target = liveSpreadsheetDataValidationTarget(context);
  return Boolean(
    target &&
      canOpenSpreadsheetDataValidation(context) &&
      context.dataValidation.open(target),
  );
}

function applySpreadsheetDataValidationCommand(
  context: SpreadsheetCommandContext,
  request: SpreadsheetDataValidationRequest,
): boolean {
  if (!context.editable) return false;
  const next = applySpreadsheetDataValidation(context.content, request);
  if (!next) return false;
  context.onChange(next);
  return true;
}

function removeSpreadsheetDataValidationCommand(
  context: SpreadsheetCommandContext,
  target: SpreadsheetDataValidationTarget,
): boolean {
  if (!context.editable) return false;
  const next = removeSpreadsheetDataValidation(context.content, target);
  if (!next) return false;
  context.onChange(next);
  return true;
}

function liveSpreadsheetDataValidationTarget(
  context: SpreadsheetCommandContext,
): SpreadsheetDataValidationTarget | null {
  if (!context.workbook || !context.targetSheetId) return null;
  const selections = context.workbook.getSelection();
  const sourceSelections = selections?.length
    ? selections
    : [context.fallbackRange];
  const ranges = sourceSelections.flatMap((selection) => {
    const range = normalizeSpreadsheetCellRange(selection);
    return range ? [range] : [];
  });
  const focusedSelection = sourceSelections.at(-1);
  const focusedRange = focusedSelection
    ? normalizeSpreadsheetCellRange(focusedSelection)
    : null;
  if (!ranges.length || !focusedSelection || !focusedRange) return null;
  return {
    sheetId: context.targetSheetId,
    ranges,
    activeCell: {
      row: clampSpreadsheetDataValidationFocus(
        focusedSelection.row_focus,
        focusedRange.row,
      ),
      column: clampSpreadsheetDataValidationFocus(
        focusedSelection.column_focus,
        focusedRange.column,
      ),
    },
  };
}

function clampSpreadsheetDataValidationFocus(
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
