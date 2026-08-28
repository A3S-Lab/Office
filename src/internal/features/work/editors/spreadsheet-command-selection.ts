import type {
  SpreadsheetCommandContext,
  SpreadsheetCommandSelection,
  SpreadsheetCommandRange,
  SpreadsheetSelectionRef,
} from './spreadsheet-command-controller';
import {
  finiteSpreadsheetSelection,
  spreadsheetSingleRange,
} from './spreadsheet-editor-support';

export function canEditSpreadsheetSelection(
  context: SpreadsheetCommandContext,
): boolean {
  return Boolean(context.editable && context.workbook && context.targetSheetId);
}

export function spreadsheetLiveCommandRange(
  context: SpreadsheetCommandContext,
): SpreadsheetCommandRange {
  const selection = spreadsheetLiveCommandSelection(context);
  return spreadsheetSingleRange(selection ?? context.fallbackRange);
}

/**
 * Returns the most recent selection known to the Office bridge. The ref is
 * intentionally checked before Fortune's imperative API because that API is
 * backed by React state and can lag behind a just-issued keyboard command.
 */
export function spreadsheetLiveCommandSelection(
  context: SpreadsheetCommandContext,
): SpreadsheetCommandRange | undefined {
  return spreadsheetLiveCommandSelections(context)?.at(-1);
}

/**
 * Returns all current selections while preserving the workbook's multi-range
 * semantics. A synchronous command selection is represented as one range.
 */
export function spreadsheetLiveCommandSelections(
  context: SpreadsheetCommandContext,
): SpreadsheetCommandRange[] | undefined {
  const shadow = context.selectionRef?.current;
  if (shadow?.sheetId === context.targetSheetId) return [shadow.selection];
  return context.workbook?.getSelection();
}

/**
 * Accept a Fortune selection callback unless it is older than a pending
 * programmatic selection request. The returned value is the normalized state
 * that should be published to React and collaboration consumers.
 */
export function acceptSpreadsheetSelectionChange(
  selectionRef: SpreadsheetSelectionRef,
  sheetId: string,
  selection: SpreadsheetCommandRange,
): SpreadsheetCommandSelection | null {
  const nextSelection = finiteSpreadsheetSelection(selection);
  const requested = selectionRef.requested;
  if (
    requested &&
    (requested.sheetId !== sheetId ||
      !sameSpreadsheetSelectionValue(requested.selection, nextSelection))
  ) {
    return null;
  }
  const next = { sheetId, selection: nextSelection };
  selectionRef.current = next;
  return next;
}

export function rememberSpreadsheetCommandSelection(
  context: SpreadsheetCommandContext,
  selection: SpreadsheetCommandRange,
): void {
  if (!context.selectionRef || !context.targetSheetId) return;
  const next = {
    ...selection,
    row: [...selection.row],
    column: [...selection.column],
  };
  context.selectionRef.current = {
    sheetId: context.targetSheetId,
    selection: next,
  };
  context.selectionRef.requested = {
    sheetId: context.targetSheetId,
    selection: { ...next, row: [...next.row], column: [...next.column] },
  };
}

/** Clear a pending programmatic request when a fresh user intent arrives. */
export function releaseSpreadsheetSelectionRequest(
  selectionRef: SpreadsheetSelectionRef,
): void {
  selectionRef.requested = null;
}

export function sameSpreadsheetSelectionValue(
  left: SpreadsheetCommandRange,
  right: SpreadsheetCommandRange,
): boolean {
  const normalizedLeft = finiteSpreadsheetSelection(left);
  const normalizedRight = finiteSpreadsheetSelection(right);
  return (
    normalizedLeft.row[0] === normalizedRight.row[0] &&
    normalizedLeft.row[1] === normalizedRight.row[1] &&
    normalizedLeft.column[0] === normalizedRight.column[0] &&
    normalizedLeft.column[1] === normalizedRight.column[1] &&
    normalizedLeft.row_focus === normalizedRight.row_focus &&
    normalizedLeft.column_focus === normalizedRight.column_focus
  );
}
