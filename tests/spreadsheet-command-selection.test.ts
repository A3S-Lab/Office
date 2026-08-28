import { expect, test } from '@rstest/core';
import {
  acceptSpreadsheetSelectionChange,
  rememberSpreadsheetCommandSelection,
  releaseSpreadsheetSelectionRequest,
  spreadsheetLiveCommandSelection,
} from '../src/internal/features/work/editors/spreadsheet-command-selection';
import type {
  SpreadsheetCommandContext,
  SpreadsheetCommandRange,
  SpreadsheetSelectionRef,
} from '../src/internal/features/work/editors/spreadsheet-command-controller';
import { moveSpreadsheetKeyboardSelection } from '../src/internal/features/work/editors/spreadsheet-keyboard-navigation';
import type { WorkSpreadsheetSheet } from '../src/internal/features/work/work-types';

test('keeps consecutive keyboard navigation on the synchronous selection shadow', () => {
  const initial: SpreadsheetCommandRange = {
    row: [0, 2],
    column: [1, 1],
    row_focus: 0,
    column_focus: 1,
  };
  const selectionRef: SpreadsheetSelectionRef = {
    current: null,
    requested: null,
  };
  const context = selectionContext(selectionRef, initial);
  const sheet: WorkSpreadsheetSheet = {
    id: 'sheet-1',
    name: 'Sheet 1',
    row: 20,
    column: 8,
    data: [],
  };

  rememberSpreadsheetCommandSelection(context, initial);
  let live = spreadsheetLiveCommandSelection(context);
  expect(live).toEqual(initial);

  live = moveSpreadsheetKeyboardSelection(
    sheet,
    live ?? initial,
    'sheet-start',
    false,
  );
  rememberSpreadsheetCommandSelection(context, live);
  live = moveSpreadsheetKeyboardSelection(sheet, live, 'right', false);
  rememberSpreadsheetCommandSelection(context, live);
  live = moveSpreadsheetKeyboardSelection(sheet, live, 'down', false);

  expect(live).toEqual({
    row: [1, 1],
    column: [1, 1],
    row_focus: 1,
    column_focus: 1,
  });
  expect(spreadsheetLiveCommandSelection(context)).toEqual({
    row: [0, 0],
    column: [1, 1],
    row_focus: 0,
    column_focus: 1,
  });
});

test('ignores a delayed callback that predates the requested selection', () => {
  const requested: SpreadsheetCommandRange = {
    row: [4, 4],
    column: [2, 2],
    row_focus: 4,
    column_focus: 2,
  };
  const stale: SpreadsheetCommandRange = {
    row: [3, 3],
    column: [2, 2],
    row_focus: 3,
    column_focus: 2,
  };
  const selectionRef: SpreadsheetSelectionRef = {
    current: null,
    requested: null,
  };

  rememberSpreadsheetCommandSelection(
    selectionContext(selectionRef, stale),
    requested,
  );

  expect(acceptSpreadsheetSelectionChange(selectionRef, 'sheet-1', stale)).toBe(
    null,
  );
  expect(selectionRef.current).toEqual({
    sheetId: 'sheet-1',
    selection: requested,
  });
  expect(
    acceptSpreadsheetSelectionChange(selectionRef, 'sheet-1', requested),
  ).toEqual(selectionRef.current);
});

test('allows a fresh pointer or native keyboard intent to replace the guard', () => {
  const selectionRef: SpreadsheetSelectionRef = {
    current: null,
    requested: null,
  };
  const requested: SpreadsheetCommandRange = {
    row: [4, 4],
    column: [2, 2],
    row_focus: 4,
    column_focus: 2,
  };
  const userSelection: SpreadsheetCommandRange = {
    row: [1, 1],
    column: [0, 0],
    row_focus: 1,
    column_focus: 0,
  };

  rememberSpreadsheetCommandSelection(
    selectionContext(selectionRef, userSelection),
    requested,
  );
  releaseSpreadsheetSelectionRequest(selectionRef);

  expect(
    acceptSpreadsheetSelectionChange(selectionRef, 'sheet-1', userSelection),
  ).toEqual({ sheetId: 'sheet-1', selection: userSelection });
  expect(selectionRef.current).toEqual({
    sheetId: 'sheet-1',
    selection: userSelection,
  });
});

function selectionContext(
  selectionRef: SpreadsheetSelectionRef,
  staleSelection: SpreadsheetCommandRange,
): SpreadsheetCommandContext {
  return {
    selectionRef,
    targetSheetId: 'sheet-1',
    workbook: {
      getSelection: () => [staleSelection],
    },
  } as SpreadsheetCommandContext;
}
