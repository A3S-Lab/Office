import type { CellMatrix } from '@fortune-sheet/core';
import { expect, test } from '@rstest/core';
import { synchronizeSpreadsheetWorkbookInPlace } from '../src/internal/features/work/editors/spreadsheet-in-place-workbook-sync';
import { registerImportedSpreadsheetMatrix } from '../src/internal/features/work/work-spreadsheet-matrix-profile';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

test('adopts a certified simple workbook without remounting', () => {
  const previous = workbook('sheet-1', [[{ v: 'Loading' }]]);
  const next = certifiedWorkbook('sheet-1', [[{ v: 'Ready' }]]);
  const updates: WorkSpreadsheetContent['sheets'][] = [];

  const synchronized = synchronizeSpreadsheetWorkbookInPlace(
    { updateSheet: (sheets) => updates.push(sheets) },
    previous,
    next,
    next.sheets,
    false,
  );

  expect(synchronized).toBe(true);
  expect(updates).toEqual([next.sheets]);
});

test('keeps the remount path for previews and identity changes', () => {
  const previous = workbook('sheet-1', [[{ v: 'Loading' }]]);
  const next = certifiedWorkbook('sheet-1', [[{ v: 'Ready' }]]);
  const updateSheet = () => {
    throw new Error('must not update');
  };

  expect(
    synchronizeSpreadsheetWorkbookInPlace(
      { updateSheet },
      previous,
      next,
      next.sheets,
      true,
    ),
  ).toBe(false);
  expect(
    synchronizeSpreadsheetWorkbookInPlace(
      { updateSheet },
      previous,
      certifiedWorkbook('sheet-2', [[{ v: 'Ready' }]]),
      next.sheets,
      false,
    ),
  ).toBe(false);
});

test('keeps the remount path for uncertified or stateful worksheets', () => {
  const previous = workbook('sheet-1', [[{ v: 'Loading' }]]);
  const uncertified = workbook('sheet-1', [[{ v: 'Ready' }]]);
  const protectedWorkbook = certifiedWorkbook('sheet-1', [[{ v: 'Ready' }]]);
  protectedWorkbook.sheets[0].config = { authority: { sheet: 1 } };
  const updateSheet = () => {
    throw new Error('must not update');
  };

  expect(
    synchronizeSpreadsheetWorkbookInPlace(
      { updateSheet },
      previous,
      uncertified,
      uncertified.sheets,
      false,
    ),
  ).toBe(false);
  expect(
    synchronizeSpreadsheetWorkbookInPlace(
      { updateSheet },
      previous,
      protectedWorkbook,
      protectedWorkbook.sheets,
      false,
    ),
  ).toBe(false);
});

test('falls back to remounting if the Fortune update rejects the workbook', () => {
  const previous = workbook('sheet-1', [[{ v: 'Loading' }]]);
  const next = certifiedWorkbook('sheet-1', [[{ v: 'Ready' }]]);

  expect(
    synchronizeSpreadsheetWorkbookInPlace(
      {
        updateSheet: () => {
          throw new Error('rejected');
        },
      },
      previous,
      next,
      next.sheets,
      false,
    ),
  ).toBe(false);
});

function certifiedWorkbook(
  sheetId: string,
  data: CellMatrix,
): WorkSpreadsheetContent {
  registerImportedSpreadsheetMatrix(data, {
    columnCount: data[0]?.length ?? 0,
    formulaCells: [],
    fortuneReady: true,
    populatedCellCount: 1,
    protectionCellKey: '',
    rowCount: data.length,
    shownCommentCells: [],
  });
  return workbook(sheetId, data);
}

function workbook(sheetId: string, data: CellMatrix): WorkSpreadsheetContent {
  return {
    type: 'spreadsheet',
    sheets: [
      {
        id: sheetId,
        name: 'Sheet 1',
        status: 1,
        order: 0,
        row: data.length,
        column: data[0]?.length ?? 0,
        data,
      },
    ],
  };
}
