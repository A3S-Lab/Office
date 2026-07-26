import { expect, test } from '@rstest/core';
import { spreadsheetSelectionSummary } from '../src/internal/features/work/editors/spreadsheet-editor-support';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

test('summarizes populated and numeric cells in a spreadsheet selection', () => {
  const sheet: WorkSpreadsheetContent['sheets'][number] = {
    id: 'sheet-1',
    name: 'Sheet 1',
    data: [
      [{ v: 'Owner' }, { v: 12 }, { v: 18 }],
      [{ v: 'Lin' }, { v: 6 }, null],
      [null, { v: '' }, { v: 24 }],
    ],
  };

  expect(
    spreadsheetSelectionSummary(sheet, {
      row: [0, 2],
      column: [0, 2],
    }),
  ).toEqual({
    average: 15,
    nonEmptyCount: 6,
    numericCount: 4,
    sum: 60,
  });
});

test('summarizes sparse spreadsheet cells without scanning empty coordinates', () => {
  const sheet: WorkSpreadsheetContent['sheets'][number] = {
    id: 'sheet-1',
    name: 'Sheet 1',
    celldata: [
      { r: 2, c: 1, v: { v: 2 } },
      { r: 50_000, c: 1, v: { v: 8 } },
      { r: 2, c: 8, v: { v: 100 } },
    ],
  };

  expect(
    spreadsheetSelectionSummary(sheet, {
      row: [0, 100_000],
      column: [0, 2],
    }),
  ).toEqual({
    average: 5,
    nonEmptyCount: 2,
    numericCount: 2,
    sum: 10,
  });
});
