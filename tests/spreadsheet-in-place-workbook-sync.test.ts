import type { CellMatrix } from '@fortune-sheet/core';
import { expect, test } from '@rstest/core';
import {
  spreadsheetCellTextEdits,
  synchronizeSpreadsheetWorkbookInPlace,
} from '../src/internal/features/work/editors/spreadsheet-in-place-workbook-sync';
import { registerImportedSpreadsheetMatrix } from '../src/internal/features/work/work-spreadsheet-matrix-profile';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

test('paints leaf-typing glyph frames via setCellValue without remounting', () => {
  const previous = certifiedWorkbook('sheet-1', [[{ v: 'Hi', m: 'Hi' }]]);
  // Leaf typing clones row arrays and drops the WeakMap profile.
  const next = workbook('sheet-1', [[{ v: 'Hit', m: 'Hit' }]]);
  const cellValues: Array<[number, number, unknown, unknown]> = [];
  const updates: WorkSpreadsheetContent['sheets'][] = [];

  expect(
    synchronizeSpreadsheetWorkbookInPlace(
      {
        setCellValue: (row, column, value, options) => {
          cellValues.push([row, column, value, options]);
        },
        updateSheet: (sheets) => updates.push(sheets),
      },
      previous,
      next,
      next.sheets,
      false,
    ),
  ).toBe(true);
  expect(cellValues).toEqual([[0, 0, 'Hit', { id: 'sheet-1' }]]);
  expect(updates).toEqual([]);
});

test('adopts a certified simple workbook without remounting', () => {
  const previous = workbook('sheet-1', [[{ v: 'Loading', m: 'Loading' }]]);
  const next = certifiedWorkbook('sheet-1', [[{ v: 'Ready', m: 'Ready' }]]);
  const cellValues: unknown[] = [];
  const updates: WorkSpreadsheetContent['sheets'][] = [];

  const synchronized = synchronizeSpreadsheetWorkbookInPlace(
    {
      setCellValue: (...args) => cellValues.push(args),
      updateSheet: (sheets) => updates.push(sheets),
    },
    previous,
    next,
    next.sheets,
    false,
  );

  expect(synchronized).toBe(true);
  expect(cellValues).toEqual([[0, 0, 'Ready', { id: 'sheet-1' }]]);
  expect(updates).toEqual([]);
});

test('keeps the remount path for previews and identity changes', () => {
  const previous = workbook('sheet-1', [[{ v: 'Loading' }]]);
  const next = certifiedWorkbook('sheet-1', [[{ v: 'Ready' }]]);
  const updateSheet = () => {
    throw new Error('must not update');
  };
  const setCellValue = () => {
    throw new Error('must not set cell');
  };

  expect(
    synchronizeSpreadsheetWorkbookInPlace(
      { updateSheet, setCellValue },
      previous,
      next,
      next.sheets,
      true,
    ),
  ).toBe(false);
  expect(
    synchronizeSpreadsheetWorkbookInPlace(
      { updateSheet, setCellValue },
      previous,
      certifiedWorkbook('sheet-2', [[{ v: 'Ready' }]]),
      next.sheets,
      false,
    ),
  ).toBe(false);
});

test('updates cell text in place even when the sheet already has structure', () => {
  const previous = certifiedWorkbook('sheet-1', [[{ v: 'Loading', m: 'Loading' }]]);
  previous.sheets[0].config = { merge: {} };
  previous.sheets[0].defaultRowHeight = 24;
  const next = certifiedWorkbook('sheet-1', [[{ v: 'Ready', m: 'Ready' }]]);
  next.sheets[0].config = { merge: {} };
  next.sheets[0].defaultRowHeight = 24;
  const cellValues: unknown[] = [];

  expect(
    synchronizeSpreadsheetWorkbookInPlace(
      {
        setCellValue: (...args) => cellValues.push(args),
        updateSheet: () => {
          throw new Error('must not updateSheet');
        },
      },
      previous,
      next,
      next.sheets,
      false,
    ),
  ).toBe(true);
  expect(cellValues).toEqual([[0, 0, 'Ready', { id: 'sheet-1' }]]);
});

test('falls back to updateSheet when cell style changes with the value', () => {
  const previous = certifiedWorkbook('sheet-1', [[{ v: 'Hi', m: 'Hi' }]]);
  const next = certifiedWorkbook('sheet-1', [
    [{ v: 'Hit', m: 'Hit', bl: 1 }],
  ]);
  const updates: WorkSpreadsheetContent['sheets'][] = [];

  expect(spreadsheetCellTextEdits(previous, next)).toBeNull();
  expect(
    synchronizeSpreadsheetWorkbookInPlace(
      {
        setCellValue: () => {
          throw new Error('must not setCellValue');
        },
        updateSheet: (sheets) => updates.push(sheets),
      },
      previous,
      next,
      next.sheets,
      false,
    ),
  ).toBe(true);
  expect(updates).toEqual([next.sheets]);
});

test('keeps the remount path when sheet structure changes', () => {
  const previous = certifiedWorkbook('sheet-1', [[{ v: 'Loading' }]]);
  const next = certifiedWorkbook('sheet-1', [[{ v: 'Ready' }]]);
  next.sheets[0].config = { authority: { sheet: 1 } };
  const updateSheet = () => {
    throw new Error('must not update');
  };
  const setCellValue = () => {
    throw new Error('must not set cell');
  };

  expect(
    synchronizeSpreadsheetWorkbookInPlace(
      { updateSheet, setCellValue },
      previous,
      next,
      next.sheets,
      false,
    ),
  ).toBe(false);
});

test('falls back to remounting if the Fortune update rejects the workbook', () => {
  const previous = workbook('sheet-1', [
    [{ v: 'Loading', m: 'Loading', bl: 1 }],
  ]);
  const next = certifiedWorkbook('sheet-1', [
    [{ v: 'Ready', m: 'Ready', bl: 1 }],
  ]);

  expect(
    synchronizeSpreadsheetWorkbookInPlace(
      {
        setCellValue: () => {
          throw new Error('must not setCellValue');
        },
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
