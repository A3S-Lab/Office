import { expect, test } from '@rstest/core';
import {
  spreadsheetFormulaResultText,
  spreadsheetSelectionSummary,
} from '../src/internal/features/work/editors/spreadsheet-editor-support';
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

test('shows a formula cell display value and hides the formula text', () => {
  expect(spreadsheetFormulaResultText(null)).toBeNull();
  expect(spreadsheetFormulaResultText({ v: 30, m: '30' })).toBeNull();
  expect(
    spreadsheetFormulaResultText({ f: '=SUM(A1:A2)', m: '=SUM(A1:A2)' }),
  ).toBeNull();
  expect(
    spreadsheetFormulaResultText({ f: '=SUM(A1:A2)', v: 30, m: '=SUM(A1:A2)' }),
  ).toBe('30');
  expect(
    spreadsheetFormulaResultText({ f: '=SUM(A1:A2)', v: 30, m: '30' }),
  ).toBe('30');
  expect(spreadsheetFormulaResultText({ f: '=SUM(A1:A2)', v: 30 })).toBe('30');
});
