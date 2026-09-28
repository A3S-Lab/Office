import { expect, test } from '@rstest/core';
import { spreadsheetActiveCellAnnouncement } from '../src/internal/features/work/editors/spreadsheet-active-cell-announcement';

test('spreadsheetActiveCellAnnouncement names blank and valued cells', () => {
  expect(
    spreadsheetActiveCellAnnouncement({ row: 0, column: 0 }),
  ).toBe('A1, blank');
  expect(
    spreadsheetActiveCellAnnouncement({
      row: 1,
      column: 2,
      displayValue: '42',
    }),
  ).toBe('C2, 42');
  expect(
    spreadsheetActiveCellAnnouncement({
      row: 0,
      column: 0,
      formula: '=SUM(A2:A4)',
      displayValue: '9',
    }),
  ).toBe('A1, formula =SUM(A2:A4)');
});

test('spreadsheetActiveCellAnnouncement stays quiet without a focus cell', () => {
  expect(spreadsheetActiveCellAnnouncement({})).toBe('');
  expect(
    spreadsheetActiveCellAnnouncement({ row: Number.NaN, column: 0 }),
  ).toBe('');
});
