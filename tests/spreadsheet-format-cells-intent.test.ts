import { expect, test } from '@rstest/core';
import { normalizeSpreadsheetFormatCellsOpenIntent } from '../src/internal/features/work/editors/spreadsheet-format-cells-intent';

test('normalizes default and font-focused Format Cells routes', () => {
  expect(normalizeSpreadsheetFormatCellsOpenIntent(undefined)).toEqual({
    tab: 'number',
  });
  expect(
    normalizeSpreadsheetFormatCellsOpenIntent({
      tab: 'font',
      focus: 'fontFamily',
    }),
  ).toEqual({ tab: 'font', focus: 'fontFamily' });
  expect(
    normalizeSpreadsheetFormatCellsOpenIntent({
      tab: 'font',
      focus: 'fontSize',
    }),
  ).toEqual({ tab: 'font', focus: 'fontSize' });
  expect(normalizeSpreadsheetFormatCellsOpenIntent({ tab: 'border' })).toEqual({
    tab: 'border',
  });
});

test('rejects malformed or contradictory Format Cells routes', () => {
  for (const value of [
    null,
    'font',
    {},
    { tab: 'unknown' },
    { tab: 'number', focus: 'fontSize' },
    { tab: 'font', focus: 'unknown' },
    { tab: 'font', extra: true },
  ]) {
    expect(normalizeSpreadsheetFormatCellsOpenIntent(value)).toBeNull();
  }
});
