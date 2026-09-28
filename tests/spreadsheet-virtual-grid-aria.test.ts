import { expect, test } from '@rstest/core';
import {
  spreadsheetVirtualGridActiveCellId,
  spreadsheetVirtualGridAria,
} from '../src/internal/features/work/editors/spreadsheet-virtual-grid-aria';

test('spreadsheetVirtualGridAria exposes a real grid widget for the active cell', () => {
  const activeCellId = spreadsheetVirtualGridActiveCellId(2, 3);
  expect(activeCellId).toBe('a3s-ss-cell-2-3');
  expect(
    spreadsheetVirtualGridAria({
      sheetName: 'Sheet1',
      rowCount: 1_048_576,
      columnCount: 16_384,
      activeRow: 2,
      activeColumn: 3,
    }),
  ).toEqual({
    role: 'grid',
    'aria-label': 'Sheet1',
    'aria-rowcount': 1_048_576,
    'aria-colcount': 16_384,
    'aria-activedescendant': 'a3s-ss-cell-2-3',
  });
});

test('spreadsheetVirtualGridAria clamps active descendant into the sheet', () => {
  expect(
    spreadsheetVirtualGridAria({
      sheetName: 'Budget',
      rowCount: 10,
      columnCount: 5,
      activeRow: 99,
      activeColumn: -1,
    })['aria-activedescendant'],
  ).toBe('a3s-ss-cell-9-0');
});
