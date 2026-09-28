import { expect, test } from '@rstest/core';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';
import { spreadsheetVirtualGridCommitCell } from '../src/internal/features/work/editors/spreadsheet-virtual-grid-commit';

const base: WorkSpreadsheetContent = {
  type: 'spreadsheet',
  sheets: [
    {
      id: 'sheet-1',
      name: 'Sheet1',
      row: 60,
      column: 26,
      data: [],
    },
  ],
};

test('spreadsheetVirtualGridCommitCell writes sparse literals without densifying', () => {
  const next = spreadsheetVirtualGridCommitCell(base, 'sheet-1', 100, 5, '42');
  const sheet = next.sheets[0];
  expect(Object.keys(sheet?.data ?? [])).toEqual(['100']);
  expect(Object.keys(sheet?.data?.[100] ?? [])).toEqual(['5']);
  expect(sheet?.data?.[100]?.[5]).toEqual({ v: 42, m: '42' });
});

test('spreadsheetVirtualGridCommitCell replaces formulas with literals and clears blanks', () => {
  const withFormula: WorkSpreadsheetContent = {
    ...base,
    sheets: [
      {
        ...base.sheets[0]!,
        data: [[{ f: '=A2', v: 1, m: '1' }]],
      },
    ],
  };
  const replaced = spreadsheetVirtualGridCommitCell(
    withFormula,
    'sheet-1',
    0,
    0,
    'hello',
  );
  expect(replaced.sheets[0]?.data?.[0]?.[0]).toEqual({ v: 'hello', m: 'hello' });

  const cleared = spreadsheetVirtualGridCommitCell(
    replaced,
    'sheet-1',
    0,
    0,
    '   ',
  );
  expect(cleared.sheets[0]?.data?.[0]?.[0]).toBeNull();
});
