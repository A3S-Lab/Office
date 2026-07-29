import { expect, test } from '@rstest/core';
import {
  spreadsheetFindMatches,
  type SpreadsheetFindMatch,
} from '../src/internal/features/work/editors/spreadsheet-find';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

const sheet: WorkSpreadsheetContent['sheets'][number] = {
  id: 'sheet-1',
  name: '执行看板',
  data: [
    [
      { v: 'alpha source', m: 'Alpha 计划' },
      { v: 2, m: '2.00', f: '=SUM(A1:A1)' },
    ],
  ],
  celldata: [
    { r: 0, c: 0, v: { v: 'duplicate alpha' } },
    { r: 2, c: 1, v: { v: 'Sparse Alpha' } },
  ],
};

test('finds displayed and sparse cell text without duplicating coordinates', () => {
  expect(compactMatches(spreadsheetFindMatches(sheet, 'ALPHA'))).toEqual([
    { reference: 'A1', text: 'Alpha 计划' },
    { reference: 'B3', text: 'Sparse Alpha' },
  ]);
});

test('finds raw values and formulas case-insensitively', () => {
  expect(compactMatches(spreadsheetFindMatches(sheet, 'alpha source'))).toEqual(
    [{ reference: 'A1', text: 'Alpha 计划' }],
  );
  expect(compactMatches(spreadsheetFindMatches(sheet, 'sum('))).toEqual([
    { reference: 'B1', text: '2.00' },
  ]);
  expect(spreadsheetFindMatches(sheet, '')).toEqual([]);
});

function compactMatches(matches: readonly SpreadsheetFindMatch[]) {
  return matches.map(({ reference, text }) => ({ reference, text }));
}
