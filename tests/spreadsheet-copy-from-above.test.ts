import type { Sheet } from '@fortune-sheet/core';
import { expect, test } from '@rstest/core';
import {
  planSpreadsheetCopyFromAbove,
  type SpreadsheetCopyFromAboveKind,
} from '../src/internal/features/work/editors/spreadsheet-copy-from-above';

test('plans exact formula and cached-value copies without copying source styles', () => {
  const sheet: Sheet = {
    id: 'sheet-1',
    name: 'Sheet 1',
    row: 3,
    column: 2,
    data: [
      [
        { v: 2 },
        {
          f: '=$A$1+A1',
          v: 4,
          m: '4.00',
          bl: 1,
          bg: '#e2f0d9',
          ct: { fa: '0.00', t: 'n' },
        },
      ],
      [null, { v: 'Replace me', it: 1, bg: '#ddebf7' }],
    ],
  };

  expect(
    planSpreadsheetCopyFromAbove(sheet, { row: 1, column: 1 }, 'formula'),
  ).toEqual({
    formulaBarValue: '=$A$1+A1',
    kind: 'formula',
    source: { row: 0, column: 1 },
    targetRange: { row: [1, 1], column: [1, 1] },
    value: '=$A$1+A1',
  });
  expect(
    planSpreadsheetCopyFromAbove(sheet, { row: 1, column: 1 }, 'value'),
  ).toEqual({
    formulaBarValue: 4,
    kind: 'value',
    source: { row: 0, column: 1 },
    targetRange: { row: [1, 1], column: [1, 1] },
    value: 4,
  });
});

test('copies constants and sparse cached values through the same scalar plan', () => {
  const constant: Sheet = {
    id: 'sheet-1',
    name: 'Sheet 1',
    data: [[{ v: '001', m: '001' }], [{ v: 'Old' }]],
  };
  expect(
    planSpreadsheetCopyFromAbove(constant, { row: 1, column: 0 }, 'formula')
      ?.value,
  ).toBe('001');

  const sparse: Sheet = {
    id: 'sheet-1',
    name: 'Sheet 1',
    celldata: [
      { r: 8, c: 4, v: { f: '=A1+1', m: '7' } },
      { r: 9, c: 4, v: { v: 'Old' } },
    ],
  };
  expect(
    planSpreadsheetCopyFromAbove(sparse, { row: 9, column: 4 }, 'value'),
  ).toMatchObject({ formulaBarValue: '7', value: '7' });

  const blank: Sheet = {
    id: 'sheet-1',
    name: 'Sheet 1',
    data: [[], [{ f: '=A1', v: 1 }]],
  };
  expect(
    planSpreadsheetCopyFromAbove(blank, { row: 1, column: 0 }, 'value'),
  ).toMatchObject({ formulaBarValue: '', value: null });
});

test('rejects unavailable, no-op, protected, merged, and pivot targets', () => {
  const base: Sheet = {
    id: 'sheet-1',
    name: 'Sheet 1',
    row: 2,
    column: 2,
    data: [[{ v: 2 }], [{ v: 2 }]],
  };

  expect(
    planSpreadsheetCopyFromAbove(base, { row: 0, column: 0 }, 'value'),
  ).toBeNull();
  expect(
    planSpreadsheetCopyFromAbove(base, { row: 1_048_576, column: 0 }, 'value'),
  ).toBeNull();
  expect(
    planSpreadsheetCopyFromAbove(base, { row: 1, column: 0 }, 'value'),
  ).toBeNull();
  expect(
    planSpreadsheetCopyFromAbove(
      base,
      { row: 1, column: 0 },
      'unsupported' as SpreadsheetCopyFromAboveKind,
    ),
  ).toBeNull();

  for (const sheet of [
    { ...base, config: { authority: { sheet: 1 } } },
    {
      ...base,
      config: { merge: { '1_0': { r: 1, c: 0, rs: 1, cs: 2 } } },
    },
    { ...base, pivotTable: { enabled: true } },
  ]) {
    expect(
      planSpreadsheetCopyFromAbove(sheet, { row: 1, column: 0 }, 'formula'),
    ).toBeNull();
  }
});

test('fails closed for array, data-table, and external formula sources', () => {
  for (const sheet of [
    {
      id: 'sheet-1',
      name: 'Sheet 1',
      data: [[{ f: '=SEQUENCE(2)', v: 1 }], [{ v: 'Old' }]],
      formulaMetadata: {
        ranges: [
          {
            type: 'dynamic-array' as const,
            anchor: 'A1',
            reference: 'A1:A2',
            formula: '=SEQUENCE(2)',
          },
        ],
      },
    },
    {
      id: 'sheet-1',
      name: 'Sheet 1',
      data: [[{ f: '=[Book.xlsx]Sheet1!A1', v: 1 }], [{ v: 'Old' }]],
    },
    {
      id: 'sheet-1',
      name: 'Sheet 1',
      data: [[{ f: 'SUM(A1:A2)', v: 1 }], [{ v: 'Old' }]],
    },
  ]) {
    expect(
      planSpreadsheetCopyFromAbove(sheet, { row: 1, column: 0 }, 'formula'),
    ).toBeNull();
    expect(
      planSpreadsheetCopyFromAbove(sheet, { row: 1, column: 0 }, 'value'),
    ).not.toBeNull();
  }
});
