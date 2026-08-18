import { expect, test } from '@rstest/core';
import {
  sameSpreadsheetHistoryContent,
  sameSpreadsheetWorkbookState,
  spreadsheetSheetsForFortune,
  spreadsheetSheetsFromFortune,
} from '../src/internal/features/work/editors/spreadsheet-editor-support';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

test('projects sparse workbook cells without cloning logical empty ranges', () => {
  const data = sparseMatrix();
  const projected = spreadsheetSheetsForFortune([
    {
      id: 'sheet-1',
      name: 'Sparse',
      row: 1_000_000,
      column: 16_384,
      data,
      config: {
        merge: {
          '0_0': { r: 0, c: 0, rs: 1, cs: 2 },
        },
      },
    },
  ]);

  expect(projected[0]?.data).toBeUndefined();
  expect(projected[0]?.celldata).toEqual([
    {
      r: 0,
      c: 0,
      v: { v: 'Anchor', m: 'Anchor', mc: { r: 0, c: 0, rs: 1, cs: 2 } },
    },
    { r: 0, c: 1, v: { mc: { r: 0, c: 0 } } },
    { r: 999_999, c: 16_383, v: { v: 'Tail', m: 'Tail' } },
  ]);
  expect(Object.keys(data)).toEqual(['0', '999999']);
  expect(data[0]?.[0]).not.toHaveProperty('mc');
});

test('compares workbook state without JSON serialization', () => {
  const sentinel = {
    toJSON(): never {
      throw new Error('Workbook comparison serialized its input.');
    },
  };
  const sheet = {
    id: 'sheet-1',
    name: 'Sheet 1',
    data: [[{ v: 1, m: '1' }]],
    sentinel,
  } as WorkSpreadsheetContent['sheets'][number] & { sentinel: typeof sentinel };

  expect(sameSpreadsheetWorkbookState([sheet], [{ ...sheet }])).toBe(true);
});

test('ignores Fortune formula result normalization in workbook comparisons', () => {
  const rendered: WorkSpreadsheetContent['sheets'] = [
    {
      id: 'sheet-1',
      name: 'Formula',
      celldata: [{ r: 0, c: 0, v: { f: '=1+1', v: 2 } }],
    },
  ];
  const normalized: WorkSpreadsheetContent['sheets'] = [
    {
      ...rendered[0],
      celldata: [
        {
          r: 0,
          c: 0,
          v: { f: '=1+1', v: 2, m: '2', ct: { fa: 'General', t: 'n' } },
        },
      ],
    },
  ];

  expect(sameSpreadsheetWorkbookState(normalized, rendered)).toBe(true);
  const changedFormula: WorkSpreadsheetContent['sheets'] = [
    {
      ...normalized[0],
      celldata: [{ r: 0, c: 0, v: { f: '=2+2', v: 4, m: '4' } }],
    },
  ];
  expect(sameSpreadsheetWorkbookState(changedFormula, rendered)).toBe(false);
});

test('ignores Fortune default zoom normalization in workbook comparisons', () => {
  const rendered: WorkSpreadsheetContent['sheets'] = [
    {
      id: 'sheet-1',
      name: 'Default zoom',
      celldata: [{ r: 0, c: 0, v: { v: 'Value', m: 'Value' } }],
    },
  ];
  const normalized: WorkSpreadsheetContent['sheets'] = [
    { ...rendered[0], zoomRatio: 1 },
  ];

  expect(sameSpreadsheetWorkbookState(normalized, rendered)).toBe(true);
  expect(
    sameSpreadsheetWorkbookState(
      [{ ...rendered[0], zoomRatio: 1.25 }],
      rendered,
    ),
  ).toBe(false);
});

test('ignores Fortune calculation bookkeeping in workbook comparisons', () => {
  const rendered: WorkSpreadsheetContent['sheets'] = [
    {
      id: 'sheet-1',
      name: 'Calculated',
      celldata: [{ r: 0, c: 0, v: { f: '=1+1', v: 2 } }],
    },
  ];
  const normalized: WorkSpreadsheetContent['sheets'] = [
    {
      ...rendered[0],
      calcChain: [{ r: 0, c: 0, id: 'sheet-1' }],
      dynamicArray_compute: { '0_0': { r: 0, c: 0 } },
    },
  ];

  expect(sameSpreadsheetWorkbookState(normalized, rendered)).toBe(true);
});

test('treats Fortune visible-row caches as the same sparse workbook state', () => {
  const source = {
    id: 'sheet-1',
    name: 'Sparse',
    row: 1_000_000,
    column: 16_384,
    data: sparseMatrix(),
  };
  const projected = spreadsheetSheetsForFortune([source]);
  const fortuneData: NonNullable<
    WorkSpreadsheetContent['sheets'][number]['data']
  > = [];
  fortuneData.length = 1_000_000;
  for (let row = 0; row < 100; row += 1) {
    fortuneData[row] = [];
    fortuneData[row].length = 16_384;
  }
  fortuneData[0][0] = { v: 'Anchor', m: 'Anchor' };
  fortuneData[999_999] = [];
  fortuneData[999_999].length = 16_384;
  fortuneData[999_999][16_383] = { v: 'Tail', m: 'Tail' };
  const projectedSheet = projected[0];
  if (!projectedSheet) throw new Error('Projected sheet is unavailable.');
  const changed = [
    { ...projectedSheet, celldata: undefined, data: fortuneData },
  ];

  expect(sameSpreadsheetWorkbookState(changed, projected)).toBe(true);
  const controlled = spreadsheetSheetsFromFortune(changed, [source]);
  expect(Object.keys(controlled[0]?.data ?? [])).toEqual(['0', '999999']);
});

test('materializes only a newly edited far row at the Fortune boundary', () => {
  const source = {
    id: 'sheet-1',
    name: 'Sparse',
    row: 1_000_000,
    column: 16_384,
    data: sparseMatrix(),
  };
  const fortuneData: NonNullable<
    WorkSpreadsheetContent['sheets'][number]['data']
  > = [];
  fortuneData.length = 1_000_000;
  for (let row = 0; row < 100; row += 1) {
    fortuneData[row] = [];
    fortuneData[row].length = 16_384;
  }
  fortuneData[0][0] = source.data[0]?.[0] ?? null;
  fortuneData[800_000] = [];
  fortuneData[800_000].length = 16_384;
  fortuneData[800_000][12_000] = { v: 'Edited', m: 'Edited' };
  fortuneData[999_999] = [];
  fortuneData[999_999].length = 16_384;
  fortuneData[999_999][16_383] = source.data[999_999]?.[16_383] ?? null;

  const controlled = spreadsheetSheetsFromFortune(
    [{ ...source, data: fortuneData }],
    [source],
  );
  expect(Object.keys(controlled[0]?.data ?? [])).toEqual([
    '0',
    '800000',
    '999999',
  ]);
});

test('detects values added inside sparse history arrays', () => {
  const left = sparseHistoryWorkbook();
  const right = sparseHistoryWorkbook();
  const rightRow = right.sheets[0]?.data?.[0];
  if (!rightRow) throw new Error('Expected a sparse history row.');
  rightRow[1] = { v: 'new', m: 'new' };

  expect(sameSpreadsheetHistoryContent(left, right)).toBe(false);
});

function sparseMatrix(): NonNullable<
  WorkSpreadsheetContent['sheets'][number]['data']
> {
  const data: NonNullable<WorkSpreadsheetContent['sheets'][number]['data']> =
    [];
  data.length = 1_000_000;
  data[0] = [];
  data[0][0] = { v: 'Anchor', m: 'Anchor' };
  data[999_999] = [];
  data[999_999][16_383] = { v: 'Tail', m: 'Tail' };
  return data;
}

function sparseHistoryWorkbook(): WorkSpreadsheetContent {
  const row: NonNullable<
    WorkSpreadsheetContent['sheets'][number]['data']
  >[number] = [];
  row.length = 2;
  row[0] = { v: 'existing', m: 'existing' };
  return {
    type: 'spreadsheet',
    sheets: [{ id: 'sheet-1', name: 'Sheet 1', data: [row] }],
  };
}
