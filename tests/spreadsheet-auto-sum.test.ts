import { describe, expect, test } from '@rstest/core';
import {
  planSpreadsheetAutoSum,
  spreadsheetAutoSumMaximumTargets,
} from '../src/internal/features/work/editors/spreadsheet-auto-sum';
import type { WorkSpreadsheetSheet } from '../src/internal/features/work/work-types';

describe('spreadsheet AutoSum planning', () => {
  test('infers the nearest numeric run above a blank cell', () => {
    const sheet = matrixSheet([
      [{ v: 'Revenue' }],
      [{ v: 12 }],
      [{ v: 18 }],
      [null],
    ]);

    expect(
      planSpreadsheetAutoSum(sheet, { row: [3, 3], column: [0, 0] }, 'sum'),
    ).toEqual({
      formulaBarValue: '=SUM(A2:A3)',
      function: 'sum',
      selection: {
        row: [3, 3],
        column: [0, 0],
        row_focus: 3,
        column_focus: 0,
      },
      targetCellCount: 1,
      writes: [
        {
          range: { row: [3, 3], column: [0, 0] },
          values: [['=SUM(A2:A3)']],
        },
      ],
    });
  });

  test('falls back to the nearest numeric run on the left', () => {
    const sheet = matrixSheet([
      [{ v: 'Label' }, { v: 4 }, { f: '=B1*2', v: 8 }, null],
    ]);

    expect(
      planSpreadsheetAutoSum(sheet, { row: [0, 0], column: [3, 3] }, 'average')
        ?.formulaBarValue,
    ).toBe('=AVERAGE(B1:C1)');
  });

  test('writes a totals row without overwriting label or text columns', () => {
    const sheet = matrixSheet([
      [{ v: 'Region' }, { v: 10 }, { v: 'Owner' }, { v: 20 }],
      [{ v: 'Total' }, null, null, null],
    ]);

    expect(
      planSpreadsheetAutoSum(sheet, { row: [0, 1], column: [0, 3] }, 'sum'),
    ).toEqual({
      formulaBarValue: '=SUM(B1:B1)',
      function: 'sum',
      selection: {
        row: [1, 1],
        column: [1, 3],
        row_focus: 1,
        column_focus: 1,
      },
      targetCellCount: 2,
      writes: [
        {
          range: { row: [1, 1], column: [1, 1] },
          values: [['=SUM(B1:B1)']],
        },
        {
          range: { row: [1, 1], column: [3, 3] },
          values: [['=SUM(D1:D1)']],
        },
      ],
    });
  });

  test('writes a totals column while retaining text-only rows', () => {
    const sheet = matrixSheet([
      [{ v: 2 }, { v: 3 }, null],
      [{ v: 'Owner' }, { v: 'A3S' }, null],
      [{ v: 5 }, { v: 7 }, null],
    ]);

    expect(
      planSpreadsheetAutoSum(sheet, { row: [0, 2], column: [0, 2] }, 'max')
        ?.writes,
    ).toEqual([
      {
        range: { row: [0, 0], column: [2, 2] },
        values: [['=MAX(A1:B1)']],
      },
      {
        range: { row: [2, 2], column: [2, 2] },
        values: [['=MAX(A3:B3)']],
      },
    ]);
  });

  test('bounds combined row and column totals and resolves their shared target once', () => {
    const combined = matrixSheet([
      [{ v: 1 }, { v: 2 }, null],
      [{ v: 3 }, { v: 4 }, null],
      [null, null, null],
    ]);
    const sharedTarget = matrixSheet([
      [{ v: 'Label' }, { v: 2 }],
      [{ v: 3 }, null],
    ]);

    expect(
      planSpreadsheetAutoSum(combined, { row: [0, 2], column: [0, 2] }, 'sum')
        ?.selection,
    ).toEqual({
      row: [0, 2],
      column: [0, 2],
      row_focus: 0,
      column_focus: 2,
    });
    expect(
      planSpreadsheetAutoSum(
        sharedTarget,
        { row: [0, 1], column: [0, 1] },
        'sum',
      ),
    ).toMatchObject({
      formulaBarValue: '=SUM(B1:B1)',
      targetCellCount: 1,
      writes: [
        {
          range: { row: [1, 1], column: [1, 1] },
          values: [['=SUM(B1:B1)']],
        },
      ],
    });
  });

  test('reads sparse celldata without projecting a dense matrix', () => {
    const sheet = {
      id: 'sheet-1',
      name: 'Sparse',
      row: 1_048_576,
      column: 16_384,
      celldata: [
        { r: 10, c: 4, v: { v: 2 } },
        { r: 11, c: 4, v: { v: 3 } },
      ],
    } satisfies WorkSpreadsheetSheet;

    expect(
      planSpreadsheetAutoSum(sheet, { row: [12, 12], column: [4, 4] }, 'min')
        ?.formulaBarValue,
    ).toBe('=MIN(E11:E12)');
    expect(sheet).not.toHaveProperty('data');
  });

  test('rejects overwrites, empty sources, and excessive target counts', () => {
    const occupiedTarget = matrixSheet([[{ v: 2 }], [{ v: 'Do not replace' }]]);
    const emptySource = matrixSheet([[{ v: 'Header' }], [null]]);
    const tooWide = matrixSheet([
      Array.from({ length: spreadsheetAutoSumMaximumTargets + 1 }, () => ({
        v: 1,
      })),
      [],
    ]);

    expect(
      planSpreadsheetAutoSum(
        occupiedTarget,
        { row: [0, 1], column: [0, 0] },
        'sum',
      ),
    ).toBeNull();
    expect(
      planSpreadsheetAutoSum(
        emptySource,
        { row: [1, 1], column: [0, 0] },
        'count',
      ),
    ).toBeNull();
    expect(
      planSpreadsheetAutoSum(
        tooWide,
        {
          row: [0, 1],
          column: [0, spreadsheetAutoSumMaximumTargets],
        },
        'sum',
      ),
    ).toBeNull();
  });
});

function matrixSheet(data: WorkSpreadsheetSheet['data']): WorkSpreadsheetSheet {
  return {
    id: 'sheet-1',
    name: 'Sheet 1',
    data,
  };
}
