import type { Cell } from '@fortune-sheet/core';
import { describe, expect, test } from '@rstest/core';
import {
  createSpreadsheetSortDialogSource,
  validateSpreadsheetSortRequest,
} from '../src/internal/features/work/editors/spreadsheet-sort';
import { sortSpreadsheetMatrix } from '../src/internal/features/work/editors/spreadsheet-sort-matrix';

describe('spreadsheet left-to-right sort', () => {
  test('moves complete columns by stable row keys and translates relative column references', () => {
    const result = sortSpreadsheetMatrix(
      [
        [cell(2), cell(1), cell(1), cell(2)],
        [cell('Gamma'), cell('Alpha'), cell('Beta'), cell('Delta')],
        [
          cell('Gamma!', { f: '=A2&$D$1' }),
          cell('Alpha!', { f: '=B2&$D$1' }),
          cell('Beta!', { f: '=C2&$D$1' }),
          cell('Delta!', { f: '=D2&$D$1' }),
        ],
      ],
      {
        sheetId: 'sheet-1',
        range: { row: [0, 2], column: [0, 3] },
        orientation: 'left-to-right',
        hasHeader: false,
        keys: [
          { index: 0, direction: 'ascending' },
          { index: 1, direction: 'descending' },
        ],
      },
    );

    expect(result).toMatchObject({ ok: true });
    if (!result.ok) throw new Error(result.message);
    expect(result.rows.map((row) => row.map((item) => item?.v))).toEqual([
      [1, 1, 2, 2],
      ['Beta', 'Alpha', 'Gamma', 'Delta'],
      ['Beta!', 'Alpha!', 'Gamma!', 'Delta!'],
    ]);
    expect(result.rows[2]?.map((item) => item?.f)).toEqual([
      '=A2&$D$1',
      '=B2&$D$1',
      '=C2&$D$1',
      '=D2&$D$1',
    ]);
  });

  test('applies a custom list to an absolute row key', () => {
    const result = sortSpreadsheetMatrix(
      [
        [cell('Mar'), cell('Jan'), cell('Feb')],
        [cell(30), cell(10), cell(20)],
      ],
      {
        sheetId: 'sheet-1',
        range: { row: [4, 5], column: [2, 4] },
        orientation: 'left-to-right',
        hasHeader: false,
        keys: [{ index: 4, customList: ['Jan', 'Feb', 'Mar'] }],
      },
    );

    expect(
      result.ok && result.rows.map((row) => row.map((item) => item?.v)),
    ).toEqual([
      ['Jan', 'Feb', 'Mar'],
      [10, 20, 30],
    ]);
  });

  test('composes same-row appearance priorities with a later value key', () => {
    const result = sortSpreadsheetMatrix(
      [
        [cell('Alpha'), cell('Red B'), cell('Red A'), cell('Blue')],
        [
          cell('Ready'),
          cell('Blocked', { bg: '#fce8e6' }),
          cell('Blocked', { bg: '#fce8e6' }),
          cell('Review', { bg: '#4472c4' }),
        ],
      ],
      {
        sheetId: 'sheet-1',
        range: { row: [0, 1], column: [0, 3] },
        orientation: 'left-to-right',
        hasHeader: false,
        keys: [
          {
            index: 1,
            sortOn: 'cell-color',
            color: '#fce8e6',
            position: 'first',
          },
          { index: 0, direction: 'ascending' },
        ],
      },
    );

    expect(result.ok && result.rows[0]?.map((item) => item?.v)).toEqual([
      'Red A',
      'Red B',
      'Alpha',
      'Blue',
    ]);
  });

  test('rejects horizontal headers, row keys outside the range, and one-column ranges', () => {
    const base = {
      sheetId: 'sheet-1',
      range: {
        row: [3, 5] as [number, number],
        column: [1, 4] as [number, number],
      },
      orientation: 'left-to-right' as const,
      hasHeader: false,
    };

    expect(
      validateSpreadsheetSortRequest({
        ...base,
        hasHeader: true,
        keys: [{ index: 3, direction: 'ascending' }],
      }),
    ).toMatchObject({ ok: false, code: 'invalid-header' });
    expect(
      validateSpreadsheetSortRequest({
        ...base,
        keys: [{ index: 9, direction: 'ascending' }],
      }),
    ).toMatchObject({ ok: false, code: 'row-out-of-range' });
    expect(
      validateSpreadsheetSortRequest({
        ...base,
        range: { row: [3, 5], column: [1, 1] },
        keys: [{ index: 3, direction: 'ascending' }],
      }),
    ).toMatchObject({ ok: false, code: 'not-enough-columns' });
  });

  test('rejects the operation atomically when a moved formula would leave the sheet', () => {
    const result = sortSpreadsheetMatrix(
      [[cell('Zed'), cell('Alpha', { f: '=A1' })]],
      {
        sheetId: 'sheet-1',
        range: { row: [0, 0], column: [0, 1] },
        orientation: 'left-to-right',
        hasHeader: false,
        keys: [{ index: 0, direction: 'ascending' }],
      },
    );

    expect(result).toMatchObject({
      ok: false,
      code: 'formula-reference-out-of-range',
    });
  });

  test('exposes absolute row fields and both active axes to the dialog', () => {
    const source = createSpreadsheetSortDialogSource(
      'sheet-1',
      'Quarterly plan',
      {
        range: { row: [3, 5], column: [1, 3] },
        activeColumn: 2,
        activeRow: 4,
      },
      [
        [cell('Metric'), cell('Q1'), cell('Q2')],
        [cell('Priority'), cell(2), cell(1)],
        [cell('Revenue'), cell(200), cell(100)],
      ],
    );

    expect(source).toMatchObject({
      columns: [
        { index: 1, label: 'B（Metric）' },
        { index: 2, label: 'C（Q1）' },
        { index: 3, label: 'D（Q2）' },
      ],
      rows: [
        { index: 3, label: '行 4' },
        { index: 4, label: '行 5' },
        { index: 5, label: '行 6' },
      ],
      value: {
        orientation: 'top-to-bottom',
        caseSensitive: false,
        textMethod: 'pinyin',
        hasHeader: true,
        keys: [{ index: 2, direction: 'ascending' }],
      },
    });
  });
});

function cell(value: Cell['v'], format: Partial<Cell> = {}): Cell {
  return { v: value, ...format };
}
