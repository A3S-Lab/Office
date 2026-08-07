import type { Cell } from '@fortune-sheet/core';
import { expect, test } from '@rstest/core';
import {
  captureSpreadsheetFormatPattern,
  spreadsheetCellFormatAttributes,
  spreadsheetFormatPainterBatches,
  spreadsheetFormatPainterTargetRange,
} from '../src/internal/features/work/editors/spreadsheet-format-painter';

test('captures only native cell formatting and owns copied values', () => {
  const source: Cell = {
    v: 'Revenue',
    f: '=SUM(B2:B4)',
    m: 'Revenue',
    bg: '#fff2cc',
    bl: 1,
    ff: 'Arial',
    ct: { fa: '0.00%', t: 'n', s: [{ v: 'not-formatting' }] },
    ps: {
      left: null,
      top: null,
      width: null,
      height: null,
      value: 'Keep this comment at the source',
      isShow: false,
    },
  };

  const pattern = captureSpreadsheetFormatPattern([[source, null]]);

  expect(pattern).not.toBeNull();
  expect(pattern?.rowCount).toBe(1);
  expect(pattern?.columnCount).toBe(2);
  expect(Object.keys(pattern?.cells[0]?.[0] ?? {})).toEqual(
    spreadsheetCellFormatAttributes,
  );
  expect(pattern?.cells[0]?.[0]).toMatchObject({
    bg: '#fff2cc',
    bl: 1,
    ff: 'Arial',
    ct: { fa: '0.00%', t: 'n' },
  });
  expect(pattern?.cells[0]?.[0]).not.toHaveProperty('v');
  expect(pattern?.cells[0]?.[0]).not.toHaveProperty('f');
  expect(pattern?.cells[0]?.[0]).not.toHaveProperty('ps');
  expect(pattern?.cells[0]?.[1]?.ct).toEqual({ fa: 'General', t: 'g' });

  source.ct = { fa: '#,##0', t: 'n' };
  source.bg = '#000000';
  expect(pattern?.cells[0]?.[0]).toMatchObject({
    bg: '#fff2cc',
    ct: { fa: '0.00%', t: 'n' },
  });
});

test('expands a single target to the source pattern and clamps sheet edges', () => {
  const pattern = captureSpreadsheetFormatPattern([
    [{ bg: '#ffffff' }, { bg: '#eeeeee' }, { bg: '#dddddd' }],
    [{ bg: '#cccccc' }, { bg: '#bbbbbb' }, { bg: '#aaaaaa' }],
  ]);
  if (!pattern) throw new Error('Expected a format pattern.');

  expect(
    spreadsheetFormatPainterTargetRange(
      { row: [4, 4], column: [2, 2] },
      pattern,
      { rowCount: 10, columnCount: 8 },
    ),
  ).toEqual({ row: [4, 5], column: [2, 4] });
  expect(
    spreadsheetFormatPainterTargetRange(
      { row: [8, 9], column: [5, 7] },
      pattern,
      { rowCount: 10, columnCount: 8 },
    ),
  ).toEqual({ row: [8, 9], column: [5, 7] });
  expect(
    spreadsheetFormatPainterTargetRange(
      { row: [9, 9], column: [7, 7] },
      pattern,
      { rowCount: 10, columnCount: 8 },
    ),
  ).toEqual({ row: [9, 9], column: [7, 7] });
});

test('tiles a multi-cell format pattern without copying values or formulas', () => {
  const pattern = captureSpreadsheetFormatPattern([
    [
      { v: 'A', bg: '#f00', bl: 1 },
      { v: 'B', bg: '#00f', bl: 0 },
    ],
    [
      { f: '=1+1', bg: '#00f', bl: 0 },
      { v: 2, bg: '#f00', bl: 1 },
    ],
  ]);
  if (!pattern) throw new Error('Expected a format pattern.');

  const batches = spreadsheetFormatPainterBatches(pattern, {
    row: [3, 6],
    column: [4, 7],
  });
  const backgrounds = expandedAttributeValues(batches, 'bg');
  const bold = expandedAttributeValues(batches, 'bl');

  expect(backgrounds).toEqual([
    ['#f00', '#00f', '#f00', '#00f'],
    ['#00f', '#f00', '#00f', '#f00'],
    ['#f00', '#00f', '#f00', '#00f'],
    ['#00f', '#f00', '#00f', '#f00'],
  ]);
  expect(bold).toEqual([
    [1, 0, 1, 0],
    [0, 1, 0, 1],
    [1, 0, 1, 0],
    [0, 1, 0, 1],
  ]);
  expect(batches.every((batch) => batch.attribute !== 'v')).toBe(true);
  expect(batches.every((batch) => batch.attribute !== 'f')).toBe(true);
});

function expandedAttributeValues(
  batches: ReturnType<typeof spreadsheetFormatPainterBatches>,
  attribute: (typeof spreadsheetCellFormatAttributes)[number],
): unknown[][] {
  const values = Array.from({ length: 4 }, () => Array<unknown>(4));
  for (const batch of batches.filter((item) => item.attribute === attribute)) {
    for (const range of batch.ranges) {
      for (let row = range.row[0]; row <= range.row[1]; row += 1) {
        for (
          let column = range.column[0];
          column <= range.column[1];
          column += 1
        ) {
          values[row - 3][column - 4] = batch.value;
        }
      }
    }
  }
  return values;
}
