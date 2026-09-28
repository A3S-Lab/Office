import { expect, test } from '@rstest/core';
import type { CellMatrix } from '@fortune-sheet/core';
import { spreadsheetVirtualGridPaintCells } from '../src/internal/features/work/editors/spreadsheet-virtual-grid-paint';
import {
  SPREADSHEET_VIRTUAL_GRID_DEFAULT_COLUMN_WIDTH,
  SPREADSHEET_VIRTUAL_GRID_DEFAULT_ROW_HEIGHT,
} from '../src/internal/features/work/editors/spreadsheet-virtual-grid-viewport';

test('spreadsheetVirtualGridPaintCells visits only populated cells in the viewport', () => {
  const data: CellMatrix = [];
  data.length = 1_048_576;
  data[0] = [];
  data[0].length = 16_384;
  data[0][0] = { v: 'Anchor', m: 'Anchor' };
  data[100_000] = [];
  data[100_000].length = 16_384;
  data[100_000][200] = { v: 42, m: '42', bg: '#cc0000', fc: '#ffffff' };
  data[100_000][201] = { f: '=A1', m: '1', v: 1 };
  data[999_999] = [];
  data[999_999][16_383] = { v: 'Tail', m: 'Tail' };

  const cells = spreadsheetVirtualGridPaintCells({
    data,
    viewport: {
      rowStart: 100_000,
      rowEnd: 100_002,
      columnStart: 200,
      columnEnd: 202,
      offsetTop: 0,
      offsetLeft: 0,
    },
  });

  expect(cells).toEqual([
    {
      row: 100_000,
      column: 200,
      text: '42',
      background: '#cc0000',
      color: '#ffffff',
      x: 0,
      y: 0,
      width: SPREADSHEET_VIRTUAL_GRID_DEFAULT_COLUMN_WIDTH,
      height: SPREADSHEET_VIRTUAL_GRID_DEFAULT_ROW_HEIGHT,
    },
    {
      row: 100_000,
      column: 201,
      text: '1',
      background: undefined,
      color: undefined,
      x: SPREADSHEET_VIRTUAL_GRID_DEFAULT_COLUMN_WIDTH,
      y: 0,
      width: SPREADSHEET_VIRTUAL_GRID_DEFAULT_COLUMN_WIDTH,
      height: SPREADSHEET_VIRTUAL_GRID_DEFAULT_ROW_HEIGHT,
    },
  ]);
});

test('spreadsheetVirtualGridPaintCells spans a merge and skips covered cells', () => {
  const data: CellMatrix = [];
  data[0] = [];
  data[0][0] = { v: 'Title', m: 'Title', bg: '#0f7b3b' };
  data[0][1] = { v: 'covered', m: 'covered' };

  const cells = spreadsheetVirtualGridPaintCells({
    data,
    merges: [{ r: 0, c: 0, rs: 1, cs: 2 }],
    viewport: {
      rowStart: 0,
      rowEnd: 2,
      columnStart: 0,
      columnEnd: 2,
      offsetTop: 0,
      offsetLeft: 0,
    },
  });

  expect(cells).toEqual([
    {
      row: 0,
      column: 0,
      text: 'Title',
      background: '#0f7b3b',
      color: undefined,
      x: 0,
      y: 0,
      width: SPREADSHEET_VIRTUAL_GRID_DEFAULT_COLUMN_WIDTH * 2,
      height: SPREADSHEET_VIRTUAL_GRID_DEFAULT_ROW_HEIGHT,
    },
  ]);
});

test('spreadsheetVirtualGridPaintCells skips empty viewport windows', () => {
  const data: CellMatrix = [];
  data.length = 100;
  data[0] = [];
  data[0][0] = { v: 'Only', m: 'Only' };

  expect(
    spreadsheetVirtualGridPaintCells({
      data,
      viewport: {
        rowStart: 10,
        rowEnd: 12,
        columnStart: 3,
        columnEnd: 5,
        offsetTop: 0,
        offsetLeft: 0,
      },
    }),
  ).toEqual([]);
});
