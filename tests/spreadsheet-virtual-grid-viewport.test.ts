import { expect, test } from '@rstest/core';
import {
  SPREADSHEET_VIRTUAL_GRID_DEFAULT_COLUMN_WIDTH,
  SPREADSHEET_VIRTUAL_GRID_DEFAULT_ROW_HEIGHT,
  spreadsheetVirtualGridViewport,
} from '../src/internal/features/work/editors/spreadsheet-virtual-grid-viewport';

test('spreadsheetVirtualGridViewport covers only the visible window', () => {
  const viewport = spreadsheetVirtualGridViewport({
    scrollTop: 0,
    scrollLeft: 0,
    clientHeight: SPREADSHEET_VIRTUAL_GRID_DEFAULT_ROW_HEIGHT * 3,
    clientWidth: SPREADSHEET_VIRTUAL_GRID_DEFAULT_COLUMN_WIDTH * 2,
    rowCount: 1_048_576,
    columnCount: 16_384,
  });

  expect(viewport).toEqual({
    rowStart: 0,
    rowEnd: 3,
    columnStart: 0,
    columnEnd: 2,
    offsetTop: 0,
    offsetLeft: 0,
  });
});

test('spreadsheetVirtualGridViewport scrolls into a sparse far window without densifying', () => {
  const viewport = spreadsheetVirtualGridViewport({
    scrollTop: SPREADSHEET_VIRTUAL_GRID_DEFAULT_ROW_HEIGHT * 100_000,
    scrollLeft: SPREADSHEET_VIRTUAL_GRID_DEFAULT_COLUMN_WIDTH * 200,
    clientHeight: SPREADSHEET_VIRTUAL_GRID_DEFAULT_ROW_HEIGHT * 2 + 12,
    clientWidth: SPREADSHEET_VIRTUAL_GRID_DEFAULT_COLUMN_WIDTH + 40,
    rowCount: 1_048_576,
    columnCount: 16_384,
  });

  expect(viewport.rowStart).toBe(100_000);
  expect(viewport.rowEnd).toBe(100_003);
  expect(viewport.columnStart).toBe(200);
  expect(viewport.columnEnd).toBe(202);
  expect(viewport.offsetTop).toBe(0);
  expect(viewport.offsetLeft).toBe(0);
});

test('spreadsheetVirtualGridViewport clamps to sheet bounds and applies overscan', () => {
  const viewport = spreadsheetVirtualGridViewport({
    scrollTop: SPREADSHEET_VIRTUAL_GRID_DEFAULT_ROW_HEIGHT * 8 + 6,
    scrollLeft: SPREADSHEET_VIRTUAL_GRID_DEFAULT_COLUMN_WIDTH * 4 + 10,
    clientHeight: SPREADSHEET_VIRTUAL_GRID_DEFAULT_ROW_HEIGHT * 2,
    clientWidth: SPREADSHEET_VIRTUAL_GRID_DEFAULT_COLUMN_WIDTH * 2,
    rowCount: 10,
    columnCount: 6,
    overscan: 1,
  });

  // Overscan expands the painted range; offsets are relative to rowStart/columnStart
  // so the first overscanned row sits above the scroll origin.
  expect(viewport).toEqual({
    rowStart: 7,
    rowEnd: 10,
    columnStart: 3,
    columnEnd: 6,
    offsetTop: 30,
    offsetLeft: 106,
  });
});

test('spreadsheetVirtualGridViewport stays empty for zero-sized sheets', () => {
  expect(
    spreadsheetVirtualGridViewport({
      scrollTop: 0,
      scrollLeft: 0,
      clientHeight: 100,
      clientWidth: 100,
      rowCount: 0,
      columnCount: 0,
    }),
  ).toEqual({
    rowStart: 0,
    rowEnd: 0,
    columnStart: 0,
    columnEnd: 0,
    offsetTop: 0,
    offsetLeft: 0,
  });
});
