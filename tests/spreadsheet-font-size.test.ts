import type { Cell } from '@fortune-sheet/core';
import { describe, expect, test } from '@rstest/core';
import {
  canAdjustSpreadsheetFontSize,
  MAX_SPREADSHEET_FONT_SIZE_CELLS,
  nextSpreadsheetFontSize,
  spreadsheetFontSizeApiCalls,
} from '../src/internal/features/work/editors/spreadsheet-font-size-command';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

describe('spreadsheet font-size stepping', () => {
  test('moves custom and standard sizes through the shared WPS scale', () => {
    expect(nextSpreadsheetFontSize(10, 'grow')).toBe(11);
    expect(nextSpreadsheetFontSize(10.5, 'grow')).toBe(11);
    expect(nextSpreadsheetFontSize(10.5, 'shrink')).toBe(10);
    expect(nextSpreadsheetFontSize(9, 'shrink')).toBeNull();
    expect(nextSpreadsheetFontSize(72, 'grow')).toBeNull();
    expect(nextSpreadsheetFontSize(100, 'shrink')).toBe(72);
    expect(nextSpreadsheetFontSize(500, 'shrink')).toBeNull();
  });

  test('compacts mixed per-cell sizes into deterministic native rectangles', () => {
    const cells: (Cell | null)[][] = [
      [{ fs: 9 }, { fs: 10 }, null],
      [{ fs: 9 }, { fs: 10 }, { fs: 72 }],
    ];

    expect(
      spreadsheetFontSizeApiCalls(
        cells,
        { row: [0, 1], column: [0, 2] },
        'sheet-1',
        'grow',
      ),
    ).toEqual([
      {
        name: 'setCellFormatByRange',
        args: ['fs', 10, { row: [0, 1], column: [0, 0] }, { id: 'sheet-1' }],
      },
      {
        name: 'setCellFormatByRange',
        args: ['fs', 11, { row: [0, 0], column: [1, 2] }, { id: 'sheet-1' }],
      },
      {
        name: 'setCellFormatByRange',
        args: ['fs', 11, { row: [1, 1], column: [1, 1] }, { id: 'sheet-1' }],
      },
    ]);
  });

  test('reads sparse content without densifying it and enforces the bound', () => {
    const content: WorkSpreadsheetContent = {
      type: 'spreadsheet',
      sheets: [
        {
          id: 'sheet-1',
          name: 'Sheet 1',
          row: 1_048_576,
          column: 16_384,
          celldata: [{ r: 999_999, c: 2, v: { fs: 14, v: 'A3S' } }],
        },
      ],
    };

    expect(
      canAdjustSpreadsheetFontSize(
        content,
        'sheet-1',
        { row: [999_999, 999_999], column: [2, 2] },
        'grow',
      ),
    ).toBe(true);
    expect(
      canAdjustSpreadsheetFontSize(
        content,
        'sheet-1',
        {
          row: [0, MAX_SPREADSHEET_FONT_SIZE_CELLS],
          column: [0, 0],
        },
        'grow',
      ),
    ).toBe(false);
    expect(content.sheets[0]?.celldata).toHaveLength(1);
  });
});
