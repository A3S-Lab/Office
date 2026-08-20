import type { Cell } from '@fortune-sheet/core';
import { describe, expect, test } from '@rstest/core';
import {
  canAdjustSpreadsheetDecimalPlaces,
  MAX_SPREADSHEET_DECIMAL_FORMAT_CELLS,
  spreadsheetDecimalFormatApiCalls,
} from '../src/internal/features/work/editors/spreadsheet-number-format-command';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

describe('spreadsheet decimal-place commands', () => {
  test('preserves each number-format family and compacts vertical runs', () => {
    const cells: (Cell | null)[][] = [
      [
        { ct: { fa: '[$¥-804]#,##0.00', t: 'n' }, v: 12.5 },
        { ct: { fa: '0.0%', t: 'n' }, v: 0.25 },
        { ct: { fa: 'yyyy-MM-dd', t: 'd' }, v: 45_292 },
      ],
      [
        { ct: { fa: '[$¥-804]#,##0.00', t: 'n' }, v: 20.5 },
        { ct: { fa: '0.0%', t: 'n' }, v: 0.5 },
        { ct: { fa: '# ?/?', t: 'n' }, v: 1.5 },
      ],
    ];

    expect(
      spreadsheetDecimalFormatApiCalls(
        cells,
        { row: [4, 5], column: [2, 4] },
        'sheet-1',
        'increase',
      ),
    ).toEqual([
      {
        name: 'setCellFormatByRange',
        args: [
          'ct',
          { fa: '[$¥-804]#,##0.000', t: 'n' },
          { row: [4, 5], column: [2, 2] },
          { id: 'sheet-1' },
        ],
      },
      {
        name: 'setCellFormatByRange',
        args: [
          'ct',
          { fa: '0.00%', t: 'n' },
          { row: [4, 5], column: [3, 3] },
          { id: 'sheet-1' },
        ],
      },
    ]);
  });

  test('bounds sparse materialization and rejects a complete no-op', () => {
    const content: WorkSpreadsheetContent = {
      type: 'spreadsheet',
      sheets: [
        {
          column: 1,
          data: [[{ ct: { fa: 'yyyy-MM-dd', t: 'd' }, v: 45_292 }]],
          id: 'sheet-1',
          name: 'Sheet 1',
          row: 1,
        },
      ],
    };

    expect(
      canAdjustSpreadsheetDecimalPlaces(
        content,
        'sheet-1',
        { row: [0, 0], column: [0, 0] },
        'increase',
      ),
    ).toBe(false);
    expect(
      canAdjustSpreadsheetDecimalPlaces(
        content,
        'sheet-1',
        {
          row: [0, 0],
          column: [0, MAX_SPREADSHEET_DECIMAL_FORMAT_CELLS],
        },
        'increase',
      ),
    ).toBe(false);
    expect(
      canAdjustSpreadsheetDecimalPlaces(
        content,
        'sheet-1',
        { row: [1, 1], column: [0, 0] },
        'decrease',
      ),
    ).toBe(true);
  });

  test('reads sparse celldata when deciding whether a selection can change', () => {
    const content: WorkSpreadsheetContent = {
      type: 'spreadsheet',
      sheets: [
        {
          celldata: [
            {
              c: 0,
              r: 0,
              v: { ct: { fa: 'yyyy-MM-dd', t: 'd' }, v: 45_292 },
            },
            {
              c: 1,
              r: 0,
              v: { ct: { fa: '0.00%', t: 'n' }, v: 0.25 },
            },
          ],
          column: 2,
          id: 'sheet-1',
          name: 'Sheet 1',
          row: 1,
        },
      ],
    };

    expect(
      canAdjustSpreadsheetDecimalPlaces(
        content,
        'sheet-1',
        { row: [0, 0], column: [0, 0] },
        'increase',
      ),
    ).toBe(false);
    expect(
      canAdjustSpreadsheetDecimalPlaces(
        content,
        'sheet-1',
        { row: [0, 0], column: [1, 1] },
        'increase',
      ),
    ).toBe(true);
  });
});
