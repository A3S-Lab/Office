import { describe, expect, test } from '@rstest/core';
import {
  canApplySpreadsheetCellFill,
  planSpreadsheetCellFill,
  spreadsheetCellFillMaximumCells,
} from '../src/internal/features/work/editors/spreadsheet-cell-fill';
import type { WorkSpreadsheetSheet } from '../src/internal/features/work/work-types';

describe('spreadsheet cell fill', () => {
  test('plans edge-copy ranges for all four WPS fill directions', () => {
    const selection = { row: [4, 1], column: [3, 2] };

    expect(planSpreadsheetCellFill(selection, 'down')).toEqual({
      applyRange: { row: [2, 4], column: [2, 3] },
      copyRange: { row: [1, 1], column: [2, 3] },
      direction: 'down',
      targetCellCount: 6,
    });
    expect(planSpreadsheetCellFill(selection, 'up')).toEqual({
      applyRange: { row: [1, 3], column: [2, 3] },
      copyRange: { row: [4, 4], column: [2, 3] },
      direction: 'up',
      targetCellCount: 6,
    });
    expect(planSpreadsheetCellFill(selection, 'right')).toEqual({
      applyRange: { row: [1, 4], column: [3, 3] },
      copyRange: { row: [1, 4], column: [2, 2] },
      direction: 'right',
      targetCellCount: 4,
    });
    expect(planSpreadsheetCellFill(selection, 'left')).toEqual({
      applyRange: { row: [1, 4], column: [2, 2] },
      copyRange: { row: [1, 4], column: [3, 3] },
      direction: 'left',
      targetCellCount: 4,
    });
  });

  test('rejects one-cell axes, invalid ranges, and oversized fill targets', () => {
    expect(
      planSpreadsheetCellFill({ row: [0, 0], column: [0, 3] }, 'down'),
    ).toBeNull();
    expect(
      planSpreadsheetCellFill({ row: [0, 3], column: [0, 0] }, 'right'),
    ).toBeNull();
    expect(
      planSpreadsheetCellFill({ row: [-1, 2], column: [0, 1] }, 'down'),
    ).toBeNull();
    expect(
      planSpreadsheetCellFill(
        { row: [0, spreadsheetCellFillMaximumCells], column: [0, 1] },
        'down',
      ),
    ).toBeNull();
    expect(
      planSpreadsheetCellFill({ row: [0, 2], column: [0, 1] }, 'down', 4),
    ).not.toBeNull();
    expect(
      planSpreadsheetCellFill({ row: [0, 2], column: [0, 1] }, 'down', 3),
    ).toBeNull();
  });

  test('preflights merges, pivot sheets, and read-only axes', () => {
    const plan = planSpreadsheetCellFill(
      { row: [0, 2], column: [0, 1] },
      'down',
    );
    if (!plan) throw new Error('Expected a fill plan.');

    expect(canApplySpreadsheetCellFill(sheet(), plan)).toBe(true);
    expect(
      canApplySpreadsheetCellFill(
        sheet({ config: { merge: { '0_0': { r: 0, c: 0, rs: 2, cs: 2 } } } }),
        plan,
      ),
    ).toBe(false);
    expect(
      canApplySpreadsheetCellFill(sheet({ isPivotTable: true }), plan),
    ).toBe(false);
    expect(
      canApplySpreadsheetCellFill(
        sheet({ config: { rowReadOnly: { 1: 1 } } }),
        plan,
      ),
    ).toBe(false);
    expect(
      canApplySpreadsheetCellFill(
        sheet({ config: { colReadOnly: { 1: 1 } } }),
        plan,
      ),
    ).toBe(false);
  });

  test('matches Fortune protection semantics before invoking native fill', () => {
    const plan = planSpreadsheetCellFill(
      { row: [0, 1], column: [0, 1] },
      'down',
    );
    if (!plan) throw new Error('Expected a fill plan.');

    expect(
      canApplySpreadsheetCellFill(
        sheet({
          config: { authority: { sheet: 1 } },
          data: [
            [{ lo: 0 }, { lo: 0 }],
            [{ lo: 0 }, { lo: 0 }],
          ],
        }),
        plan,
      ),
    ).toBe(true);
    expect(
      canApplySpreadsheetCellFill(
        sheet({
          config: { authority: { sheet: 1 } },
          data: [[{ lo: 0 }, { lo: 0 }], [{ lo: 0 }]],
        }),
        plan,
      ),
    ).toBe(false);
    expect(
      canApplySpreadsheetCellFill(
        sheet({
          config: {
            authority: {
              cellProtectionRanges: [
                {
                  range: { row: [1, 1], column: [1, 1] },
                  locked: true,
                },
              ],
              sheet: 0,
            },
          },
        }),
        plan,
      ),
    ).toBe(false);
    expect(
      canApplySpreadsheetCellFill(
        sheet({
          data: [
            [{ lo: 1 }, null],
            [null, null],
          ],
        }),
        plan,
      ),
    ).toBe(false);
  });
});

function sheet(
  overrides: Partial<WorkSpreadsheetSheet> = {},
): WorkSpreadsheetSheet {
  return {
    id: 'sheet-1',
    name: 'Sheet 1',
    ...overrides,
  };
}
