import { checkCellIsLocked, type Context } from '@fortune-sheet/core';
import { expect, test } from '@rstest/core';
import {
  sheetProtectionAuthority,
  withEditableRange,
  withSheetProtection,
} from '../src/internal/features/work/work-spreadsheet-protection';

test('keeps editable ranges compact on sparse worksheets', () => {
  const sheet = withSheetProtection(
    withEditableRange(
      {
        id: 'sheet-1',
        name: 'Sheet 1',
        row: 1_048_576,
        column: 16_384,
        data: [],
      },
      null,
      { name: 'InputCells', sqref: 'B2:B1048576' },
    ),
    true,
  );
  const authority = sheetProtectionAuthority(sheet);
  const context = {
    currentSheetId: 'sheet-1',
    luckysheetfile: [sheet],
  } as unknown as Context;

  expect(sheet.data).toHaveLength(1_048_576);
  expect(Object.keys(sheet.data ?? [])).toEqual([]);
  expect(authority.cellProtectionRanges).toEqual([
    {
      range: { row: [1, 1_048_575], column: [1, 1] },
      locked: false,
      hidden: false,
    },
  ]);
  expect(checkCellIsLocked(context, 700_000, 1, 'sheet-1')).toBe(false);
  expect(checkCellIsLocked(context, 700_000, 2, 'sheet-1')).toBe(true);
});
