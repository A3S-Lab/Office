import { expect, test } from '@rstest/core';
import { withEditableRange } from '../src/internal/features/work/work-spreadsheet-protection';

test('fills sparse worksheet rows when an editable range starts below row one', () => {
  const sheet = withEditableRange(
    {
      id: 'sheet-1',
      name: '工作表 1',
      row: 30,
      column: 10,
      data: [],
    },
    null,
    { name: 'InputCells', sqref: 'B2:B10' },
  );

  expect(sheet.data).toHaveLength(30);
  expect(sheet.data?.every((row) => Array.isArray(row))).toBe(true);
  expect(sheet.data?.every((row) => row.length === 10)).toBe(true);
  expect(sheet.data?.[1]?.[1]).toMatchObject({ lo: 0 });
});
