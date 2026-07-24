import type { Op } from '@fortune-sheet/core';
import { expect, test } from '@rstest/core';
import { act, renderHook, waitFor } from '@testing-library/react';
import { useSpreadsheetWorkbookSync } from '../src/internal/features/work/editors/use-spreadsheet-workbook-sync';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

test('remounts only for external controlled Spreadsheet values', async () => {
  const initial = workbook(1);
  const { result, rerender } = renderHook(
    ({ content }) => useSpreadsheetWorkbookSync(content),
    { initialProps: { content: initial } },
  );
  expect(result.current.mountRevision).toBe(0);

  const local = workbook(2);
  act(() => result.current.acceptContent(local));
  rerender({ content: local });
  expect(result.current.mountRevision).toBe(0);

  const external = workbook(3);
  rerender({ content: external });
  await waitFor(() => expect(result.current.mountRevision).toBe(1));

  const liveButRejected = workbook(4);
  act(() => result.current.acceptContent(liveButRejected));
  rerender({ content: external });
  await waitFor(() => expect(result.current.mountRevision).toBe(2));
});

test('preserves Fortune operation order until the matching workbook change', () => {
  const value = workbook(1);
  const { result } = renderHook(() => useSpreadsheetWorkbookSync(value));
  const operations: Op[] = [
    {
      id: 'sheet-1',
      op: 'replace',
      path: ['data', 0, 0, 'v'],
      value: 2,
    },
    {
      id: 'sheet-1',
      op: 'replace',
      path: ['data', 0, 0, 'm'],
      value: '2',
    },
  ];

  act(() => result.current.recordOperations(operations.slice(0, 1)));
  act(() => result.current.recordOperations(operations.slice(1)));

  expect(result.current.takeOperations()).toEqual(operations);
  expect(result.current.takeOperations()).toEqual([]);
});

function workbook(value: number): WorkSpreadsheetContent {
  return {
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-1',
        name: 'Sheet 1',
        data: [[{ v: value, m: String(value) }]],
      },
    ],
  };
}
