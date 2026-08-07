import type { Cell } from '@fortune-sheet/core';
import { expect, test } from '@rstest/core';
import { act, renderHook } from '@testing-library/react';
import type { SpreadsheetWorkbookCommandPort } from '../src/internal/features/work/editors/spreadsheet-command-controller';
import {
  spreadsheetFormatPainterMaximumCells,
  useSpreadsheetFormatPainter,
} from '../src/internal/features/work/editors/use-spreadsheet-format-painter';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

test('applies once, locks for repeated targets, and cancels explicitly', () => {
  const workbook = new FormatPainterWorkbook();
  workbook.cells = [
    [
      { bg: '#fff2cc', bl: 1 },
      { bg: '#d9eaf7', bl: 0 },
    ],
  ];
  const errors: string[] = [];
  const { result } = renderHook(() =>
    useSpreadsheetFormatPainter({
      content: workbookContent(),
      editable: true,
      onError: (message) => errors.push(message),
      sourceRange: { row: [0, 0], column: [0, 1] },
      sourceSheetId: 'sheet-1',
      workbook,
    }),
  );

  act(() => {
    expect(result.current.commandPort.activate('once')).toBe(true);
  });
  expect(result.current.mode).toBe('once');
  expect(workbook.reads).toEqual([
    { range: { row: [0, 0], column: [0, 1] }, sheetId: 'sheet-1' },
  ]);

  act(() => {
    expect(
      result.current.commandPort.applySelection({
        sheetId: 'sheet-2',
        selection: { row: [2, 2], column: [1, 1] },
      }),
    ).toBe(true);
  });
  expect(result.current.mode).toBeNull();
  expect(workbook.batches).toHaveLength(1);
  expect(workbook.batches[0]).toContainEqual({
    name: 'setCellFormatByRange',
    args: [
      'bg',
      '#fff2cc',
      [{ row: [2, 2], column: [1, 1] }],
      { id: 'sheet-2' },
    ],
  });
  expect(workbook.batches[0]).toContainEqual({
    name: 'setCellFormatByRange',
    args: [
      'bg',
      '#d9eaf7',
      [{ row: [2, 2], column: [2, 2] }],
      { id: 'sheet-2' },
    ],
  });

  act(() => {
    expect(result.current.commandPort.activate('locked')).toBe(true);
  });
  expect(result.current.mode).toBe('locked');
  act(() => {
    expect(
      result.current.commandPort.applySelection({
        sheetId: 'sheet-1',
        selection: { row: [4, 4], column: [0, 0] },
      }),
    ).toBe(true);
  });
  expect(result.current.mode).toBe('locked');
  act(() => {
    expect(
      result.current.commandPort.applySelection({
        sheetId: 'sheet-1',
        selection: { row: [4, 4], column: [0, 0] },
      }),
    ).toBe(false);
  });
  expect(workbook.batches).toHaveLength(2);
  act(() => {
    expect(
      result.current.commandPort.applySelection({
        sheetId: 'sheet-1',
        selection: { row: [5, 5], column: [0, 0] },
      }),
    ).toBe(true);
  });
  expect(workbook.batches).toHaveLength(3);
  act(() => {
    expect(result.current.commandPort.cancel()).toBe(true);
  });
  expect(result.current.mode).toBeNull();
  expect(errors).toEqual([]);
});

test('rejects oversized source and target ranges without freezing the workbook', () => {
  const workbook = new FormatPainterWorkbook();
  workbook.cells = [[{ bg: '#fff2cc' }]];
  const errors: string[] = [];
  const { result, rerender } = renderHook(
    ({ sourceRange }) =>
      useSpreadsheetFormatPainter({
        content: workbookContent(spreadsheetFormatPainterMaximumCells + 10, 4),
        editable: true,
        onError: (message) => errors.push(message),
        sourceRange,
        sourceSheetId: 'sheet-1',
        workbook,
      }),
    {
      initialProps: {
        sourceRange: {
          row: [0, spreadsheetFormatPainterMaximumCells],
          column: [0, 0],
        },
      },
    },
  );

  act(() => {
    expect(result.current.commandPort.activate('once')).toBe(false);
  });
  expect(workbook.reads).toEqual([]);
  expect(errors.at(-1)).toContain('请缩小源区域');

  rerender({ sourceRange: { row: [0, 0], column: [0, 0] } });
  act(() => {
    expect(result.current.commandPort.activate('once')).toBe(true);
    expect(
      result.current.commandPort.applySelection({
        sheetId: 'sheet-1',
        selection: {
          row: [0, spreadsheetFormatPainterMaximumCells],
          column: [0, 0],
        },
      }),
    ).toBe(false);
  });
  expect(workbook.batches).toEqual([]);
  expect(result.current.mode).toBe('once');
  expect(errors.at(-1)).toContain('请缩小目标区域');
});

class FormatPainterWorkbook {
  cells: (Cell | null)[][] = [];
  batches: Array<Array<{ name: string; args: unknown[] }>> = [];
  reads: Array<{ range: unknown; sheetId: string | undefined }> = [];

  batchCallApis(apiCalls: Array<{ name: string; args: unknown[] }>): void {
    this.batches.push(apiCalls);
  }

  getCellsByRange(
    range: Parameters<SpreadsheetWorkbookCommandPort['getCellsByRange']>[0],
    options?: { id?: string },
  ): (Cell | null)[][] {
    this.reads.push({ range, sheetId: options?.id });
    return this.cells;
  }
}

function workbookContent(row = 20, column = 12): WorkSpreadsheetContent {
  return {
    type: 'spreadsheet',
    sheets: [
      { id: 'sheet-1', name: 'Sheet 1', row, column },
      { id: 'sheet-2', name: 'Sheet 2', row, column },
    ],
  };
}
