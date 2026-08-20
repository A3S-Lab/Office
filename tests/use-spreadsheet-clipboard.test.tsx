import { beforeEach, describe, expect, test } from '@rstest/core';
import { act, renderHook, waitFor } from '@testing-library/react';
import {
  clearRichSpreadsheetClipboard,
  spreadsheetClipboardSnapshotForText,
} from '../src/internal/features/work/editors/spreadsheet-clipboard';
import { useSpreadsheetClipboard } from '../src/internal/features/work/editors/use-spreadsheet-clipboard';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

beforeEach(() => clearRichSpreadsheetClipboard());

describe('spreadsheet clipboard controller', () => {
  test('preserves rich same-editor cells and commits one controlled paste', async () => {
    const contentRef = { current: workbook() };
    const changes: WorkSpreadsheetContent[] = [];
    const writes: string[] = [];
    let selection = sourceSelection();
    const { result } = renderHook(() =>
      useSpreadsheetClipboard({
        canAccessSelection: true,
        clipboard: {
          readText: async () => writes.at(-1) ?? '',
          writeText: async (value) => {
            writes.push(value);
          },
        },
        clearSelection: () => true,
        commit: (next) => {
          changes.push(next);
          contentRef.current = next;
          return true;
        },
        contentRef,
        editable: true,
        fallbackFocusTarget: () => null,
        getSelection: () => selection,
      }),
    );

    act(() => expect(result.current.commandPort.copySelection()).toBe(true));
    await waitFor(() => expect(writes).toEqual(['5']));
    expect(spreadsheetClipboardSnapshotForText('5')?.kind).toBe('rich');

    selection = targetSelection(2, 2);
    act(() => expect(result.current.commandPort.pasteSelection()).toBe(true));
    await waitFor(() => expect(changes).toHaveLength(1));

    expect(changes[0]?.sheets[0]?.data?.[2]?.[2]).toMatchObject({
      v: 5,
      m: '5',
      bg: '#fff2cc',
      fs: 15,
    });
  });

  test('uses external TSV as plain cells when it does not match the rich copy', async () => {
    const contentRef = { current: workbook() };
    const changes: WorkSpreadsheetContent[] = [];
    let selection = sourceSelection();
    let clipboardText = '5';
    const { result } = renderHook(() =>
      useSpreadsheetClipboard({
        canAccessSelection: true,
        clipboard: {
          readText: async () => clipboardText,
          writeText: async (value) => {
            clipboardText = value;
          },
        },
        clearSelection: () => true,
        commit: (next) => {
          changes.push(next);
          contentRef.current = next;
          return true;
        },
        contentRef,
        editable: true,
        fallbackFocusTarget: () => null,
        getSelection: () => selection,
      }),
    );

    act(() => expect(result.current.commandPort.copySelection()).toBe(true));
    clipboardText = '42\tNorth';
    selection = targetSelection(2, 2);
    act(() => expect(result.current.commandPort.pasteSelection()).toBe(true));
    await waitFor(() => expect(changes).toHaveLength(1));

    expect(changes[0]?.sheets[0]?.data?.[2]?.[2]).toEqual({
      v: '42',
      m: '42',
    });
    expect(changes[0]?.sheets[0]?.data?.[2]?.[3]).toEqual({
      v: 'North',
      m: 'North',
    });
  });

  test('captures the target before asynchronous clipboard access resolves', async () => {
    const contentRef = { current: workbook() };
    const changes: WorkSpreadsheetContent[] = [];
    let selection = targetSelection(2, 2);
    let resolveClipboard: ((value: string) => void) | undefined;
    const { result } = renderHook(() =>
      useSpreadsheetClipboard({
        canAccessSelection: true,
        clipboard: {
          readText: () =>
            new Promise<string>((resolve) => {
              resolveClipboard = resolve;
            }),
          writeText: async () => undefined,
        },
        clearSelection: () => true,
        commit: (next) => {
          changes.push(next);
          contentRef.current = next;
          return true;
        },
        contentRef,
        editable: true,
        fallbackFocusTarget: () => null,
        getSelection: () => selection,
      }),
    );

    act(() => expect(result.current.commandPort.pasteSelection()).toBe(true));
    selection = targetSelection(4, 4);
    await act(async () => resolveClipboard?.('9'));
    await waitFor(() => expect(changes).toHaveLength(1));

    expect(changes[0]?.sheets[0]?.data?.[2]?.[2]?.v).toBe('9');
    expect(changes[0]?.sheets[0]?.data?.[4]?.[4]?.v).toBe(100);
  });

  test('opens Paste Special with the invoking target and exact focus origin', async () => {
    const contentRef = { current: workbook() };
    const invoker = document.createElement('button');
    const grid = document.createElement('button');
    document.body.append(invoker);
    document.body.append(grid);
    invoker.focus();
    const selection = targetSelection(2, 2);
    const { result } = renderHook(() =>
      useSpreadsheetClipboard({
        canAccessSelection: true,
        clipboard: {
          readText: async () => '8',
          writeText: async () => undefined,
        },
        clearSelection: () => true,
        commit: () => true,
        contentRef,
        editable: true,
        fallbackFocusTarget: () => grid,
        getSelection: () => selection,
      }),
    );

    act(() => expect(result.current.commandPort.openPasteSpecial()).toBe(true));
    await waitFor(() => expect(result.current.dialogSource).not.toBeNull());
    expect(result.current.dialogSource).toMatchObject({
      targetSheetId: 'sheet-1',
      targetRange: { row: [2, 2], column: [2, 2] },
      invoker,
    });
    expect(result.current.restoreDialogFocusTarget()).toBe(invoker);

    act(() =>
      expect(
        result.current.applyDialog({
          content: 'values',
          operation: 'none',
          skipBlanks: false,
          transpose: false,
        }),
      ).toBe(true),
    );
    await waitFor(() => expect(grid).toHaveFocus());

    invoker.remove();
    grid.remove();
  });
});

function workbook(): WorkSpreadsheetContent {
  return {
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-1',
        name: 'Sheet 1',
        row: 8,
        column: 8,
        config: {},
        data: Array.from({ length: 6 }, (_row, row) =>
          Array.from({ length: 6 }, (_column, column) =>
            row === 0 && column === 0
              ? { v: 5, m: '5', bg: '#fff2cc', fs: 15 }
              : row >= 2 && column >= 2
                ? { v: 100, m: '100', bg: '#ddebf7' }
                : null,
          ),
        ),
      },
    ],
  };
}

function sourceSelection() {
  return {
    sheetId: 'sheet-1',
    range: { row: [0, 0], column: [0, 0] } as const,
    plainText: '5',
  };
}

function targetSelection(row: number, column: number) {
  return {
    sheetId: 'sheet-1',
    range: {
      row: [row, row] as [number, number],
      column: [column, column] as [number, number],
    },
    plainText: '100',
  };
}
