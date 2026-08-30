import type { Cell } from '@fortune-sheet/core';
import { expect, test } from '@rstest/core';
import { act, fireEvent, render, screen } from '@testing-library/react';
import { useRef } from 'react';
import type {
  SpreadsheetEditorCommands,
  SpreadsheetSortCommandPort,
} from '../src/internal/features/work/editors/spreadsheet-command-controller';
import type {
  SpreadsheetSortOpenRequest,
  SpreadsheetSortRequest,
} from '../src/internal/features/work/editors/spreadsheet-sort';
import { useSpreadsheetSort } from '../src/internal/features/work/editors/use-spreadsheet-sort';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

test('authenticates an expanded range before opening and applying Custom Sort', async () => {
  const applied: Array<{
    authorized: boolean;
    request: SpreadsheetSortRequest;
    staleSelectionAuthorized: boolean;
  }> = [];
  const portRef: { current: SpreadsheetSortCommandPort | null } = {
    current: null,
  };
  const content = spreadsheetContent();
  const commandsRef = {
    current: {
      applyCustomSort: (request: SpreadsheetSortRequest) => {
        applied.push({
          request,
          authorized:
            portRef.current?.canApply(request, {
              row: [1, 1],
              column: [1, 1],
            }) ?? false,
          staleSelectionAuthorized:
            portRef.current?.canApply(request, {
              row: [2, 2],
              column: [1, 1],
            }) ?? false,
        });
        return true;
      },
    } as SpreadsheetEditorCommands,
  };
  const officeRoot = document.createElement('div');
  officeRoot.dataset.a3sOffice = 'true';
  const invoker = document.createElement('button');
  const host = document.createElement('div');
  officeRoot.append(invoker, host);
  document.body.append(officeRoot);
  invoker.focus();

  render(
    <SortHarness
      commandsRef={commandsRef}
      content={content}
      portRef={portRef}
    />,
    { container: host },
  );

  act(() => expect(portRef.current?.open(expansionRequest())).toBe(true));
  expect(screen.getByRole('dialog', { name: '排序提醒' })).toBeVisible();
  fireEvent.click(screen.getByRole('button', { name: '排序' }));

  expect(applied).toEqual([]);
  expect(portRef.current?.canOpen).toBe(false);
  expect(
    await screen.findByRole('dialog', { name: '自定义排序' }),
  ).toHaveTextContent('执行看板!A1:C3');
  expect(screen.getByRole('checkbox', { name: '数据包含标题' })).toBeChecked();
  fireEvent.click(screen.getByRole('button', { name: '确定' }));

  expect(applied).toEqual([
    {
      authorized: true,
      staleSelectionAuthorized: false,
      request: {
        sheetId: 'sheet-1',
        range: { row: [0, 2], column: [0, 2] },
        hasHeader: true,
        keys: [{ column: 1, direction: 'ascending' }],
      },
    },
  ]);
  expect(
    portRef.current?.canApply(
      {
        sheetId: 'sheet-1',
        range: { row: [8, 9], column: [8, 9] },
        hasHeader: false,
        keys: [{ column: 8, direction: 'ascending' }],
      },
      { row: [1, 1], column: [1, 1] },
    ),
  ).toBe(false);
  officeRoot.remove();
});

test('applies a quick sort to the expanded current region with detected headers', () => {
  const applied: Array<{
    authorized: boolean;
    request: SpreadsheetSortRequest;
  }> = [];
  const portRef: { current: SpreadsheetSortCommandPort | null } = {
    current: null,
  };
  const content = spreadsheetContent();
  const commandsRef = {
    current: {
      applyCustomSort: (request: SpreadsheetSortRequest) => {
        applied.push({
          request,
          authorized:
            portRef.current?.canApply(request, {
              row: [1, 2],
              column: [1, 1],
            }) ?? false,
        });
        return true;
      },
    } as SpreadsheetEditorCommands,
  };
  const officeRoot = document.createElement('div');
  officeRoot.dataset.a3sOffice = 'true';
  const invoker = document.createElement('button');
  const host = document.createElement('div');
  officeRoot.append(invoker, host);
  document.body.append(officeRoot);
  invoker.focus();

  render(
    <SortHarness
      commandsRef={commandsRef}
      content={content}
      portRef={portRef}
    />,
    { container: host },
  );

  act(() => expect(portRef.current?.open(quickExpansionRequest())).toBe(true));
  expect(screen.getByRole('radio', { name: /扩展选定区域/ })).toBeChecked();
  fireEvent.click(screen.getByRole('button', { name: '排序' }));

  expect(applied).toEqual([
    {
      authorized: true,
      request: {
        sheetId: 'sheet-1',
        range: { row: [0, 2], column: [0, 2] },
        hasHeader: true,
        keys: [{ column: 1, direction: 'descending' }],
      },
    },
  ]);
  expect(screen.queryByRole('dialog', { name: '排序提醒' })).toBeNull();
  officeRoot.remove();
});

function SortHarness({
  commandsRef,
  content,
  portRef,
}: {
  commandsRef: { current: SpreadsheetEditorCommands | null };
  content: WorkSpreadsheetContent;
  portRef: { current: SpreadsheetSortCommandPort | null };
}) {
  const contentRef = useRef(content);
  const sort = useSpreadsheetSort({
    commandsRef,
    contentRef,
    focusGrid: () => undefined,
    getGridFocusTarget: () => null,
    getRows: ({ range, sheetId }) => {
      const sheet = contentRef.current.sheets.find(
        (candidate) => candidate.id === sheetId,
      );
      if (!sheet?.data) return null;
      const rows: (Cell | null)[][] = [];
      for (let row = range.row[0]; row <= range.row[1]; row += 1) {
        const cells: (Cell | null)[] = [];
        for (
          let column = range.column[0];
          column <= range.column[1];
          column += 1
        ) {
          cells.push(sheet.data[row]?.[column] ?? null);
        }
        rows.push(cells);
      }
      return rows;
    },
    preview: false,
  });
  portRef.current = sort.commandPort;
  return sort.dialog;
}

function expansionRequest(): SpreadsheetSortOpenRequest {
  return {
    sheetId: 'sheet-1',
    activeColumn: 1,
    intent: { type: 'custom' },
    selected: {
      range: { row: [1, 1], column: [1, 1] },
      available: false,
    },
    expanded: {
      range: { row: [0, 2], column: [0, 2] },
      available: true,
    },
  };
}

function quickExpansionRequest(): SpreadsheetSortOpenRequest {
  return {
    sheetId: 'sheet-1',
    activeColumn: 1,
    intent: { type: 'quick', direction: 'descending' },
    selected: {
      range: { row: [1, 2], column: [1, 1] },
      available: true,
    },
    expanded: {
      range: { row: [0, 2], column: [0, 2] },
      available: true,
    },
  };
}

function spreadsheetContent(): WorkSpreadsheetContent {
  return {
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-1',
        name: '执行看板',
        data: [
          [{ v: 'Name' }, { v: 'Score' }, { v: 'Owner' }],
          [{ v: 'Beta' }, { v: 80 }, { v: 'Lin' }],
          [{ v: 'Alpha' }, { v: 90 }, { v: 'Ada' }],
        ],
      },
    ],
  };
}
