import type { Cell } from '@fortune-sheet/core';
import { expect, test } from '@rstest/core';
import { act, fireEvent, render, screen, within } from '@testing-library/react';
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
        orientation: 'top-to-bottom',
        caseSensitive: false,
        textMethod: 'pinyin',
        hasHeader: true,
        keys: [{ index: 1, direction: 'ascending' }],
      },
    },
  ]);
  expect(
    portRef.current?.canApply(
      {
        sheetId: 'sheet-1',
        range: { row: [8, 9], column: [8, 9] },
        orientation: 'top-to-bottom',
        hasHeader: false,
        keys: [{ index: 8, direction: 'ascending' }],
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
        orientation: 'top-to-bottom',
        hasHeader: true,
        keys: [{ index: 1, direction: 'descending' }],
      },
    },
  ]);
  expect(screen.queryByRole('dialog', { name: '排序提醒' })).toBeNull();
  officeRoot.remove();
});

test('reuses authored custom lists without stealing focus after dialog close', async () => {
  const portRef: { current: SpreadsheetSortCommandPort | null } = {
    current: null,
  };
  const content = spreadsheetContent();
  const commandsRef = {
    current: {
      applyCustomSort: () => true,
    } as SpreadsheetEditorCommands,
  };
  const invoker = document.createElement('button');
  const nextTarget = document.createElement('button');
  document.body.append(invoker, nextTarget);
  invoker.focus();

  render(
    <SortHarness
      commandsRef={commandsRef}
      content={content}
      portRef={portRef}
    />,
  );

  act(() => expect(portRef.current?.open(customSelectionRequest())).toBe(true));
  const order = screen.getByRole('combobox', { name: '排序条件 1 次序' });
  const createOption = within(order).getByRole('option', {
    name: '新建自定义序列…',
  }) as HTMLOptionElement;
  fireEvent.change(order, { target: { value: createOption.value } });
  fireEvent.change(
    screen.getByRole('textbox', { name: '排序条件 1 自定义序列' }),
    { target: { value: '有风险\n进行中\n正常\n已完成' } },
  );
  fireEvent.click(screen.getByRole('button', { name: '使用序列' }));
  fireEvent.click(screen.getByRole('button', { name: '取消' }));

  expect(screen.queryByRole('dialog', { name: '自定义排序' })).toBeNull();
  expect(invoker).toHaveFocus();
  nextTarget.focus();
  await act(
    () =>
      new Promise<void>((resolve) => {
        requestAnimationFrame(() => resolve());
      }),
  );
  expect(nextTarget).toHaveFocus();
  act(() => expect(portRef.current?.open(customSelectionRequest())).toBe(true));
  expect(
    within(screen.getByRole('combobox', { name: '排序条件 1 次序' })).getByRole(
      'option',
      { name: '有风险 → 进行中 → 正常 → …' },
    ),
  ).toBeInTheDocument();
  fireEvent.click(screen.getByRole('button', { name: '取消' }));
  invoker.remove();
  nextTarget.remove();
});

test('authors an effective conditional-icon key from the controlled sheet snapshot', () => {
  const applied: SpreadsheetSortRequest[] = [];
  const portRef: { current: SpreadsheetSortCommandPort | null } = {
    current: null,
  };
  const content = spreadsheetContent();
  const commandsRef = {
    current: {
      applyCustomSort: (request: SpreadsheetSortRequest) => {
        applied.push(request);
        return true;
      },
    } as SpreadsheetEditorCommands,
  };

  render(
    <SortHarness
      commandsRef={commandsRef}
      content={content}
      portRef={portRef}
    />,
  );

  act(() => expect(portRef.current?.open(customSelectionRequest())).toBe(true));
  fireEvent.change(
    screen.getByRole('combobox', { name: '排序条件 1 排序依据' }),
    { target: { value: 'icon' } },
  );
  const target = screen.getByRole('combobox', {
    name: '排序条件 1 目标外观',
  });
  expect(
    within(target).getByRole('option', {
      name: /三色交通灯（实心） 3\/3/,
    }),
  ).toBeInTheDocument();
  fireEvent.change(target, {
    target: { value: 'icon:3TrafficLights1:2' },
  });
  fireEvent.click(screen.getByRole('button', { name: '确定' }));

  expect(applied).toEqual([
    {
      sheetId: 'sheet-1',
      range: { row: [0, 2], column: [0, 2] },
      orientation: 'top-to-bottom',
      caseSensitive: false,
      textMethod: 'pinyin',
      hasHeader: true,
      keys: [
        {
          index: 1,
          sortOn: 'icon',
          icon: { iconSet: '3TrafficLights1', index: 2 },
          position: 'first',
        },
      ],
    },
  ]);
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
    activeRow: 1,
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
    activeRow: 1,
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

function customSelectionRequest(): SpreadsheetSortOpenRequest {
  return {
    sheetId: 'sheet-1',
    activeColumn: 1,
    activeRow: 1,
    intent: { type: 'custom' },
    selected: {
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
        luckysheet_conditionformat_save: [
          {
            type: 'icons',
            cellrange: [{ row: [1, 2], column: [1, 1] }],
            format: {
              iconSet: '3TrafficLights1',
              showValue: true,
              reverse: false,
              percent: false,
              thresholds: [
                { type: 'min', gte: true },
                { type: 'num', value: 85, gte: true },
                { type: 'num', value: 88, gte: true },
              ],
            },
          },
        ],
      },
    ],
  };
}
