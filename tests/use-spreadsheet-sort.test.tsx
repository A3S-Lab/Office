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
import type { SpreadsheetSortCustomListStore } from '../src/internal/features/work/editors/spreadsheet-sort-custom-list-store';
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

test('carries an owned table scope from the warning into locked sort controls', async () => {
  const applied: SpreadsheetSortRequest[] = [];
  const portRef: { current: SpreadsheetSortCommandPort | null } = {
    current: null,
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
      commandsRef={{
        current: {
          applyCustomSort: (request: SpreadsheetSortRequest) => {
            applied.push(request);
            return true;
          },
        } as SpreadsheetEditorCommands,
      }}
      content={spreadsheetContent()}
      portRef={portRef}
    />,
    { container: host },
  );
  const request = expansionRequest();
  if (!request.expanded) throw new Error('Expected an expanded fixture.');
  request.expanded.scope = {
    kind: 'table',
    tableId: 'table-1',
    hasHeader: true,
  };

  act(() => expect(portRef.current?.open(request)).toBe(true));
  expect(screen.getByRole('dialog', { name: '排序提醒' })).toHaveTextContent(
    '选定单元格位于表格中',
  );
  fireEvent.click(screen.getByRole('button', { name: '排序' }));

  expect(
    await screen.findByRole('dialog', { name: '自定义排序' }),
  ).toBeVisible();
  expect(screen.getByRole('checkbox', { name: '数据包含标题' })).toBeDisabled();
  fireEvent.click(screen.getByRole('button', { name: '选项…' }));
  const options = await screen.findByRole('dialog', { name: '排序选项' });
  expect(
    within(options).getByRole('radio', { name: /按行排序/ }),
  ).toBeDisabled();
  expect(options).toHaveTextContent('结构化数据区域仅支持按列排序');
  fireEvent.click(within(options).getByRole('button', { name: '取消' }));
  fireEvent.click(screen.getByRole('button', { name: '确定' }));

  expect(applied).toEqual([
    expect.objectContaining({
      sheetId: 'sheet-1',
      range: { row: [0, 2], column: [0, 2] },
      orientation: 'top-to-bottom',
      hasHeader: true,
      scope: { kind: 'table', tableId: 'table-1', hasHeader: true },
    }),
  ]);
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

test('persists authored custom lists through a typed host store and remount', () => {
  const portRef: { current: SpreadsheetSortCommandPort | null } = {
    current: null,
  };
  const content = spreadsheetContent();
  const commandsRef = {
    current: {
      applyCustomSort: () => true,
    } as SpreadsheetEditorCommands,
  };
  const store = new RecordingCustomListStore();
  const first = render(
    <SortHarness
      commandsRef={commandsRef}
      content={content}
      customListStore={store}
      portRef={portRef}
    />,
  );
  expect(store.loads).toBe(1);

  act(() => expect(portRef.current?.open(customSelectionRequest())).toBe(true));
  const order = screen.getByRole('combobox', { name: '排序条件 1 次序' });
  fireEvent.change(order, { target: { value: 'create-custom-list' } });
  fireEvent.change(
    screen.getByRole('textbox', { name: '排序条件 1 自定义序列' }),
    { target: { value: '有风险\n进行中\n正常\n已完成' } },
  );
  fireEvent.click(screen.getByRole('button', { name: '使用序列' }));
  expect(store.saved).toEqual([[['有风险', '进行中', '正常', '已完成']]]);
  expect(
    within(order).getByRole('option', {
      name: '有风险 → 进行中 → 正常 → …',
    }),
  ).toBeInTheDocument();
  expect(order.querySelector('optgroup[label="已保存的序列"]')).not.toBeNull();
  first.unmount();

  render(
    <SortHarness
      commandsRef={commandsRef}
      content={content}
      customListStore={store}
      portRef={portRef}
    />,
  );
  expect(store.loads).toBe(2);
  act(() => expect(portRef.current?.open(customSelectionRequest())).toBe(true));
  const remountedOrder = screen.getByRole('combobox', {
    name: '排序条件 1 次序',
  });
  expect(
    within(remountedOrder).getByRole('option', {
      name: '有风险 → 进行中 → 正常 → …',
    }),
  ).toBeInTheDocument();
  expect(
    remountedOrder.querySelector('optgroup[label="已保存的序列"]'),
  ).not.toBeNull();
});

test('persists custom-list deletion and preference order across remounts', () => {
  const portRef: { current: SpreadsheetSortCommandPort | null } = {
    current: null,
  };
  const content = spreadsheetContent();
  const commandsRef = {
    current: {
      applyCustomSort: () => true,
    } as SpreadsheetEditorCommands,
  };
  const store = new RecordingCustomListStore();
  store.lists = [
    ['High', 'Medium', 'Low'],
    ['North', 'South'],
    ['Red', 'Amber', 'Green'],
  ];
  const first = render(
    <SortHarness
      commandsRef={commandsRef}
      content={content}
      customListStore={store}
      portRef={portRef}
    />,
  );

  act(() => expect(portRef.current?.open(customSelectionRequest())).toBe(true));
  const managerButton = screen.getByRole('button', {
    name: '管理自定义序列',
  });
  fireEvent.click(managerButton);
  let manager = screen.getByRole('dialog', { name: '自定义序列' });
  fireEvent.click(within(manager).getByRole('button', { name: '确定' }));
  expect(store.saved).toEqual([]);

  fireEvent.click(managerButton);
  manager = screen.getByRole('dialog', { name: '自定义序列' });
  const list = within(manager).getByRole('listbox', {
    name: '自定义序列列表',
  });
  fireEvent.change(list, { target: { value: 'user:2' } });
  fireEvent.click(within(manager).getByRole('button', { name: '上移序列' }));
  fireEvent.click(within(manager).getByRole('button', { name: '上移序列' }));
  fireEvent.change(list, { target: { value: 'user:1' } });
  fireEvent.click(within(manager).getByRole('button', { name: '删除序列' }));
  fireEvent.click(within(manager).getByRole('button', { name: '确定' }));

  expect(store.saved).toEqual([
    [
      ['Red', 'Amber', 'Green'],
      ['High', 'Medium', 'Low'],
    ],
  ]);
  fireEvent.click(screen.getByRole('button', { name: '取消' }));
  first.unmount();

  render(
    <SortHarness
      commandsRef={commandsRef}
      content={content}
      customListStore={store}
      portRef={portRef}
    />,
  );
  act(() => expect(portRef.current?.open(customSelectionRequest())).toBe(true));
  const order = screen.getByRole('combobox', { name: '排序条件 1 次序' });
  expect(
    within(order.querySelector('optgroup[label="已保存的序列"]') as HTMLElement)
      .getAllByRole('option')
      .map((option) => option.textContent),
  ).toEqual(['Red → Amber → Green', 'High → Medium → Low']);
});

test('keeps an authored list in session when the host store rejects a write', () => {
  const portRef: { current: SpreadsheetSortCommandPort | null } = {
    current: null,
  };
  const content = spreadsheetContent();
  const commandsRef = {
    current: {
      applyCustomSort: () => true,
    } as SpreadsheetEditorCommands,
  };
  const store: SpreadsheetSortCustomListStore = {
    load: () => [],
    save: () => {
      throw new Error('storage unavailable');
    },
  };

  render(
    <SortHarness
      commandsRef={commandsRef}
      content={content}
      customListStore={store}
      portRef={portRef}
    />,
  );

  act(() => expect(portRef.current?.open(customSelectionRequest())).toBe(true));
  const order = screen.getByRole('combobox', { name: '排序条件 1 次序' });
  fireEvent.change(order, { target: { value: 'create-custom-list' } });
  fireEvent.change(
    screen.getByRole('textbox', { name: '排序条件 1 自定义序列' }),
    { target: { value: '高\n中\n低' } },
  );
  fireEvent.click(screen.getByRole('button', { name: '使用序列' }));

  expect(
    order.querySelector('optgroup[label="本次会话的序列"]'),
  ).not.toBeNull();
  expect(order.querySelector('optgroup[label="已保存的序列"]')).toBeNull();
  fireEvent.click(screen.getByRole('button', { name: '取消' }));

  act(() => expect(portRef.current?.open(customSelectionRequest())).toBe(true));
  const reopenedOrder = screen.getByRole('combobox', {
    name: '排序条件 1 次序',
  });
  expect(
    within(reopenedOrder).getByRole('option', { name: '高 → 中 → 低' }),
  ).toBeInTheDocument();
  expect(
    reopenedOrder.querySelector('optgroup[label="本次会话的序列"]'),
  ).not.toBeNull();
});

test('keeps the complete managed preference set in session after a rejected write', () => {
  const portRef: { current: SpreadsheetSortCommandPort | null } = {
    current: null,
  };
  const content = spreadsheetContent();
  const commandsRef = {
    current: {
      applyCustomSort: () => true,
    } as SpreadsheetEditorCommands,
  };
  const store: SpreadsheetSortCustomListStore = {
    load: () => [
      ['High', 'Medium', 'Low'],
      ['North', 'South'],
    ],
    save: () => {
      throw new Error('storage unavailable');
    },
  };

  render(
    <SortHarness
      commandsRef={commandsRef}
      content={content}
      customListStore={store}
      portRef={portRef}
    />,
  );

  act(() => expect(portRef.current?.open(customSelectionRequest())).toBe(true));
  fireEvent.click(screen.getByRole('button', { name: '管理自定义序列' }));
  const manager = screen.getByRole('dialog', { name: '自定义序列' });
  const list = within(manager).getByRole('listbox', {
    name: '自定义序列列表',
  });
  fireEvent.change(list, { target: { value: 'user:0' } });
  fireEvent.change(
    within(manager).getByRole('textbox', { name: '自定义序列项目' }),
    { target: { value: 'Critical\nNormal' } },
  );
  fireEvent.click(within(manager).getByRole('button', { name: '保存更改' }));
  fireEvent.change(list, { target: { value: 'user:1' } });
  fireEvent.click(within(manager).getByRole('button', { name: '删除序列' }));
  fireEvent.click(within(manager).getByRole('button', { name: '确定' }));

  const order = screen.getByRole('combobox', { name: '排序条件 1 次序' });
  expect(order.querySelector('optgroup[label="已保存的序列"]')).toBeNull();
  expect(
    within(
      order.querySelector('optgroup[label="本次会话的序列"]') as HTMLElement,
    )
      .getAllByRole('option')
      .map((option) => option.textContent),
  ).toEqual(['Critical → Normal']);

  fireEvent.click(screen.getByRole('button', { name: '取消' }));
  act(() => expect(portRef.current?.open(customSelectionRequest())).toBe(true));
  expect(
    within(
      screen
        .getByRole('combobox', { name: '排序条件 1 次序' })
        .querySelector('optgroup[label="本次会话的序列"]') as HTMLElement,
    ).getByRole('option', { name: 'Critical → Normal' }),
  ).toBeInTheDocument();
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
  customListStore,
  portRef,
}: {
  commandsRef: { current: SpreadsheetEditorCommands | null };
  content: WorkSpreadsheetContent;
  customListStore?: SpreadsheetSortCustomListStore;
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
    customListStore,
  });
  portRef.current = sort.commandPort;
  return sort.dialog;
}

class RecordingCustomListStore implements SpreadsheetSortCustomListStore {
  lists: readonly (readonly string[])[] = [];
  loads = 0;
  readonly saved: Array<readonly (readonly string[])[]> = [];

  load(): readonly (readonly string[])[] {
    this.loads += 1;
    return this.lists;
  }

  save(lists: readonly (readonly string[])[]): void {
    this.lists = lists.map((list) => [...list]);
    this.saved.push(this.lists);
  }
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
