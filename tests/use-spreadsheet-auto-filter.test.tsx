import type { Selection } from '@fortune-sheet/core';
import { expect, test } from '@rstest/core';
import {
  act,
  fireEvent,
  render,
  renderHook,
  screen,
  waitFor,
} from '@testing-library/react';
import {
  type SpreadsheetAutoFilterController,
  useSpreadsheetAutoFilter,
} from '../src/internal/features/work/editors/use-spreadsheet-auto-filter';
import {
  enhanceSpreadsheetAutoFilterSurface,
  SPREADSHEET_AUTO_FILTER_MENU_ITEMS,
} from '../src/internal/features/work/editors/spreadsheet-auto-filter-menu';
import { applySpreadsheetAutoFilterCriteria } from '../src/internal/features/work/editors/spreadsheet-auto-filter';
import type { SpreadsheetEditorCommands } from '../src/internal/features/work/editors/spreadsheet-command-controller';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

const noSpreadsheetCommandsRef: {
  current: SpreadsheetEditorCommands | null;
} = { current: null };

test('publishes one controlled workbook change when AutoFilter is toggled', () => {
  const changes: WorkSpreadsheetContent[] = [];
  const canvas = document.createElement('div');
  const initial = spreadsheetContent(false);
  const { result, rerender } = renderHook(
    ({ content }) =>
      useSpreadsheetAutoFilter({
        canvasRef: { current: canvas },
        commandsRef: noSpreadsheetCommandsRef,
        content,
        editable: true,
        onChange: (next) => changes.push(next),
        selection: cellSelection(3, 1),
        sheetId: 'sheet-1',
      }),
    { initialProps: { content: initial } },
  );

  expect(result.current.active).toBe(false);
  expect(result.current.commandPort.canToggle).toBe(true);
  act(() => expect(result.current.commandPort.toggle()).toBe(true));
  expect(changes).toHaveLength(1);
  expect(changes[0]?.sheets[0]?.filter_select).toEqual({
    row: [2, 4],
    column: [0, 2],
  });

  rerender({ content: changes[0] as WorkSpreadsheetContent });
  expect(result.current.active).toBe(true);
});

test('keeps unsafe vendor row sorting out of the AutoFilter menu', () => {
  expect(SPREADSHEET_AUTO_FILTER_MENU_ITEMS).toContain('filter-by-condition');
  expect(SPREADSHEET_AUTO_FILTER_MENU_ITEMS).not.toContain('sort-by-asc');
  expect(SPREADSHEET_AUTO_FILTER_MENU_ITEMS).not.toContain('sort-by-desc');
});

test('adds the owned Top 10 action only for numeric filter columns', () => {
  const canvas = document.createElement('div');
  const triggers = Array.from({ length: 3 }, () => {
    const trigger = document.createElement('div');
    trigger.className = 'luckysheet-filter-options';
    trigger.tabIndex = 0;
    canvas.append(trigger);
    return trigger;
  });
  const menu = document.createElement('div');
  menu.className = 'fortune-filter-menu';
  const condition = document.createElement('div');
  condition.className = 'luckysheet-cols-menuitem';
  condition.tabIndex = 0;
  condition.textContent = '按条件过滤';
  menu.append(condition);
  canvas.append(menu);

  const content = spreadsheetContent(true);
  const sheet = content.sheets[0];
  if (!sheet) throw new Error('Expected the spreadsheet fixture.');
  enhanceSpreadsheetAutoFilterSurface(canvas, sheet, triggers[1] ?? null);
  const rank = canvas.querySelector<HTMLElement>('[data-a3s-auto-filter-rank]');
  expect(rank).not.toBeNull();
  expect(rank).toHaveTextContent('前 10 项');
  expect(rank).toHaveAttribute('role', 'button');
  expect(rank).toHaveAttribute('aria-haspopup', 'dialog');

  enhanceSpreadsheetAutoFilterSurface(canvas, sheet, triggers[2] ?? null);
  expect(canvas.querySelector('[data-a3s-auto-filter-rank]')).toBeNull();
});

test('opens the owned Top 10 action with a bounded default criterion', async () => {
  const canvas = document.createElement('div');
  const triggers = Array.from({ length: 3 }, () => {
    const trigger = document.createElement('div');
    trigger.className = 'luckysheet-filter-options';
    trigger.tabIndex = 0;
    canvas.append(trigger);
    return trigger;
  });
  const trigger = triggers[1] as HTMLElement;
  document.body.append(canvas);
  trigger.addEventListener('click', () => {
    if (canvas.querySelector('.fortune-filter-menu')) return;
    const menu = document.createElement('div');
    menu.className = 'fortune-filter-menu';
    const condition = document.createElement('div');
    condition.className = 'luckysheet-cols-menuitem';
    condition.tabIndex = 0;
    condition.textContent = '按条件过滤';
    const cancel = document.createElement('div');
    cancel.className = 'button-basic button-default';
    cancel.tabIndex = 0;
    cancel.textContent = '取消';
    cancel.addEventListener('click', () => menu.remove());
    menu.append(condition, cancel);
    canvas.append(menu);
  });
  const requests: Array<
    Parameters<SpreadsheetEditorCommands['applyAutoFilterCriteria']>[0]
  > = [];
  const commandsRef = {
    current: {
      applyAutoFilterCriteria: (request) => {
        requests.push(request);
        return true;
      },
    } as SpreadsheetEditorCommands,
  };
  const { result, unmount } = renderHook(() =>
    useSpreadsheetAutoFilter({
      canvasRef: { current: canvas },
      commandsRef,
      content: spreadsheetContent(true),
      editable: true,
      onChange: () => undefined,
      selection: cellSelection(2, 1),
      sheetId: 'sheet-1',
    }),
  );

  fireEvent.pointerDown(trigger);
  fireEvent.click(trigger);
  const rank = await waitFor(() => {
    const candidate = canvas.querySelector<HTMLElement>(
      '[data-a3s-auto-filter-rank]',
    );
    expect(candidate).not.toBeNull();
    return candidate as HTMLElement;
  });
  fireEvent.click(rank);
  await waitFor(() => expect(result.current.dialog).not.toBeNull());
  render(result.current.dialog);

  expect(screen.getByRole('combobox', { name: '筛选条件' })).toHaveValue('top');
  expect(screen.getByRole('textbox', { name: '项目数' })).toHaveValue('10');
  fireEvent.change(screen.getByRole('combobox', { name: '筛选条件' }), {
    target: { value: 'bottom-percent' },
  });
  fireEvent.change(screen.getByRole('textbox', { name: '百分比' }), {
    target: { value: '50' },
  });
  fireEvent.click(screen.getByRole('button', { name: '确定' }));

  expect(requests).toEqual([
    {
      sheetId: 'sheet-1',
      column: 1,
      filterRange: { row: [2, 4], column: [0, 2] },
      criteria: { type: 'bottom-percent', percent: 50 },
    },
  ]);
  unmount();
  canvas.remove();
});

test('defers worksheet range discovery until AutoFilter is toggled', () => {
  let worksheetScans = 0;
  const content = spreadsheetContent(false);
  const sourceData = content.sheets[0]?.data;
  if (!sourceData) throw new Error('Expected spreadsheet fixture data.');
  content.sheets[0] = {
    ...content.sheets[0],
    data: new Proxy(sourceData, {
      ownKeys: (target) => {
        worksheetScans += 1;
        return Reflect.ownKeys(target);
      },
    }),
  };
  const canvas = document.createElement('div');
  const { result, rerender, unmount } = renderHook(
    ({ selection }) =>
      useSpreadsheetAutoFilter({
        canvasRef: { current: canvas },
        commandsRef: noSpreadsheetCommandsRef,
        content,
        editable: true,
        onChange: () => undefined,
        selection,
        sheetId: 'sheet-1',
      }),
    { initialProps: { selection: cellSelection(3, 1) } },
  );

  rerender({ selection: cellSelection(3, 1) });

  expect(worksheetScans).toBe(0);
  act(() => expect(result.current.commandPort.toggle()).toBe(true));
  expect(worksheetScans).toBe(1);
  unmount();
});

test('adapts the vendor filter trigger and menu for keyboard use', async () => {
  const canvas = document.createElement('div');
  const triggers = Array.from({ length: 3 }, () => {
    const trigger = document.createElement('div');
    trigger.className = 'luckysheet-filter-options';
    trigger.tabIndex = 0;
    canvas.append(trigger);
    return trigger;
  });
  document.body.append(canvas);
  let cancelCount = 0;
  triggers[2]?.addEventListener('click', () => {
    if (canvas.querySelector('.fortune-filter-menu')) return;
    const menu = document.createElement('div');
    menu.className = 'fortune-context-menu fortune-filter-menu';
    const sort = document.createElement('div');
    sort.className = 'luckysheet-cols-menuitem';
    sort.tabIndex = 0;
    sort.textContent = '升序';
    const search = document.createElement('input');
    search.className = 'luckysheet-mousedown-cancel';
    const value = document.createElement('div');
    value.className = 'select-item';
    const checkbox = document.createElement('input');
    checkbox.type = 'checkbox';
    const valueLabel = document.createElement('div');
    valueLabel.textContent = '有风险';
    value.append(checkbox, valueLabel);
    const cancel = document.createElement('div');
    cancel.className = 'button-basic button-default';
    cancel.tabIndex = 0;
    cancel.textContent = '取消';
    cancel.addEventListener('click', () => {
      cancelCount += 1;
      menu.remove();
    });
    menu.append(sort, search, value, cancel);
    canvas.append(menu);
  });

  const content = spreadsheetContent(true);
  const { result, unmount } = renderHook(() =>
    useSpreadsheetAutoFilter({
      canvasRef: { current: canvas },
      commandsRef: noSpreadsheetCommandsRef,
      content,
      editable: true,
      onChange: () => undefined,
      selection: cellSelection(2, 0),
      sheetId: 'sheet-1',
      workbook: {
        getSelection: () => [cellSelection(2, 2)],
      },
    }),
  );

  await waitFor(() => {
    expect(triggers[2]).toHaveAttribute('role', 'button');
    expect(triggers[2]).toHaveAttribute('aria-label', '状态 筛选');
    expect(triggers[2]).toHaveAttribute('aria-haspopup', 'dialog');
  });
  expect(result.current.reserveAltKey()).toBe(true);
  act(() => expect(result.current.commandPort.openMenu()).toBe(true));

  const menu = await waitFor(() => {
    const candidate = canvas.querySelector<HTMLElement>('.fortune-filter-menu');
    expect(candidate).toHaveAttribute('role', 'dialog');
    expect(candidate).toHaveAttribute('aria-label', '状态 筛选');
    return candidate as HTMLElement;
  });
  expect(menu.querySelector('input:not([type="checkbox"])')).toHaveAttribute(
    'aria-label',
    '搜索筛选值',
  );
  expect(menu.querySelector('input[type="checkbox"]')).toHaveAttribute(
    'aria-label',
    '显示 有风险',
  );
  await waitFor(() =>
    expect(document.activeElement).toBe(
      menu.querySelector('.luckysheet-cols-menuitem'),
    ),
  );

  menu.dispatchEvent(
    new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      key: 'Escape',
    }),
  );
  expect(cancelCount).toBe(1);
  await waitFor(() => expect(document.activeElement).toBe(triggers[2]));

  unmount();
  canvas.remove();
});

test('opens the owned condition dialog from the vendor menu and filters controlled rows', async () => {
  const canvas = document.createElement('div');
  const triggers = Array.from({ length: 3 }, () => {
    const trigger = document.createElement('div');
    trigger.className = 'luckysheet-filter-options';
    trigger.tabIndex = 0;
    canvas.append(trigger);
    return trigger;
  });
  const trigger = triggers[2] as HTMLElement;
  document.body.append(canvas);
  trigger.addEventListener('click', () => {
    if (canvas.querySelector('.fortune-filter-menu')) return;
    const menu = document.createElement('div');
    menu.className = 'fortune-context-menu fortune-filter-menu';
    const condition = document.createElement('div');
    condition.className = 'luckysheet-cols-menuitem';
    condition.tabIndex = 0;
    condition.textContent = '按条件过滤';
    const cancel = document.createElement('div');
    cancel.className = 'button-basic button-default';
    cancel.tabIndex = 0;
    cancel.textContent = '取消';
    cancel.addEventListener('click', () => menu.remove());
    menu.append(condition, cancel);
    canvas.append(menu);
  });
  const changes: WorkSpreadsheetContent[] = [];
  const requests: Array<
    Parameters<SpreadsheetEditorCommands['applyAutoFilterCriteria']>[0]
  > = [];
  let selectionDuringApply: ReturnType<
    SpreadsheetAutoFilterController['selectionForChange']
  > = null;
  let readSelectionForChange: SpreadsheetAutoFilterController['selectionForChange'] =
    () => null;
  const content = spreadsheetContent(true);
  content.sheets[0] = {
    ...content.sheets[0],
    config: { rowhidden: { '9': 0 } },
  };
  const commandsRef = {
    current: {
      applyAutoFilterCriteria: (request) => {
        requests.push(request);
        selectionDuringApply = readSelectionForChange();
        const next = applySpreadsheetAutoFilterCriteria(
          content,
          request.sheetId,
          request.column,
          request.criteria,
        );
        if (!next) return false;
        changes.push(next);
        return true;
      },
    } as SpreadsheetEditorCommands,
  };
  const { result, unmount } = renderHook(() =>
    useSpreadsheetAutoFilter({
      canvasRef: { current: canvas },
      commandsRef,
      content,
      editable: true,
      onChange: () => {
        throw new Error('Condition dialog bypassed the command boundary.');
      },
      selection: cellSelection(2, 2),
      sheetId: 'sheet-1',
    }),
  );
  readSelectionForChange = () => result.current.selectionForChange();

  fireEvent.pointerDown(trigger);
  fireEvent.click(trigger);
  const condition = await waitFor(() => {
    const candidate = canvas.querySelector<HTMLElement>(
      '[data-a3s-auto-filter-condition]',
    );
    expect(candidate).not.toBeNull();
    return candidate as HTMLElement;
  });
  fireEvent.click(condition);
  await waitFor(() => expect(result.current.dialog).not.toBeNull());
  render(result.current.dialog);

  expect(
    screen.getByRole('dialog', { name: '自定义自动筛选' }),
  ).toHaveTextContent('季度经营!状态');
  fireEvent.change(screen.getByRole('combobox', { name: '筛选条件' }), {
    target: { value: 'matches-wildcard' },
  });
  fireEvent.change(screen.getByRole('textbox', { name: '通配符表达式' }), {
    target: { value: '有?险' },
  });
  fireEvent.click(screen.getByRole('button', { name: '确定' }));

  expect(requests).toEqual([
    {
      sheetId: 'sheet-1',
      column: 2,
      filterRange: { row: [2, 4], column: [0, 2] },
      criteria: { type: 'matches-wildcard', value: '有?险' },
    },
  ]);
  expect(selectionDuringApply).toEqual({
    sheetId: 'sheet-1',
    selections: [cellSelection(2, 2)],
  });
  expect(result.current.selectionForChange()).toBeNull();
  expect(changes).toHaveLength(1);
  expect(changes[0]?.sheets[0]?.config?.rowhidden).toEqual({
    '3': 0,
    '9': 0,
  });
  expect(changes[0]?.sheets[0]?.filter?.['2']).toMatchObject({
    cindex: 2,
    rowhidden: { '3': 0 },
  });
  unmount();
  canvas.remove();
});

function cellSelection(row: number, column: number): Selection {
  return {
    row: [row, row],
    column: [column, column],
    row_focus: row,
    column_focus: column,
  };
}

function spreadsheetContent(active: boolean): WorkSpreadsheetContent {
  return {
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-1',
        name: '季度经营',
        row: 20,
        column: 8,
        data: [
          [],
          [],
          [{ v: '季度' }, { v: '收入' }, { v: '状态' }],
          [{ v: 'Q1' }, { v: 120 }, { v: '正常' }],
          [{ v: 'Q2' }, { v: 128 }, { v: '有风险' }],
        ],
        ...(active
          ? {
              filter: {},
              filter_select: { row: [2, 4], column: [0, 2] },
            }
          : {}),
      },
    ],
  };
}
