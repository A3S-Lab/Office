import type { Selection } from '@fortune-sheet/core';
import { expect, test } from '@rstest/core';
import { act, renderHook, waitFor } from '@testing-library/react';
import { useSpreadsheetAutoFilter } from '../src/internal/features/work/editors/use-spreadsheet-auto-filter';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

test('publishes one controlled workbook change when AutoFilter is toggled', () => {
  const changes: WorkSpreadsheetContent[] = [];
  const canvas = document.createElement('div');
  const initial = spreadsheetContent(false);
  const { result, rerender } = renderHook(
    ({ content }) =>
      useSpreadsheetAutoFilter({
        canvasRef: { current: canvas },
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
