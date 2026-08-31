import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, within } from '@testing-library/react';
import type {
  SpreadsheetSortCustomList,
  SpreadsheetSortDialogSource,
  SpreadsheetSortDialogValue,
} from '../src/internal/features/work/editors/spreadsheet-sort';
import {
  createSpreadsheetSortCustomList,
  MAX_SPREADSHEET_SORT_USER_CUSTOM_LISTS,
  mergeSpreadsheetSortCustomLists,
  SPREADSHEET_SORT_BUILT_IN_CUSTOM_LISTS,
} from '../src/internal/features/work/editors/spreadsheet-sort-custom-list';
import { SpreadsheetSortDialog } from '../src/internal/features/work/editors/spreadsheet-sort-dialog';

test('authors, reorders, and applies accessible WPS multi-key sort levels', () => {
  const applied: SpreadsheetSortDialogValue[] = [];
  render(
    <SpreadsheetSortDialog
      source={sortSource()}
      restoreFocusTarget={() => null}
      onApply={(value) => {
        applied.push(value);
        return true;
      }}
      onClose={() => undefined}
    />,
  );

  const dialog = screen.getByRole('dialog', { name: '自定义排序' });
  expect(dialog).toHaveTextContent('Sales!A1:C5');
  expect(screen.getByRole('checkbox', { name: '数据包含标题' })).toBeChecked();
  expect(screen.getByRole('combobox', { name: '排序条件 1 列' })).toHaveValue(
    '0',
  );
  expect(screen.getByRole('combobox', { name: '排序条件 1 次序' })).toHaveValue(
    'ascending',
  );

  fireEvent.click(screen.getByRole('button', { name: '添加条件' }));
  expect(screen.getByRole('combobox', { name: '排序条件 2 列' })).toHaveValue(
    '1',
  );
  fireEvent.change(screen.getByRole('combobox', { name: '排序条件 2 次序' }), {
    target: { value: 'descending' },
  });
  fireEvent.click(screen.getByRole('button', { name: '上移条件 2' }));

  const levels = dialog.querySelectorAll('.work-spreadsheet-sort-level');
  expect(levels).toHaveLength(2);
  expect(
    within(levels[0] as HTMLElement).getByRole('combobox', {
      name: '排序条件 1 列',
    }),
  ).toHaveValue('1');
  expect(
    within(levels[0] as HTMLElement).getByRole('combobox', {
      name: '排序条件 1 次序',
    }),
  ).toHaveValue('descending');
  expect(
    within(levels[1] as HTMLElement).getByRole('combobox', {
      name: '排序条件 2 列',
    }),
  ).toHaveValue('0');

  fireEvent.click(screen.getByRole('button', { name: '确定' }));
  expect(applied).toEqual([
    {
      orientation: 'top-to-bottom',
      caseSensitive: false,
      textMethod: 'pinyin',
      hasHeader: true,
      keys: [
        { index: 1, direction: 'descending' },
        { index: 0, direction: 'ascending' },
      ],
    },
  ]);
});

test('removes levels and exposes the compact header contract', () => {
  render(
    <SpreadsheetSortDialog
      source={{
        ...sortSource(),
        value: {
          orientation: 'top-to-bottom',
          caseSensitive: false,
          textMethod: 'pinyin',
          hasHeader: false,
          keys: [
            { index: 0, direction: 'ascending' },
            { index: 1, direction: 'descending' },
            { index: 2, direction: 'ascending' },
          ],
        },
      }}
      restoreFocusTarget={() => null}
      onApply={() => true}
      onClose={() => undefined}
    />,
  );

  fireEvent.click(screen.getByRole('button', { name: '删除条件 2' }));
  expect(screen.queryByRole('combobox', { name: '排序条件 3 列' })).toBeNull();
  expect(screen.getByRole('button', { name: '添加条件' })).toBeEnabled();
  expect(
    screen.getByRole('checkbox', { name: '数据包含标题' }),
  ).not.toBeChecked();
});

test('adds distinct appearance priorities on a one-column range', () => {
  const applied: SpreadsheetSortDialogValue[] = [];
  const source = sortSource();
  render(
    <SpreadsheetSortDialog
      source={{
        ...source,
        range: { row: [0, 4], column: [0, 0] },
        rangeReference: 'A1:A5',
        columns: source.columns.slice(0, 1),
        appearanceRows: source.appearanceRows.map((row) => row.slice(0, 1)),
        value: {
          orientation: 'top-to-bottom',
          caseSensitive: false,
          textMethod: 'pinyin',
          hasHeader: true,
          keys: [
            {
              index: 0,
              sortOn: 'cell-color',
              color: '#eef4ff',
              position: 'first',
            },
          ],
        },
      }}
      restoreFocusTarget={() => null}
      onApply={(value) => {
        applied.push(value);
        return true;
      }}
      onClose={() => undefined}
    />,
  );

  const add = screen.getByRole('button', { name: '添加条件' });
  expect(add).toBeEnabled();
  fireEvent.click(add);
  expect(screen.getByRole('combobox', { name: '排序条件 2 列' })).toHaveValue(
    '0',
  );
  expect(
    screen.getByRole('combobox', { name: '排序条件 2 排序依据' }),
  ).toHaveValue('cell-color');
  expect(
    screen.getByRole('combobox', { name: '排序条件 2 目标外观' }),
  ).toHaveValue('cell-color:none');

  fireEvent.click(screen.getByRole('button', { name: '确定' }));
  expect(applied).toEqual([
    {
      orientation: 'top-to-bottom',
      caseSensitive: false,
      textMethod: 'pinyin',
      hasHeader: true,
      keys: [
        {
          index: 0,
          sortOn: 'cell-color',
          color: '#eef4ff',
          position: 'first',
        },
        {
          index: 0,
          sortOn: 'cell-color',
          color: null,
          position: 'first',
        },
      ],
    },
  ]);
});

test('creates and applies a reusable custom-list order without leaving the dialog', () => {
  const applied: SpreadsheetSortDialogValue[] = [];
  const remembered: SpreadsheetSortCustomList[] = [];
  render(
    <SpreadsheetSortDialog
      source={sortSource()}
      restoreFocusTarget={() => null}
      onApply={(value) => {
        applied.push(value);
        return true;
      }}
      onRememberCustomList={(list) => {
        remembered.push(list);
        return undefined;
      }}
      onClose={() => undefined}
    />,
  );

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

  expect(
    within(order).getByRole('option', {
      name: '有风险 → 进行中 → 正常 → …',
    }),
  ).toBeInTheDocument();
  expect(remembered).toEqual([
    {
      source: 'session',
      label: '有风险 → 进行中 → 正常 → …',
      entries: ['有风险', '进行中', '正常', '已完成'],
    },
  ]);

  fireEvent.click(screen.getByRole('button', { name: '确定' }));
  expect(applied).toEqual([
    {
      orientation: 'top-to-bottom',
      caseSensitive: false,
      textMethod: 'pinyin',
      hasHeader: true,
      keys: [
        {
          index: 0,
          customList: ['有风险', '进行中', '正常', '已完成'],
        },
      ],
    },
  ]);
});

test('keeps a stored identity when the initial key carries the same list', () => {
  const entries = ['有风险', '进行中', '正常', '已完成'];
  const stored = createSpreadsheetSortCustomList(entries, 'stored');
  expect(stored).not.toBeNull();
  const source = sortSource();

  render(
    <SpreadsheetSortDialog
      source={{
        ...source,
        customLists: mergeSpreadsheetSortCustomLists(stored ? [stored] : []),
        value: {
          ...source.value,
          keys: [{ index: 0, customList: entries }],
        },
      }}
      restoreFocusTarget={() => null}
      onApply={() => true}
      onClose={() => undefined}
    />,
  );

  const order = screen.getByRole('combobox', { name: '排序条件 1 次序' });
  expect(order).toHaveValue('custom-list:7');
  expect(order.querySelector('optgroup[label="已保存的序列"]')).not.toBeNull();
  expect(order.querySelector('optgroup[label="本次会话的序列"]')).toBeNull();
});

test('reconciles an active custom-list key after preference edits and deletion', () => {
  const original = ['High', 'Medium', 'Low'];
  const stored = createSpreadsheetSortCustomList(original, 'stored');
  expect(stored).not.toBeNull();
  const source = sortSource();
  const updates: Array<readonly (readonly string[])[]> = [];

  render(
    <SpreadsheetSortDialog
      source={{
        ...source,
        customLists: mergeSpreadsheetSortCustomLists(stored ? [stored] : []),
        value: {
          ...source.value,
          keys: [{ index: 0, customList: original }],
        },
      }}
      restoreFocusTarget={() => null}
      onApply={() => true}
      onUpdateCustomLists={(lists) => {
        updates.push(lists);
        return lists
          .map((entries) => createSpreadsheetSortCustomList(entries, 'stored'))
          .filter((list): list is SpreadsheetSortCustomList => list !== null);
      }}
      onClose={() => undefined}
    />,
  );

  const managerButton = screen.getByRole('button', {
    name: '管理自定义序列',
  });
  fireEvent.click(managerButton);
  let manager = screen.getByRole('dialog', { name: '自定义序列' });
  fireEvent.change(
    within(manager).getByRole('listbox', { name: '自定义序列列表' }),
    { target: { value: 'user:0' } },
  );
  fireEvent.change(
    within(manager).getByRole('textbox', { name: '自定义序列项目' }),
    { target: { value: 'Critical\nNormal' } },
  );
  fireEvent.click(within(manager).getByRole('button', { name: '保存更改' }));
  fireEvent.click(within(manager).getByRole('button', { name: '确定' }));

  const order = screen.getByRole('combobox', { name: '排序条件 1 次序' });
  expect(
    within(order).getByRole('option', { name: 'Critical → Normal' }),
  ).toBeInTheDocument();
  expect(order).toHaveValue('custom-list:7');
  expect(updates).toEqual([[['Critical', 'Normal']]]);
  expect(managerButton).toHaveFocus();

  fireEvent.click(managerButton);
  manager = screen.getByRole('dialog', { name: '自定义序列' });
  fireEvent.change(
    within(manager).getByRole('listbox', { name: '自定义序列列表' }),
    { target: { value: 'user:0' } },
  );
  fireEvent.click(within(manager).getByRole('button', { name: '删除序列' }));
  fireEvent.click(within(manager).getByRole('button', { name: '确定' }));

  expect(order).toHaveValue('ascending');
  expect(updates).toEqual([[['Critical', 'Normal']], []]);
});

test('rejects another authored list after the mounted-editor user-list bound', () => {
  const storedLists = Array.from(
    { length: MAX_SPREADSHEET_SORT_USER_CUSTOM_LISTS },
    (_, index) =>
      createSpreadsheetSortCustomList(
        [`First ${index}`, `Second ${index}`],
        'stored',
      ),
  ).filter((list): list is SpreadsheetSortCustomList => list !== null);
  const remembered: SpreadsheetSortCustomList[] = [];
  render(
    <SpreadsheetSortDialog
      source={{
        ...sortSource(),
        customLists: mergeSpreadsheetSortCustomLists(storedLists),
      }}
      restoreFocusTarget={() => null}
      onApply={() => true}
      onRememberCustomList={(list) => {
        remembered.push(list);
        return undefined;
      }}
      onClose={() => undefined}
    />,
  );

  const order = screen.getByRole('combobox', { name: '排序条件 1 次序' });
  const createOption = within(order).getByRole('option', {
    name: '新建自定义序列…',
  }) as HTMLOptionElement;
  fireEvent.change(order, { target: { value: createOption.value } });
  fireEvent.change(
    screen.getByRole('textbox', { name: '排序条件 1 自定义序列' }),
    { target: { value: 'Overflow first\nOverflow second' } },
  );
  fireEvent.click(screen.getByRole('button', { name: '使用序列' }));

  expect(screen.getByRole('alert')).toHaveTextContent(
    `当前编辑器最多保留 ${MAX_SPREADSHEET_SORT_USER_CUSTOM_LISTS} 个自定义序列。`,
  );
  expect(
    screen.getByRole('textbox', { name: '排序条件 1 自定义序列' }),
  ).toHaveAttribute('aria-invalid', 'true');
  expect(remembered).toEqual([]);
});

test('authors cell-color, font-color, and conditional-icon sort levels', () => {
  const applied: SpreadsheetSortDialogValue[] = [];
  render(
    <SpreadsheetSortDialog
      source={sortSource()}
      restoreFocusTarget={() => null}
      onApply={(value) => {
        applied.push(value);
        return true;
      }}
      onClose={() => undefined}
    />,
  );

  fireEvent.change(
    screen.getByRole('combobox', { name: '排序条件 1 排序依据' }),
    { target: { value: 'cell-color' } },
  );
  expect(
    screen.getByRole('combobox', { name: '排序条件 1 目标外观' }),
  ).toHaveAccessibleName('排序条件 1 目标外观');
  fireEvent.change(
    screen.getByRole('combobox', { name: '排序条件 1 目标外观' }),
    { target: { value: 'cell-color:#eef4ff' } },
  );
  fireEvent.change(screen.getByRole('combobox', { name: '排序条件 1 位置' }), {
    target: { value: 'last' },
  });

  fireEvent.click(screen.getByRole('button', { name: '添加条件' }));
  fireEvent.change(
    screen.getByRole('combobox', { name: '排序条件 2 排序依据' }),
    { target: { value: 'font-color' } },
  );
  fireEvent.change(
    screen.getByRole('combobox', { name: '排序条件 2 目标外观' }),
    { target: { value: 'font-color:#d84b4f' } },
  );

  fireEvent.click(screen.getByRole('button', { name: '添加条件' }));
  fireEvent.change(
    screen.getByRole('combobox', { name: '排序条件 3 排序依据' }),
    { target: { value: 'icon' } },
  );
  expect(
    within(
      screen.getByRole('combobox', { name: '排序条件 3 目标外观' }),
    ).getByRole('option', { name: /三色交通灯（实心） 3\/3/ }),
  ).toBeInTheDocument();
  fireEvent.change(
    screen.getByRole('combobox', { name: '排序条件 3 目标外观' }),
    { target: { value: 'icon:3TrafficLights1:2' } },
  );

  fireEvent.click(screen.getByRole('button', { name: '确定' }));
  expect(applied).toEqual([
    {
      orientation: 'top-to-bottom',
      caseSensitive: false,
      textMethod: 'pinyin',
      hasHeader: true,
      keys: [
        {
          index: 0,
          sortOn: 'cell-color',
          color: '#eef4ff',
          position: 'last',
        },
        {
          index: 1,
          sortOn: 'font-color',
          color: '#d84b4f',
          position: 'first',
        },
        {
          index: 2,
          sortOn: 'icon',
          icon: { iconSet: '3TrafficLights1', index: 2 },
          position: 'first',
        },
      ],
    },
  ]);
});

test('reselects an available appearance when the retained-header boundary changes', () => {
  render(
    <SpreadsheetSortDialog
      source={sortSource()}
      restoreFocusTarget={() => null}
      onApply={() => true}
      onClose={() => undefined}
    />,
  );

  fireEvent.click(screen.getByRole('checkbox', { name: '数据包含标题' }));
  fireEvent.change(
    screen.getByRole('combobox', { name: '排序条件 1 排序依据' }),
    { target: { value: 'cell-color' } },
  );
  expect(
    screen.getByRole('combobox', { name: '排序条件 1 目标外观' }),
  ).toHaveValue('cell-color:#4472c4');

  fireEvent.click(screen.getByRole('checkbox', { name: '数据包含标题' }));
  expect(
    screen.getByRole('combobox', { name: '排序条件 1 目标外观' }),
  ).toHaveValue('cell-color:#eef4ff');
});

function sortSource(): SpreadsheetSortDialogSource {
  return {
    sheetId: 'sheet-1',
    sheetName: 'Sales',
    range: { row: [0, 4], column: [0, 2] },
    rangeReference: 'A1:C5',
    activeRow: 1,
    columns: [
      { index: 0, label: 'A（Team）' },
      { index: 1, label: 'B（Score）' },
      { index: 2, label: 'C（Owner）' },
    ],
    rows: [
      { index: 0, label: '行 1' },
      { index: 1, label: '行 2' },
      { index: 2, label: '行 3' },
      { index: 3, label: '行 4' },
      { index: 4, label: '行 5' },
    ],
    customLists: SPREADSHEET_SORT_BUILT_IN_CUSTOM_LISTS,
    appearanceRows: [
      [
        { cellColor: '#4472c4', fontColor: '#ffffff', icon: null },
        { cellColor: '#4472c4', fontColor: '#ffffff', icon: null },
        { cellColor: '#4472c4', fontColor: '#ffffff', icon: null },
      ],
      [
        { cellColor: '#eef4ff', fontColor: null, icon: null },
        { cellColor: null, fontColor: '#d84b4f', icon: null },
        {
          cellColor: null,
          fontColor: null,
          icon: { iconSet: '3TrafficLights1', index: 2 },
        },
      ],
      [
        { cellColor: null, fontColor: null, icon: null },
        { cellColor: null, fontColor: null, icon: null },
        {
          cellColor: null,
          fontColor: null,
          icon: { iconSet: '3TrafficLights1', index: 1 },
        },
      ],
      [
        { cellColor: '#eef4ff', fontColor: null, icon: null },
        { cellColor: null, fontColor: '#d84b4f', icon: null },
        {
          cellColor: null,
          fontColor: null,
          icon: { iconSet: '3TrafficLights1', index: 0 },
        },
      ],
      [
        { cellColor: null, fontColor: null, icon: null },
        { cellColor: null, fontColor: null, icon: null },
        { cellColor: null, fontColor: null, icon: null },
      ],
    ],
    value: {
      orientation: 'top-to-bottom',
      caseSensitive: false,
      textMethod: 'pinyin',
      hasHeader: true,
      keys: [{ index: 0, direction: 'ascending' }],
    },
  };
}
