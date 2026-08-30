import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, within } from '@testing-library/react';
import { SpreadsheetSortDialog } from '../src/internal/features/work/editors/spreadsheet-sort-dialog';
import {
  createSpreadsheetSortCustomList,
  MAX_SPREADSHEET_SORT_SESSION_CUSTOM_LISTS,
  mergeSpreadsheetSortCustomLists,
  SPREADSHEET_SORT_BUILT_IN_CUSTOM_LISTS,
} from '../src/internal/features/work/editors/spreadsheet-sort-custom-list';
import type {
  SpreadsheetSortCustomList,
  SpreadsheetSortDialogSource,
  SpreadsheetSortDialogValue,
} from '../src/internal/features/work/editors/spreadsheet-sort';

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
      hasHeader: true,
      keys: [
        { column: 1, direction: 'descending' },
        { column: 0, direction: 'ascending' },
      ],
    },
  ]);
});

test('removes levels, prevents duplicate keys, and exposes the compact header contract', () => {
  render(
    <SpreadsheetSortDialog
      source={{
        ...sortSource(),
        value: {
          hasHeader: false,
          keys: [
            { column: 0, direction: 'ascending' },
            { column: 1, direction: 'descending' },
            { column: 2, direction: 'ascending' },
          ],
        },
      }}
      restoreFocusTarget={() => null}
      onApply={() => true}
      onClose={() => undefined}
    />,
  );

  expect(screen.getByRole('button', { name: '添加条件' })).toBeDisabled();
  const secondColumn = screen.getByRole('combobox', {
    name: '排序条件 2 列',
  });
  expect(
    within(secondColumn).getByRole('option', { name: 'A（Team）' }),
  ).toBeDisabled();
  fireEvent.click(screen.getByRole('button', { name: '删除条件 2' }));
  expect(screen.queryByRole('combobox', { name: '排序条件 3 列' })).toBeNull();
  expect(screen.getByRole('button', { name: '添加条件' })).toBeEnabled();
  expect(
    screen.getByRole('checkbox', { name: '数据包含标题' }),
  ).not.toBeChecked();
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
      onRememberCustomList={(list) => remembered.push(list)}
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
      hasHeader: true,
      keys: [
        {
          column: 0,
          customList: ['有风险', '进行中', '正常', '已完成'],
        },
      ],
    },
  ]);
});

test('rejects another authored list after the mounted-editor session bound', () => {
  const sessionLists = Array.from(
    { length: MAX_SPREADSHEET_SORT_SESSION_CUSTOM_LISTS },
    (_, index) =>
      createSpreadsheetSortCustomList(
        [`First ${index}`, `Second ${index}`],
        'session',
      ),
  ).filter((list): list is SpreadsheetSortCustomList => list !== null);
  const remembered: SpreadsheetSortCustomList[] = [];
  render(
    <SpreadsheetSortDialog
      source={{
        ...sortSource(),
        customLists: mergeSpreadsheetSortCustomLists(sessionLists),
      }}
      restoreFocusTarget={() => null}
      onApply={() => true}
      onRememberCustomList={(list) => remembered.push(list)}
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
    `当前编辑器会话最多保留 ${MAX_SPREADSHEET_SORT_SESSION_CUSTOM_LISTS} 个自定义序列。`,
  );
  expect(
    screen.getByRole('textbox', { name: '排序条件 1 自定义序列' }),
  ).toHaveAttribute('aria-invalid', 'true');
  expect(remembered).toEqual([]);
});

function sortSource(): SpreadsheetSortDialogSource {
  return {
    sheetId: 'sheet-1',
    sheetName: 'Sales',
    range: { row: [0, 4], column: [0, 2] },
    rangeReference: 'A1:C5',
    columns: [
      { column: 0, label: 'A（Team）' },
      { column: 1, label: 'B（Score）' },
      { column: 2, label: 'C（Owner）' },
    ],
    customLists: SPREADSHEET_SORT_BUILT_IN_CUSTOM_LISTS,
    value: {
      hasHeader: true,
      keys: [{ column: 0, direction: 'ascending' }],
    },
  };
}
