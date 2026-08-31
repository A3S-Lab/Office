import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, within } from '@testing-library/react';
import {
  createSpreadsheetSortCustomList,
  MAX_SPREADSHEET_SORT_USER_CUSTOM_LISTS,
  mergeSpreadsheetSortCustomLists,
  type SpreadsheetSortCustomList,
} from '../src/internal/features/work/editors/spreadsheet-sort-custom-list';
import {
  type SpreadsheetSortCustomListManagementResult,
  SpreadsheetSortCustomListManagerDialog,
} from '../src/internal/features/work/editors/spreadsheet-sort-custom-list-manager';

test('keeps built-in lists read-only and rejects a duplicate user list', () => {
  render(
    <SpreadsheetSortCustomListManagerDialog
      customLists={customLists([
        ['High', 'Medium', 'Low'],
        ['North', 'South'],
      ])}
      restoreFocusTarget={() => null}
      onApply={() => undefined}
      onClose={() => undefined}
    />,
  );

  const list = screen.getByRole('listbox', { name: '自定义序列列表' });
  expect(list).toHaveValue('built-in:0');
  expect(
    screen.getByRole('textbox', { name: '自定义序列项目' }),
  ).toHaveAttribute('readonly');
  expect(screen.getByRole('button', { name: '删除序列' })).toBeDisabled();
  expect(screen.getByRole('button', { name: '上移序列' })).toBeDisabled();
  expect(screen.getByRole('button', { name: '下移序列' })).toBeDisabled();

  fireEvent.click(screen.getByRole('button', { name: '新建序列' }));
  fireEvent.change(screen.getByRole('textbox', { name: '自定义序列项目' }), {
    target: { value: ' high \nMEDIUM\nlow' },
  });
  fireEvent.click(screen.getByRole('button', { name: '添加序列' }));

  expect(screen.getByRole('alert')).toHaveTextContent('该自定义序列已存在');
  expect(
    screen.getByRole('textbox', { name: '自定义序列项目' }),
  ).toHaveAttribute('aria-invalid', 'true');
});

test('adds, edits, reorders, and deletes user lists as one atomic update', () => {
  const applied: SpreadsheetSortCustomListManagementResult[] = [];
  render(
    <SpreadsheetSortCustomListManagerDialog
      customLists={customLists([
        ['High', 'Medium', 'Low'],
        ['North', 'South'],
        ['Red', 'Amber', 'Green'],
      ])}
      restoreFocusTarget={() => null}
      onApply={(value) => applied.push(value)}
      onClose={() => undefined}
    />,
  );

  const list = screen.getByRole('listbox', { name: '自定义序列列表' });
  fireEvent.change(list, { target: { value: 'user:0' } });
  fireEvent.change(screen.getByRole('textbox', { name: '自定义序列项目' }), {
    target: { value: 'Critical\nNormal' },
  });
  fireEvent.click(screen.getByRole('button', { name: '保存更改' }));
  fireEvent.click(screen.getByRole('button', { name: '下移序列' }));

  expect(
    within(list)
      .getAllByRole('option')
      .slice(-3)
      .map((option) => option.textContent),
  ).toEqual(['North → South', 'Critical → Normal', 'Red → Amber → Green']);

  fireEvent.change(list, { target: { value: 'user:2' } });
  fireEvent.click(screen.getByRole('button', { name: '删除序列' }));
  fireEvent.click(screen.getByRole('button', { name: '新建序列' }));
  fireEvent.change(screen.getByRole('textbox', { name: '自定义序列项目' }), {
    target: { value: 'Queued\nRunning\nDone' },
  });
  fireEvent.click(screen.getByRole('button', { name: '添加序列' }));
  fireEvent.click(screen.getByRole('button', { name: '确定' }));

  expect(applied).toEqual([
    {
      lists: [
        ['North', 'South'],
        ['Critical', 'Normal'],
        ['Queued', 'Running', 'Done'],
      ],
      changes: [
        {
          previous: ['High', 'Medium', 'Low'],
          next: ['Critical', 'Normal'],
        },
        { previous: ['Red', 'Amber', 'Green'], next: null },
      ],
    },
  ]);
});

test('disables new-list authoring at the bounded preference limit', () => {
  const entries = Array.from(
    { length: MAX_SPREADSHEET_SORT_USER_CUSTOM_LISTS },
    (_, index) => [`Priority ${index + 1}`, `Fallback ${index + 1}`],
  );

  render(
    <SpreadsheetSortCustomListManagerDialog
      customLists={customLists(entries)}
      restoreFocusTarget={() => null}
      onApply={() => undefined}
      onClose={() => undefined}
    />,
  );

  expect(screen.getByRole('button', { name: '新建序列' })).toBeDisabled();
  expect(
    within(
      screen.getByRole('listbox', { name: '自定义序列列表' }),
    ).getAllByRole('option'),
  ).toHaveLength(MAX_SPREADSHEET_SORT_USER_CUSTOM_LISTS + 7);
});

test('discards staged edits when the manager is cancelled', () => {
  const applied: SpreadsheetSortCustomListManagementResult[] = [];
  let closeCount = 0;
  render(
    <SpreadsheetSortCustomListManagerDialog
      customLists={customLists([['High', 'Medium', 'Low']])}
      restoreFocusTarget={() => null}
      onApply={(value) => applied.push(value)}
      onClose={() => {
        closeCount += 1;
      }}
    />,
  );

  fireEvent.change(screen.getByRole('listbox', { name: '自定义序列列表' }), {
    target: { value: 'user:0' },
  });
  fireEvent.change(screen.getByRole('textbox', { name: '自定义序列项目' }), {
    target: { value: 'Critical\nNormal' },
  });
  fireEvent.click(screen.getByRole('button', { name: '取消' }));

  expect(applied).toEqual([]);
  expect(closeCount).toBe(1);
});

function customLists(
  entries: readonly (readonly string[])[],
): readonly SpreadsheetSortCustomList[] {
  const stored = entries
    .map((list) => createSpreadsheetSortCustomList(list, 'stored'))
    .filter((list): list is SpreadsheetSortCustomList => list !== null);
  return mergeSpreadsheetSortCustomLists(stored);
}
