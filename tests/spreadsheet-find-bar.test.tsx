import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import { SpreadsheetFindBar } from '../src/internal/features/work/editors/spreadsheet-find-bar';
import type { SpreadsheetFindMatch } from '../src/internal/features/work/editors/spreadsheet-find';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

const sheet: WorkSpreadsheetContent['sheets'][number] = {
  id: 'sheet-1',
  name: '执行看板',
  data: [[{ v: 'Alpha' }], [{ v: 'Beta Alpha' }]],
};

test('moves through matches while keeping the query ready for keyboard use', () => {
  const selected: SpreadsheetFindMatch[] = [];
  const view = render(
    <SpreadsheetFindBar
      sheet={sheet}
      focusRequest={1}
      onClose={() => undefined}
      onSelectMatch={(match) => selected.push(match)}
    />,
  );
  const query = screen.getByRole('textbox', { name: '查找当前工作表' });
  expect(query).toHaveFocus();

  fireEvent.change(query, { target: { value: 'alpha' } });
  expect(screen.getByText('2 个匹配')).toBeVisible();

  fireEvent.keyDown(query, { key: 'Enter' });
  expect(selected.at(-1)?.reference).toBe('A1');
  expect(screen.getByText('1/2')).toBeVisible();
  expect(query).toHaveFocus();

  fireEvent.click(screen.getByRole('button', { name: '下一个匹配' }));
  expect(selected.at(-1)?.reference).toBe('A2');
  expect(screen.getByText('2/2')).toBeVisible();
  expect(query).toHaveFocus();

  fireEvent.keyDown(query, { key: 'Enter', shiftKey: true });
  expect(selected.at(-1)?.reference).toBe('A1');

  query.blur();
  view.rerender(
    <SpreadsheetFindBar
      sheet={sheet}
      focusRequest={2}
      onClose={() => undefined}
      onSelectMatch={(match) => selected.push(match)}
    />,
  );
  expect(query).toHaveFocus();
  expect(query).toHaveValue('alpha');
});

test('closes from Escape without leaking the key to the spreadsheet', () => {
  let closeCount = 0;
  render(
    <SpreadsheetFindBar
      sheet={sheet}
      focusRequest={1}
      onClose={() => {
        closeCount += 1;
      }}
      onSelectMatch={() => undefined}
    />,
  );
  const query = screen.getByRole('textbox', { name: '查找当前工作表' });
  const event = fireEvent.keyDown(query, { key: 'Escape' });
  expect(event).toBe(false);
  expect(closeCount).toBe(1);
});
