import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, within } from '@testing-library/react';
import { SpreadsheetDateTimeMenu } from '../src/internal/features/work/editors/spreadsheet-date-time-ribbon';

test('exposes discoverable WPS date and time commands with keyboard metadata', () => {
  const inserts: string[] = [];
  render(
    <SpreadsheetDateTimeMenu
      can={{ insertCurrentDateTime: () => true }}
      commands={{
        insertCurrentDateTime: (kind) => {
          inserts.push(kind);
          return true;
        },
      }}
    />,
  );

  const trigger = screen.getByRole('button', { name: '日期和时间' });
  expect(trigger).toHaveAttribute(
    'title',
    '插入当前日期或时间（Ctrl+; / Ctrl+Shift+;）',
  );
  fireEvent.click(trigger);

  const menu = screen.getByRole('menu', { name: '插入日期和时间' });
  const date = within(menu).getByRole('menuitem', { name: '插入当前日期' });
  const time = within(menu).getByRole('menuitem', { name: '插入当前时间' });
  expect(date).toHaveAttribute('aria-keyshortcuts', 'Control+;');
  expect(time).toHaveAttribute('aria-keyshortcuts', 'Control+Shift+;');
  expect(date).toHaveTextContent('Ctrl+;');
  expect(time).toHaveTextContent('Ctrl+Shift+;');
  fireEvent.click(date);
  expect(screen.queryByRole('menu', { name: '插入日期和时间' })).toBeNull();

  fireEvent.click(trigger);
  fireEvent.click(
    within(screen.getByRole('menu', { name: '插入日期和时间' })).getByRole(
      'menuitem',
      { name: '插入当前时间' },
    ),
  );
  expect(inserts).toEqual(['date', 'time']);
});

test('disables the date and time menu when the active cell is not editable', () => {
  render(
    <SpreadsheetDateTimeMenu
      can={{ insertCurrentDateTime: () => false }}
      commands={{ insertCurrentDateTime: () => false }}
    />,
  );

  expect(screen.getByRole('button', { name: '日期和时间' })).toBeDisabled();
});
