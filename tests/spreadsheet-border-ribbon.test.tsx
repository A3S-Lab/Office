import { expect, test } from '@rstest/core';
import {
  fireEvent,
  render,
  screen,
  waitFor,
  within,
} from '@testing-library/react';
import type {
  SpreadsheetEditorCanCommands,
  SpreadsheetEditorCommands,
} from '../src/internal/features/work/editors/spreadsheet-command-controller';
import { SpreadsheetBorderRibbon } from '../src/internal/features/work/editors/spreadsheet-border-ribbon';
import type { SpreadsheetCellBorderFormat } from '../src/internal/features/work/editors/spreadsheet-cell-border';

test('applies the current border from an accessible split control', () => {
  const formats: SpreadsheetCellBorderFormat[] = [];
  render(
    <SpreadsheetBorderRibbon
      can={borderCan()}
      commands={borderCommands((format) => {
        formats.push(format);
        return true;
      })}
    />,
  );

  const primary = screen.getByRole('button', { name: '所有框线' });
  const disclosure = screen.getByRole('button', { name: '更多框线' });
  expect(primary).toHaveAttribute('title', '所有框线（细实线，#000000）');
  expect(disclosure).toHaveAttribute('aria-haspopup', 'dialog');

  fireEvent.click(primary);
  expect(formats).toEqual([{ target: 'all', color: '#000000', style: 'thin' }]);
});

test('supports keyboard border selection with persistent line and color settings', async () => {
  const formats: SpreadsheetCellBorderFormat[] = [];
  render(
    <SpreadsheetBorderRibbon
      can={borderCan()}
      commands={borderCommands((format) => {
        formats.push(format);
        return true;
      })}
    />,
  );

  const disclosure = screen.getByRole('button', { name: '更多框线' });
  fireEvent.click(disclosure);
  const dialog = screen.getByRole('dialog', { name: '框线设置' });
  const menu = within(dialog).getByRole('menu', { name: '框线位置' });
  const top = within(menu).getByRole('menuitemradio', { name: '上框线' });
  const diagonalDown = within(menu).getByRole('menuitemradio', {
    name: '斜下框线',
  });
  const diagonalUp = within(menu).getByRole('menuitemradio', {
    name: '斜上框线',
  });
  await waitFor(() => expect(top).toHaveFocus());
  const all = within(menu).getByRole('menuitemradio', { name: '所有框线' });
  const none = within(menu).getByRole('menuitemradio', { name: '无框线' });
  const outside = within(menu).getByRole('menuitemradio', {
    name: '外侧框线',
  });
  expect(top).toHaveAttribute('tabindex', '0');
  expect(all).toHaveAttribute('tabindex', '-1');
  expect(all).toHaveAttribute('aria-checked', 'true');
  expect(none).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+_ Meta+Shift+_',
  );
  expect(none.querySelector('kbd')).toHaveTextContent('Cmd/Ctrl+Shift+_');
  expect(outside).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+& Meta+Shift+&',
  );
  expect(outside.querySelector('kbd')).toHaveTextContent('Cmd/Ctrl+Shift+&');

  fireEvent.keyDown(menu, { key: 'End' });
  expect(diagonalUp).toHaveFocus();
  fireEvent.keyDown(menu, { key: 'ArrowUp' });
  expect(diagonalDown).toHaveFocus();
  fireEvent.keyDown(menu, { key: 'Home' });
  expect(top).toHaveFocus();

  fireEvent.change(within(dialog).getByRole('combobox', { name: '框线样式' }), {
    target: { value: 'thick' },
  });
  fireEvent.change(within(dialog).getByLabelText('框线颜色'), {
    target: { value: '#b42318' },
  });
  fireEvent.click(outside);

  expect(formats).toEqual([
    { target: 'outside', color: '#b42318', style: 'thick' },
  ]);
  expect(screen.queryByRole('dialog', { name: '框线设置' })).toBeNull();
  expect(disclosure).toHaveFocus();

  const primary = screen.getByRole('button', { name: '外侧框线' });
  expect(primary).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+& Meta+Shift+&',
  );
  expect(primary).toHaveAttribute(
    'title',
    '外侧框线（粗实线，#B42318；Cmd/Ctrl+Shift+&）',
  );
  fireEvent.click(primary);
  expect(formats).toEqual([
    { target: 'outside', color: '#b42318', style: 'thick' },
    { target: 'outside', color: '#b42318', style: 'thick' },
  ]);
});

test('disables only border targets rejected by the command capability', () => {
  const can = borderCan();
  can.setSelectedCellBorders = (format) =>
    format.target !== 'diagonalDown' && format.target !== 'diagonalUp';
  render(
    <SpreadsheetBorderRibbon can={can} commands={borderCommands(() => true)} />,
  );

  fireEvent.click(screen.getByRole('button', { name: '更多框线' }));
  const menu = screen.getByRole('menu', { name: '框线位置' });
  expect(
    within(menu).getByRole('menuitemradio', { name: '所有框线' }),
  ).toBeEnabled();
  expect(
    within(menu).getByRole('menuitemradio', { name: '斜下框线' }),
  ).toBeDisabled();
  expect(
    within(menu).getByRole('menuitemradio', { name: '斜上框线' }),
  ).toBeDisabled();
});

function borderCan(): SpreadsheetEditorCanCommands {
  return {
    setSelectedCellBorders: () => true,
  } as SpreadsheetEditorCanCommands;
}

function borderCommands(
  setSelectedCellBorders: SpreadsheetEditorCommands['setSelectedCellBorders'],
): SpreadsheetEditorCommands {
  return { setSelectedCellBorders } as SpreadsheetEditorCommands;
}
