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
  const group = within(dialog).getByRole('radiogroup', { name: '框线位置' });
  const top = within(group).getByRole('radio', { name: '上框线' });
  const diagonalUp = within(group).getByRole('radio', {
    name: '斜上框线',
  });
  const all = within(group).getByRole('radio', { name: '所有框线' });
  const none = within(group).getByRole('radio', { name: '无框线' });
  const outside = within(group).getByRole('radio', {
    name: '外侧框线',
  });
  await waitFor(() => expect(all).toHaveFocus());
  expect(all).toHaveAttribute('tabindex', '0');
  expect(top).toHaveAttribute('tabindex', '-1');
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

  fireEvent.keyDown(group, { key: 'End' });
  expect(diagonalUp).toHaveFocus();
  fireEvent.keyDown(group, { key: 'ArrowUp' });
  expect(
    within(group).getByRole('radio', { name: '内部竖框线' }),
  ).toHaveFocus();
  fireEvent.keyDown(group, { key: 'ArrowLeft' });
  expect(
    within(group).getByRole('radio', { name: '内部横框线' }),
  ).toHaveFocus();
  fireEvent.keyDown(group, { key: 'ArrowRight' });
  expect(
    within(group).getByRole('radio', { name: '内部竖框线' }),
  ).toHaveFocus();
  fireEvent.keyDown(group, { key: 'Home' });
  expect(top).toHaveFocus();
  fireEvent.keyDown(group, { key: 'ArrowRight' });
  expect(within(group).getByRole('radio', { name: '下框线' })).toHaveFocus();
  fireEvent.keyDown(group, { key: 'ArrowDown' });
  expect(within(group).getByRole('radio', { name: '右框线' })).toHaveFocus();
  fireEvent.keyDown(group, { key: 'Home' });
  expect(top).toHaveFocus();

  fireEvent.click(within(dialog).getByRole('combobox', { name: '框线样式' }));
  fireEvent.click(screen.getByRole('option', { name: '粗实线' }));
  fireEvent.click(within(dialog).getByRole('button', { name: '框线颜色' }));
  fireEvent.click(screen.getByRole('option', { name: '颜色 #c00000' }));
  fireEvent.click(outside);

  expect(formats).toEqual([
    { target: 'outside', color: '#c00000', style: 'thick' },
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
    '外侧框线（粗实线，#C00000；Cmd/Ctrl+Shift+&）',
  );
  fireEvent.click(primary);
  expect(formats).toEqual([
    { target: 'outside', color: '#c00000', style: 'thick' },
    { target: 'outside', color: '#c00000', style: 'thick' },
  ]);
});

test('restores dirty border style and color drafts on Escape without closing', async () => {
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
  const style = within(dialog).getByRole('combobox', { name: '框线样式' });
  await waitFor(() =>
    expect(
      within(dialog).getByRole('radio', { name: '所有框线' }),
    ).toHaveFocus(),
  );

  fireEvent.click(style);
  fireEvent.click(screen.getByRole('option', { name: '粗实线' }));
  expect(style).toHaveTextContent('粗实线');

  const escapeRoot = dialog.querySelector(
    '[data-office-escape-consumer="true"]',
  );
  expect(escapeRoot).toBeTruthy();
  fireEvent.keyDown(escapeRoot as Element, { key: 'Escape' });

  expect(screen.getByRole('dialog', { name: '框线设置' })).toBeTruthy();
  expect(style).toHaveTextContent('细实线');
  expect(
    dialog.querySelector('[data-office-escape-consumer="true"]'),
  ).toBeNull();

  fireEvent.keyDown(dialog, { key: 'Escape' });
  expect(screen.queryByRole('dialog', { name: '框线设置' })).toBeNull();
  expect(disclosure).toHaveFocus();
  expect(formats).toEqual([]);
});

test('keeps radiogroup tab stop on the checked border target after reopen', async () => {
  render(
    <SpreadsheetBorderRibbon
      can={borderCan()}
      commands={borderCommands(() => true)}
    />,
  );

  fireEvent.click(screen.getByRole('button', { name: '更多框线' }));
  fireEvent.click(screen.getByRole('radio', { name: '外侧框线' }));
  expect(screen.queryByRole('dialog', { name: '框线设置' })).toBeNull();

  fireEvent.click(screen.getByRole('button', { name: '更多框线' }));
  const dialog = screen.getByRole('dialog', { name: '框线设置' });
  const group = within(dialog).getByRole('radiogroup', { name: '框线位置' });
  const outside = within(group).getByRole('radio', { name: '外侧框线' });
  await waitFor(() => expect(outside).toHaveFocus());
  expect(outside).toHaveAttribute('tabindex', '0');
  expect(outside).toHaveAttribute('aria-checked', 'true');
  expect(within(group).getByRole('radio', { name: '上框线' })).toHaveAttribute(
    'tabindex',
    '-1',
  );
});

test('disables only border targets rejected by the command capability', () => {
  const can = borderCan();
  can.setSelectedCellBorders = (format) =>
    format.target !== 'diagonalDown' && format.target !== 'diagonalUp';
  render(
    <SpreadsheetBorderRibbon can={can} commands={borderCommands(() => true)} />,
  );

  fireEvent.click(screen.getByRole('button', { name: '更多框线' }));
  const group = screen.getByRole('radiogroup', { name: '框线位置' });
  expect(within(group).getByRole('radio', { name: '所有框线' })).toBeEnabled();
  expect(within(group).getByRole('radio', { name: '斜下框线' })).toBeDisabled();
  expect(within(group).getByRole('radio', { name: '斜上框线' })).toBeDisabled();
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
