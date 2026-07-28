import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import { SpreadsheetSheetBar } from '../src/internal/features/work/editors/spreadsheet-sheet-bar';

test('uses one accessible worksheet bar for creation, activation, and menus', () => {
  const calls: string[] = [];
  render(
    <SpreadsheetSheetBar
      activeSheetId="sheet-1"
      editable
      sheets={[
        { id: 'sheet-1', name: '执行看板', status: 1 },
        { id: 'sheet-2', name: '风险台账', status: 0 },
      ]}
      onActivate={(id) => calls.push(`activate:${id}`)}
      onCreate={() => calls.push('create')}
      onDelete={(id) => calls.push(`delete:${id}`)}
      onDuplicate={(id) => calls.push(`duplicate:${id}`)}
      onHide={(id) => calls.push(`hide:${id}`)}
      onMove={(id, direction) => calls.push(`move:${id}:${direction}`)}
      onRename={(id, name) => calls.push(`rename:${id}:${name}`)}
      onSetColor={(id, color) => calls.push(`color:${id}:${color ?? 'none'}`)}
      onShow={(id) => calls.push(`show:${id}`)}
    />,
  );

  expect(
    screen.getByRole('navigation', { name: '工作表' }),
  ).toBeInTheDocument();
  expect(screen.getByRole('tab', { name: '执行看板' })).toHaveAttribute(
    'aria-selected',
    'true',
  );

  fireEvent.click(screen.getByRole('button', { name: '新建工作表' }));
  fireEvent.click(screen.getByRole('tab', { name: '风险台账' }));
  expect(calls).toEqual(['create', 'activate:sheet-2']);

  const options = screen.getByRole('button', { name: '执行看板选项' });
  fireEvent.click(options);
  expect(
    screen.getByRole('menu', { name: '执行看板工作表操作' }),
  ).toBeInTheDocument();

  fireEvent.keyDown(document, { key: 'Escape' });
  expect(
    screen.queryByRole('menu', { name: '执行看板工作表操作' }),
  ).not.toBeInTheDocument();
  expect(options).toHaveFocus();
});

test('supports inline worksheet rename and compact color controls', () => {
  const calls: string[] = [];
  render(
    <SpreadsheetSheetBar
      activeSheetId="sheet-1"
      editable
      sheets={[{ id: 'sheet-1', name: '执行看板', status: 1 }]}
      onActivate={() => undefined}
      onCreate={() => undefined}
      onDelete={() => undefined}
      onDuplicate={() => undefined}
      onHide={() => undefined}
      onMove={() => undefined}
      onRename={(id, name) => calls.push(`rename:${id}:${name}`)}
      onSetColor={(id, color) => calls.push(`color:${id}:${color ?? 'none'}`)}
      onShow={() => undefined}
    />,
  );

  fireEvent.doubleClick(screen.getByRole('tab', { name: '执行看板' }));
  const input = screen.getByRole('textbox', { name: '重命名执行看板' });
  fireEvent.change(input, { target: { value: '季度看板' } });
  fireEvent.keyDown(input, { key: 'Enter' });
  expect(calls).toEqual(['rename:sheet-1:季度看板']);

  fireEvent.click(screen.getByRole('button', { name: '执行看板选项' }));
  fireEvent.click(screen.getByRole('button', { name: '红色标签' }));
  expect(calls).toEqual(['rename:sheet-1:季度看板', 'color:sheet-1:#e06c53']);
});
