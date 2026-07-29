import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { SpreadsheetSheetBar } from '../src/internal/features/work/editors/spreadsheet-sheet-bar';

test('uses one accessible worksheet bar for creation, activation, and menus', async () => {
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

  fireEvent.click(options);
  const reopenedMenu = screen.getByRole('menu', {
    name: '执行看板工作表操作',
  });
  await waitFor(() =>
    expect(screen.getByRole('menuitem', { name: '重命名' })).toHaveFocus(),
  );
  fireEvent.keyDown(reopenedMenu, { key: 'Tab' });
  await waitFor(() =>
    expect(screen.getByRole('button', { name: '风险台账选项' })).toHaveFocus(),
  );
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
  const red = screen.getByRole('menuitemradio', { name: '红色标签' });
  expect(red).toHaveAttribute('aria-checked', 'false');
  fireEvent.click(red);
  expect(calls).toEqual(['rename:sheet-1:季度看板', 'color:sheet-1:#e06c53']);
});

test('cancels worksheet rename with Escape and restores tab keyboard focus', async () => {
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
      onSetColor={() => undefined}
      onShow={() => undefined}
    />,
  );

  fireEvent.doubleClick(screen.getByRole('tab', { name: '执行看板' }));
  const input = screen.getByRole('textbox', { name: '重命名执行看板' });
  fireEvent.change(input, { target: { value: '不应保存' } });
  fireEvent.keyDown(input, { key: 'Escape' });

  await waitFor(() =>
    expect(screen.getByRole('tab', { name: '执行看板' })).toHaveFocus(),
  );
  expect(calls).toEqual([]);
  expect(screen.queryByRole('textbox', { name: '重命名执行看板' })).toBeNull();
});

test('navigates worksheet tabs with arrow, Home, and End keys', async () => {
  const calls: string[] = [];
  render(
    <SpreadsheetSheetBar
      activeSheetId="sheet-1"
      editable
      sheets={[
        { id: 'sheet-1', name: '执行看板', status: 1 },
        { id: 'sheet-2', name: '风险台账', status: 0 },
        { id: 'sheet-3', name: '资源计划', status: 0 },
      ]}
      onActivate={(id) => calls.push(id)}
      onCreate={() => undefined}
      onDelete={() => undefined}
      onDuplicate={() => undefined}
      onHide={() => undefined}
      onMove={() => undefined}
      onRename={() => undefined}
      onSetColor={() => undefined}
      onShow={() => undefined}
    />,
  );

  const first = screen.getByRole('tab', { name: '执行看板' });
  fireEvent.keyDown(first, { key: 'ArrowRight' });
  await waitFor(() => expect(calls).toEqual(['sheet-2']));
  await waitFor(() =>
    expect(screen.getByRole('tab', { name: '风险台账' })).toHaveFocus(),
  );

  fireEvent.keyDown(screen.getByRole('tab', { name: '风险台账' }), {
    key: 'End',
  });
  await waitFor(() => expect(calls).toEqual(['sheet-2', 'sheet-3']));
  await waitFor(() =>
    expect(screen.getByRole('tab', { name: '资源计划' })).toHaveFocus(),
  );

  fireEvent.keyDown(screen.getByRole('tab', { name: '资源计划' }), {
    key: 'Home',
  });
  await waitFor(() => expect(calls).toEqual(['sheet-2', 'sheet-3', 'sheet-1']));
  await waitFor(() => expect(first).toHaveFocus());
});

test('offers direct previous and next worksheet navigation', () => {
  const calls: string[] = [];
  render(
    <SpreadsheetSheetBar
      activeSheetId="sheet-1"
      editable
      sheets={[
        { id: 'sheet-1', name: '执行看板', status: 1 },
        { id: 'sheet-2', name: '风险台账', status: 0 },
        { id: 'sheet-3', name: '资源计划', status: 0 },
      ]}
      onActivate={(id) => calls.push(id)}
      onCreate={() => undefined}
      onDelete={() => undefined}
      onDuplicate={() => undefined}
      onHide={() => undefined}
      onMove={() => undefined}
      onRename={() => undefined}
      onSetColor={() => undefined}
      onShow={() => undefined}
    />,
  );

  const previous = screen.getByRole('button', { name: '上一个工作表' });
  const next = screen.getByRole('button', { name: '下一个工作表' });
  expect(previous).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+PageUp Meta+PageUp',
  );
  expect(next).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+PageDown Meta+PageDown',
  );
  fireEvent.click(previous);
  fireEvent.click(next);
  expect(calls).toEqual(['sheet-3', 'sheet-2']);
});

test('opens the worksheet menu from the standard keyboard and context-menu gestures', async () => {
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
      onRename={() => undefined}
      onSetColor={() => undefined}
      onShow={() => undefined}
    />,
  );

  const tab = screen.getByRole('tab', { name: '执行看板' });
  const options = screen.getByRole('button', { name: '执行看板选项' });
  expect(options).toHaveAttribute('aria-keyshortcuts', 'Shift+F10');

  fireEvent.keyDown(tab, { key: 'F10', shiftKey: true });
  await waitFor(() =>
    expect(
      screen.getByRole('menu', { name: '执行看板工作表操作' }),
    ).toBeInTheDocument(),
  );
  await waitFor(() =>
    expect(screen.getByRole('menuitem', { name: '重命名' })).toHaveFocus(),
  );
  fireEvent.keyDown(screen.getByRole('menu', { name: '执行看板工作表操作' }), {
    key: 'Escape',
  });
  await waitFor(() => expect(options).toHaveFocus());

  fireEvent.contextMenu(tab);
  await waitFor(() =>
    expect(
      screen.getByRole('menu', { name: '执行看板工作表操作' }),
    ).toBeInTheDocument(),
  );
});

test('keeps read-only worksheet navigation free of disabled edit controls', () => {
  render(
    <SpreadsheetSheetBar
      activeSheetId="sheet-1"
      editable={false}
      sheets={[
        { id: 'sheet-1', name: '执行看板', status: 1 },
        { id: 'sheet-2', name: '风险台账', status: 0 },
        { id: 'sheet-3', name: '隐藏底稿', status: 0, hide: 1 },
      ]}
      onActivate={() => undefined}
      onCreate={() => undefined}
      onDelete={() => undefined}
      onDuplicate={() => undefined}
      onHide={() => undefined}
      onMove={() => undefined}
      onRename={() => undefined}
      onSetColor={() => undefined}
      onShow={() => undefined}
    />,
  );

  expect(screen.queryByRole('button', { name: '新建工作表' })).toBeNull();
  expect(screen.getByRole('button', { name: '工作表列表' })).toBeEnabled();
  expect(screen.getByRole('tab', { name: '执行看板' })).toBeEnabled();
  expect(screen.queryByRole('button', { name: '执行看板选项' })).toBeNull();
  fireEvent.click(screen.getByRole('button', { name: '工作表列表' }));
  expect(screen.queryByText('隐藏底稿')).toBeNull();
});
