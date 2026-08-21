import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { useState } from 'react';
import type { SpreadsheetContent } from '../src/core';
import { SpreadsheetEditor } from '../src/react';

test('creates one controlled table, enters Table Design, and undoes once', async () => {
  const changes: SpreadsheetContent[] = [];
  const { container } = render(<ControlledSpreadsheet changes={changes} />);
  const grid = await waitFor(() => {
    const target = container.querySelector<HTMLElement>(
      '.fortune-sheet-overlay',
    );
    expect(target).not.toBeNull();
    return target as HTMLElement;
  });

  fireEvent.click(screen.getByRole('tab', { name: '插入' }));
  const launcher = screen.getByRole('button', { name: '表格' });
  expect(launcher).toHaveAttribute('aria-keyshortcuts', 'Control+T Meta+T');
  launcher.focus();
  fireEvent.click(launcher);
  expect(screen.getByRole('dialog', { name: '创建表格' })).toBeInTheDocument();
  expect(screen.getByRole('textbox', { name: '表格区域' })).toHaveValue(
    'A1:C3',
  );
  fireEvent.click(screen.getByRole('button', { name: '取消' }));
  await waitFor(() => expect(launcher).toHaveFocus());
  expect(changes).toHaveLength(0);

  grid.focus();
  fireEvent.keyDown(grid, { ctrlKey: true, key: 't' });
  fireEvent.click(screen.getByRole('button', { name: '确定' }));

  await waitFor(() =>
    expect(screen.getByTestId('table-state')).toHaveTextContent(
      'Table1:medium:2',
    ),
  );
  expect(screen.getByTestId('table-selection')).toHaveTextContent('1:1');
  expect(changes).toHaveLength(1);
  await waitFor(() =>
    expect(screen.getByRole('tab', { name: '表格设计' })).toHaveAttribute(
      'aria-selected',
      'true',
    ),
  );
  await waitFor(() =>
    expect(
      container.querySelector<HTMLElement>('.fortune-sheet-overlay'),
    ).toHaveFocus(),
  );

  const undo = screen.getByRole('button', { name: '撤销' });
  await waitFor(() => expect(undo).toBeEnabled());
  fireEvent.click(undo);
  await waitFor(() =>
    expect(screen.getByTestId('table-state')).toHaveTextContent('none'),
  );
  expect(changes).toHaveLength(2);
});

function ControlledSpreadsheet({ changes }: { changes: SpreadsheetContent[] }) {
  const [content, setContent] = useState<SpreadsheetContent>(() => ({
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-1',
        name: 'Sales',
        status: 1,
        row: 8,
        column: 6,
        data: [
          [{ v: 'Region' }, { v: 'Revenue' }, { v: 'Status' }],
          [{ v: 'East' }, { v: 10 }, { v: 'Ready' }],
          [{ v: 'West' }, { v: 12 }, { v: 'Blocked' }],
        ],
        luckysheet_select_save: [
          {
            row: [1, 1],
            column: [1, 1],
            row_focus: 1,
            column_focus: 1,
          },
        ],
      },
    ],
  }));
  const sheet = content.sheets[0];
  const table = sheet?.tables?.[0];
  const selection = sheet?.luckysheet_select_save?.at(-1);
  return (
    <>
      <SpreadsheetEditor
        content={content}
        onChange={(next) => {
          changes.push(next);
          setContent(next);
        }}
        theme="light"
      />
      <output data-testid="table-state">
        {table
          ? `${table.name}:${table.style.family}:${table.style.family === 'none' ? 0 : table.style.number}`
          : 'none'}
      </output>
      <output data-testid="table-selection">
        {selection
          ? `${selection.row_focus ?? selection.row[0]}:${selection.column_focus ?? selection.column[0]}`
          : 'none'}
      </output>
    </>
  );
}
