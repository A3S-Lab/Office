import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { useState } from 'react';
import type { SpreadsheetContent } from '../src/core';
import { SpreadsheetEditor } from '../src/react';

test('commits Format Cells once, restores focus, and undoes in one step', async () => {
  const changes: SpreadsheetContent[] = [];
  const { container } = render(<ControlledSpreadsheet changes={changes} />);
  const grid = await waitFor(() => {
    const target = container.querySelector<HTMLElement>(
      '.fortune-sheet-overlay',
    );
    expect(target).not.toBeNull();
    return target as HTMLElement;
  });
  const launcher = screen.getByRole('button', {
    name: '设置单元格格式',
  });

  launcher.focus();
  fireEvent.click(launcher);
  const firstDialog = screen.getByRole('dialog', {
    name: '设置单元格格式',
  });
  fireEvent.keyDown(firstDialog, { key: 'Escape' });

  await waitFor(() =>
    expect(screen.queryByRole('dialog', { name: '设置单元格格式' })).toBeNull(),
  );
  await waitFor(() => expect(launcher).toHaveFocus());
  expect(changes).toHaveLength(0);
  expect(screen.getByRole('button', { name: '撤销' })).toBeDisabled();

  grid.focus();
  fireEvent.keyDown(grid, { key: 'ArrowDown' });
  fireEvent.keyDown(grid, { key: 'ArrowRight' });
  await waitFor(() =>
    expect(container.querySelector('.fortune-name-box')).toHaveTextContent(
      'B2',
    ),
  );
  fireEvent.keyDown(grid, { ctrlKey: true, key: '1' });
  expect(
    screen.getByRole('dialog', { name: '设置单元格格式' }),
  ).toBeInTheDocument();
  fireEvent.change(screen.getByRole('textbox', { name: '数字格式代码' }), {
    target: { value: '0.000' },
  });
  fireEvent.click(screen.getByRole('button', { name: '应用' }));

  await waitFor(() =>
    expect(screen.getByTestId('format-cells-value')).toHaveTextContent('0.000'),
  );
  expect(screen.getByTestId('format-cells-selection')).toHaveTextContent('1:1');
  await waitFor(() =>
    expect(
      container.querySelector<HTMLElement>('.fortune-sheet-overlay'),
    ).toHaveFocus(),
  );
  expect(changes).toHaveLength(1);

  const undo = screen.getByRole('button', { name: '撤销' });
  await waitFor(() => expect(undo).toBeEnabled());
  fireEvent.click(undo);

  await waitFor(() =>
    expect(screen.getByTestId('format-cells-value')).toHaveTextContent(
      'General',
    ),
  );
  expect(changes).toHaveLength(2);
  await waitFor(() => expect(undo).toBeDisabled());
  expect(screen.getByTestId('format-cells-selection')).toHaveTextContent('1:1');
});

function ControlledSpreadsheet({ changes }: { changes: SpreadsheetContent[] }) {
  const [content, setContent] = useState<SpreadsheetContent>(() => ({
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-1',
        name: 'Inputs',
        status: 1,
        row: 2,
        column: 2,
        data: [
          [{ v: 'Label', m: 'Label' }, null],
          [null, { v: 12.5, m: '12.5' }],
        ],
      },
    ],
  }));
  const sheet = content.sheets[0];
  const cell = sheet?.data?.[1]?.[1];
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
      <output data-testid="format-cells-value">
        {cell?.ct?.fa ?? 'General'}
      </output>
      <output data-testid="format-cells-selection">
        {selection
          ? `${selection.row[0] ?? 0}:${selection.column[0] ?? 0}`
          : 'none'}
      </output>
    </>
  );
}
