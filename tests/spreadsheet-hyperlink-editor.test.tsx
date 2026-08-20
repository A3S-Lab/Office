import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { useState } from 'react';
import type { SpreadsheetContent } from '../src/core';
import { SpreadsheetEditor } from '../src/react';

test('commits a hyperlink once, restores exact focus, and undoes in one step', async () => {
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
  const launcher = screen.getByRole('button', { name: '超链接' });
  launcher.focus();
  fireEvent.click(launcher);
  expect(
    screen.getByRole('dialog', { name: '插入超链接' }),
  ).toBeInTheDocument();
  fireEvent.click(screen.getByRole('button', { name: '取消' }));

  await waitFor(() =>
    expect(screen.queryByRole('dialog', { name: '插入超链接' })).toBeNull(),
  );
  await waitFor(() => expect(launcher).toHaveFocus());
  expect(changes).toHaveLength(0);

  grid.focus();
  fireEvent.keyDown(grid, { key: 'ArrowDown' });
  fireEvent.keyDown(grid, { key: 'ArrowRight' });
  await waitFor(() =>
    expect(container.querySelector('.fortune-name-box')).toHaveTextContent(
      'B2',
    ),
  );
  fireEvent.keyDown(grid, { ctrlKey: true, key: 'k' });

  expect(screen.getByRole('dialog', { name: '插入超链接' })).toHaveTextContent(
    'Inputs!B2',
  );
  expect(screen.getByRole('textbox', { name: '显示文本' })).toHaveValue(
    'A3S Office',
  );
  fireEvent.change(screen.getByRole('textbox', { name: '地址' }), {
    target: { value: 'a3s.dev/office' },
  });
  fireEvent.click(screen.getByRole('button', { name: '确定' }));

  await waitFor(() =>
    expect(screen.getByTestId('hyperlink-address')).toHaveTextContent(
      'https://a3s.dev/office',
    ),
  );
  expect(screen.getByTestId('hyperlink-cell')).toHaveTextContent(
    'A3S Office:#fff2cc:1:linked',
  );
  expect(screen.getByTestId('hyperlink-selection')).toHaveTextContent('1:1');
  expect(changes).toHaveLength(1);
  await waitFor(() =>
    expect(
      container.querySelector<HTMLElement>('.fortune-sheet-overlay'),
    ).toHaveFocus(),
  );

  const undo = screen.getByRole('button', { name: '撤销' });
  await waitFor(() => expect(undo).toBeEnabled());
  fireEvent.click(undo);

  await waitFor(() =>
    expect(screen.getByTestId('hyperlink-address')).toHaveTextContent('none'),
  );
  expect(screen.getByTestId('hyperlink-cell')).toHaveTextContent(
    'A3S Office:#fff2cc:1:none',
  );
  expect(screen.getByTestId('hyperlink-selection')).toHaveTextContent('1:1');
  expect(changes).toHaveLength(2);
  await waitFor(() => expect(undo).toBeDisabled());
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
          [null, { v: 'A3S Office', m: 'A3S Office', bg: '#fff2cc', bl: 1 }],
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
      <output data-testid="hyperlink-address">
        {sheet?.hyperlink?.['1_1']?.linkAddress ?? 'none'}
      </output>
      <output data-testid="hyperlink-cell">
        {`${cell?.m ?? 'none'}:${cell?.bg ?? 'none'}:${cell?.bl ?? 'none'}:${cell?.hl ? 'linked' : 'none'}`}
      </output>
      <output data-testid="hyperlink-selection">
        {selection
          ? `${selection.row[0] ?? 0}:${selection.column[0] ?? 0}`
          : 'none'}
      </output>
    </>
  );
}
