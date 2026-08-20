import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { useState } from 'react';
import type { SpreadsheetContent } from '../src/core';
import { SpreadsheetEditor } from '../src/react';

test('commits Alt+= once, preserves target style, and undoes in one step', async () => {
  const changes: SpreadsheetContent[] = [];
  const { container } = render(<ControlledAutoSum changes={changes} />);
  const grid = await waitFor(() => {
    const target = container.querySelector<HTMLElement>(
      '.fortune-sheet-overlay',
    );
    expect(target).not.toBeNull();
    return target as HTMLElement;
  });

  grid.focus();
  for (let index = 0; index < 3; index += 1) {
    fireEvent.keyDown(grid, { key: 'ArrowDown' });
  }
  await waitFor(() =>
    expect(container.querySelector('.fortune-name-box')).toHaveTextContent(
      'A4',
    ),
  );
  fireEvent.keyDown(grid, { altKey: true, code: 'Equal', key: '=' });

  await waitFor(() =>
    expect(screen.getByTestId('auto-sum-formula')).toHaveTextContent(
      '=SUM(A2:A3)',
    ),
  );
  expect(screen.getByTestId('auto-sum-style')).toHaveTextContent('#fff2cc:1');
  expect(screen.getByTestId('auto-sum-selection')).toHaveTextContent('3:0');
  expect(changes).toHaveLength(1);
  await waitFor(() => expect(grid).toHaveFocus());

  const undo = screen.getByRole('button', { name: '撤销' });
  await waitFor(() => expect(undo).toBeEnabled());
  fireEvent.click(undo);

  await waitFor(() =>
    expect(screen.getByTestId('auto-sum-formula')).toHaveTextContent('none'),
  );
  expect(screen.getByTestId('auto-sum-style')).toHaveTextContent('#fff2cc:1');
  expect(changes).toHaveLength(2);
  await waitFor(() => expect(undo).toBeDisabled());
});

function ControlledAutoSum({ changes }: { changes: SpreadsheetContent[] }) {
  const [content, setContent] = useState<SpreadsheetContent>(() => ({
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-1',
        name: 'Revenue',
        status: 1,
        row: 4,
        column: 2,
        data: [
          [{ v: 'Revenue' }, null],
          [{ v: 12 }, null],
          [{ v: 18 }, null],
          [{ bg: '#fff2cc', bl: 1 }, null],
        ],
      },
    ],
  }));
  const sheet = content.sheets[0];
  const target = sheet?.data?.[3]?.[0];
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
      <output data-testid="auto-sum-formula">{target?.f ?? 'none'}</output>
      <output data-testid="auto-sum-style">
        {`${target?.bg ?? 'none'}:${target?.bl ?? 'none'}`}
      </output>
      <output data-testid="auto-sum-selection">
        {selection
          ? `${selection.row[0] ?? 0}:${selection.column[0] ?? 0}`
          : 'none'}
      </output>
    </>
  );
}
