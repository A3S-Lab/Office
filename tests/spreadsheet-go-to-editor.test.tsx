import { expect, test } from '@rstest/core';
import {
  fireEvent,
  render,
  screen,
  waitFor,
  within,
} from '@testing-library/react';
import { useState } from 'react';
import type { SpreadsheetContent } from '../src/core';
import { SpreadsheetEditor } from '../src/react';

test('navigates to a cross-sheet named range without changing workbook content or history', async () => {
  const changes: SpreadsheetContent[] = [];
  const { container } = render(<ControlledGoToSpreadsheet changes={changes} />);
  const grid = await waitFor(() => {
    const target = container.querySelector<HTMLElement>(
      '.fortune-sheet-overlay',
    );
    expect(target).not.toBeNull();
    return target as HTMLElement;
  });
  const trigger = screen.getByRole('button', { name: '查找和选择' });

  trigger.focus();
  fireEvent.click(trigger);
  fireEvent.click(
    within(screen.getByRole('menu', { name: '查找和选择选项' })).getByRole(
      'menuitem',
      { name: '定位' },
    ),
  );
  expect(screen.getByRole('dialog', { name: '定位' })).toBeInTheDocument();
  fireEvent.click(screen.getByRole('button', { name: '取消' }));
  await waitFor(() =>
    expect(screen.queryByRole('dialog', { name: '定位' })).toBeNull(),
  );
  await waitFor(() => expect(trigger).toHaveFocus());
  expect(changes).toHaveLength(0);

  grid.focus();
  fireEvent.keyDown(grid, { ctrlKey: true, key: 'f' });
  const findQuery = screen.getByRole('textbox', { name: '查找当前工作表' });
  await waitFor(() => expect(findQuery).toHaveFocus());
  fireEvent.keyDown(findQuery, { ctrlKey: true, key: 'f' });
  await waitFor(() => expect(findQuery).toHaveFocus());
  fireEvent.keyDown(findQuery, { key: 'Escape' });
  await waitFor(() =>
    expect(
      screen.queryByRole('textbox', { name: '查找当前工作表' }),
    ).toBeNull(),
  );
  await waitFor(() => expect(grid).toHaveFocus());

  fireEvent.keyDown(grid, { key: 'F5' });
  const reference = screen.getByRole('textbox', { name: '引用位置' });
  expect(reference).toHaveValue('A1');
  fireEvent.change(reference, { target: { value: 'A50' } });
  expect(
    screen.getByText('引用超出了目标工作表的有效边界。'),
  ).toBeInTheDocument();
  expect(screen.getByRole('button', { name: '定位' })).toBeDisabled();

  fireEvent.change(reference, { target: { value: 'ArchiveBlock' } });
  fireEvent.click(screen.getByRole('button', { name: '定位' }));

  const archiveTab = screen.getByRole('tab', { name: 'Archive 2025' });
  await waitFor(() =>
    expect(archiveTab).toHaveAttribute('aria-selected', 'true'),
  );
  await waitFor(() =>
    expect(container.querySelector('.fortune-name-box')).toHaveTextContent(
      'C9:E12',
    ),
  );
  expect(
    screen.getByRole('status', { name: '表格选区状态' }),
  ).toHaveTextContent('C9:E12');
  expect(screen.getByTestId('content-active-sheet')).toHaveTextContent(
    'sheet-1',
  );
  expect(changes).toHaveLength(0);
  expect(screen.getByRole('button', { name: '撤销' })).toBeDisabled();
  await waitFor(() => expect(grid).toHaveFocus());
});

function ControlledGoToSpreadsheet({
  changes,
}: {
  changes: SpreadsheetContent[];
}) {
  const [content, setContent] = useState<SpreadsheetContent>(() => ({
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-1',
        name: 'Inputs',
        status: 1,
        row: 12,
        column: 8,
        data: spreadsheetMatrix(12, 8, 'Input'),
      },
      {
        id: 'sheet-2',
        name: 'Archive 2025',
        status: 0,
        row: 20,
        column: 8,
        data: spreadsheetMatrix(20, 8, 'Archive'),
      },
    ],
    namedRanges: [
      {
        id: 'archive-block',
        name: 'ArchiveBlock',
        reference: "'Archive 2025'!$C$9:$E$12",
      },
    ],
  }));
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
      <output data-testid="content-active-sheet">
        {content.sheets.find((sheet) => sheet.status === 1)?.id ?? 'none'}
      </output>
    </>
  );
}

function spreadsheetMatrix(rows: number, columns: number, value: string) {
  const matrix = Array.from({ length: rows }, () =>
    Array.from({ length: columns }, () => null),
  );
  const firstRow = matrix[0];
  if (firstRow) firstRow[0] = { v: value, m: value };
  return matrix;
}
