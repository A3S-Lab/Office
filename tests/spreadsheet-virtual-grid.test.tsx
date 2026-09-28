import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import type { CellMatrix } from '@fortune-sheet/core';
import { useState } from 'react';
import { SpreadsheetVirtualGrid } from '../src/internal/features/work/editors/spreadsheet-virtual-grid';

function Harness() {
  const [active, setActive] = useState({ row: 0, column: 0 });
  const data: CellMatrix = [];
  data[0] = [];
  data[0][0] = { v: 'A1', m: 'A1' };
  data[1] = [];
  data[1][0] = { v: 'A2', m: 'A2' };
  return (
    <SpreadsheetVirtualGrid
      sheetName="Sheet1"
      data={data}
      rowCount={60}
      columnCount={26}
      activeRow={active.row}
      activeColumn={active.column}
      onActiveCellChange={(row, column) => setActive({ row, column })}
    />
  );
}

test('SpreadsheetVirtualGrid owns grid ARIA and moves the active cell with arrows', () => {
  render(<Harness />);
  const grid = screen.getByRole('grid', { name: 'Sheet1' });
  expect(grid).toHaveAttribute('aria-activedescendant', 'a3s-ss-cell-0-0');
  expect(grid).toHaveAttribute('aria-rowcount', '60');
  expect(grid).toHaveAttribute('aria-colcount', '26');

  grid.focus();
  fireEvent.keyDown(grid, { key: 'ArrowDown' });
  expect(grid).toHaveAttribute('aria-activedescendant', 'a3s-ss-cell-1-0');
  fireEvent.keyDown(grid, { key: 'ArrowRight' });
  expect(grid).toHaveAttribute('aria-activedescendant', 'a3s-ss-cell-1-1');
});

test('SpreadsheetVirtualGrid commits typed literals through onCommitCell', () => {
  const commits: Array<{ row: number; column: number; raw: string }> = [];
  const data: CellMatrix = [];
  render(
    <SpreadsheetVirtualGrid
      sheetName="Sheet1"
      data={data}
      rowCount={10}
      columnCount={10}
      activeRow={0}
      activeColumn={0}
      onActiveCellChange={() => undefined}
      onCommitCell={(row, column, raw) => commits.push({ row, column, raw })}
    />,
  );
  const grid = screen.getByRole('grid', { name: 'Sheet1' });
  grid.focus();
  fireEvent.keyDown(grid, { key: 'F2' });
  const editor = screen.getByLabelText('Sheet1 cell editor');
  fireEvent.change(editor, { target: { value: '99' } });
  fireEvent.keyDown(editor, { key: 'Enter' });
  expect(commits).toEqual([{ row: 0, column: 0, raw: '99' }]);
});
