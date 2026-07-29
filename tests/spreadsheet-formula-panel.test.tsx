import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import { SpreadsheetFormulaPanel } from '../src/internal/features/work/editors/spreadsheet-formula-panel';
import type { WorkSpreadsheetContent } from '../src/internal/features/work/work-types';

const content: WorkSpreadsheetContent = {
  type: 'spreadsheet',
  sheets: [{ id: 'sheet-1', name: '工作表 1', data: [] }],
  calculation: {
    mode: 'automatic',
    fullCalculationOnLoad: false,
    forceFullCalculation: false,
    iterativeCalculation: true,
    maximumIterations: 100,
    maximumChange: 0.001,
    fullPrecision: true,
  },
};

test('keeps incomplete calculation settings out of the saved workbook', () => {
  const changes: WorkSpreadsheetContent[] = [];
  render(
    <SpreadsheetFormulaPanel
      content={content}
      canRecalculateSelection
      canRecalculateWorkbook
      onChange={(next) => changes.push(next)}
      onRecalculate={() => true}
    />,
  );

  const iterations = screen.getByRole('textbox', {
    name: '最大迭代次数',
  });
  fireEvent.change(iterations, { target: { value: '' } });
  expect(iterations).toHaveValue('');
  expect(changes).toEqual([]);
  fireEvent.blur(iterations);
  expect(iterations).toHaveValue('100');

  fireEvent.change(iterations, { target: { value: '240.' } });
  fireEvent.keyDown(iterations, { key: 'Enter' });
  expect(iterations).toHaveValue('240');
  expect(changes).toEqual([]);

  fireEvent.click(screen.getByRole('button', { name: '保存计算设置' }));
  expect(changes).toHaveLength(1);
  expect(changes[0].calculation?.maximumIterations).toBe(240);
});

test('cancels dirty calculation settings before Escape can close the task pane', () => {
  const changes: WorkSpreadsheetContent[] = [];
  render(
    <SpreadsheetFormulaPanel
      content={content}
      canRecalculateSelection
      canRecalculateWorkbook
      onChange={(next) => changes.push(next)}
      onRecalculate={() => true}
    />,
  );

  const fullPrecision = screen.getByRole('checkbox', {
    name: '使用完整精度',
  });
  const save = screen.getByRole('button', { name: '保存计算设置' });
  const cancel = screen.getByRole('button', { name: '取消更改' });
  expect(save).toBeDisabled();
  expect(cancel).toBeDisabled();

  fireEvent.click(fullPrecision);
  expect(fullPrecision).not.toBeChecked();
  expect(save).toBeEnabled();
  expect(cancel).toBeEnabled();

  fireEvent.keyDown(fullPrecision, { key: 'Escape' });
  expect(fullPrecision).toBeChecked();
  expect(save).toBeDisabled();
  expect(cancel).toBeDisabled();
  expect(changes).toEqual([]);
});
