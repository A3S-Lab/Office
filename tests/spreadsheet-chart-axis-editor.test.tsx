import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import { useState } from 'react';
import { SpreadsheetChartAxisEditor } from '../src/internal/features/work/editors/spreadsheet-chart-axis-editor';
import type { WorkSpreadsheetChartAxes } from '../src/internal/features/work/work-types';

test('commits spreadsheet axis values only after a complete edit', () => {
  const changes: WorkSpreadsheetChartAxes[] = [];

  function Fixture() {
    const [axes, setAxes] = useState<WorkSpreadsheetChartAxes>({
      left: { minimum: 10, maximum: 20, majorUnit: 2 },
    });
    return (
      <SpreadsheetChartAxisEditor
        axes={axes}
        chartType="column"
        showSecondaryAxes={false}
        onChange={(next) => {
          changes.push(next);
          setAxes(next);
        }}
      />
    );
  }

  render(<Fixture />);
  const minimum = screen.getByRole('textbox', {
    name: '纵坐标轴最小值',
  });
  fireEvent.change(minimum, { target: { value: '-' } });
  expect(changes).toEqual([]);
  fireEvent.blur(minimum);
  expect(minimum).toHaveValue('10');

  fireEvent.change(minimum, { target: { value: '12.5' } });
  fireEvent.keyDown(minimum, { key: 'Enter' });
  expect(changes.at(-1)?.left?.minimum).toBe(12.5);

  const majorUnit = screen.getByRole('textbox', {
    name: '纵坐标轴主单位',
  });
  fireEvent.change(majorUnit, { target: { value: '0' } });
  fireEvent.blur(majorUnit);
  expect(majorUnit).toHaveValue('2');
  expect(changes).toHaveLength(1);
});

test('rejects spreadsheet axis bounds that invert the visible range', () => {
  const changes: WorkSpreadsheetChartAxes[] = [];

  render(
    <SpreadsheetChartAxisEditor
      axes={{ left: { minimum: 10, maximum: 20, majorUnit: 2 } }}
      chartType="column"
      showSecondaryAxes={false}
      onChange={(next) => changes.push(next)}
    />,
  );

  const minimum = screen.getByRole('textbox', {
    name: '纵坐标轴最小值',
  });
  fireEvent.change(minimum, { target: { value: '25' } });
  expect(minimum).toHaveAttribute('aria-invalid', 'true');
  fireEvent.keyDown(minimum, { key: 'Enter' });

  expect(minimum).toHaveValue('10');
  expect(minimum).not.toHaveAttribute('aria-invalid');
  expect(changes).toEqual([]);
});
