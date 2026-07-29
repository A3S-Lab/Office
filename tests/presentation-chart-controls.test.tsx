import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import { useState } from 'react';
import { PresentationChartLayoutEditor } from '../src/internal/features/work/editors/presentation-chart-layout-editor';
import { PresentationChartPanel } from '../src/internal/features/work/editors/presentation-chart-panel';
import type { WorkSlideChart } from '../src/internal/features/work/work-types';

test('commits presentation chart spacing only after complete input', () => {
  const changes: WorkSlideChart[] = [];

  function Fixture() {
    const [chart, setChart] = useState<WorkSlideChart>({
      type: 'column',
      categories: ['一月'],
      series: [{ name: '收入', values: [1] }],
      gapWidth: 150,
      overlap: 0,
    });
    return (
      <PresentationChartLayoutEditor
        chart={chart}
        onChange={(next) => {
          changes.push(next);
          setChart(next);
        }}
      />
    );
  }

  render(<Fixture />);
  const gap = screen.getByRole('textbox', {
    name: '演示图表分类间距（%）',
  });
  fireEvent.change(gap, { target: { value: '' } });
  expect(changes).toEqual([]);
  fireEvent.blur(gap);
  expect(gap).toHaveValue('150');

  fireEvent.change(gap, { target: { value: '185.' } });
  fireEvent.keyDown(gap, { key: 'Enter' });
  expect(gap).toHaveValue('185');
  expect(changes.at(-1)?.gapWidth).toBe(185);
});

test('keeps an incomplete doughnut hole size out of the chart model', () => {
  const changes: WorkSlideChart[] = [];

  function Fixture() {
    const [chart, setChart] = useState<WorkSlideChart>({
      type: 'doughnut',
      categories: ['产品 A'],
      series: [{ name: '收入', values: [1] }],
      doughnutHoleSize: 50,
    });
    return (
      <PresentationChartPanel
        chart={chart}
        onChange={(next) => {
          changes.push(next);
          setChart(next);
        }}
        onDelete={() => undefined}
        onClose={() => undefined}
      />
    );
  }

  render(<Fixture />);
  const hole = screen.getByRole('textbox', { name: '圆环孔径' });
  fireEvent.change(hole, { target: { value: '' } });
  expect(changes).toEqual([]);
  fireEvent.blur(hole);
  expect(hole).toHaveValue('50');

  fireEvent.change(hole, { target: { value: '72.6' } });
  fireEvent.keyDown(hole, { key: 'Enter' });
  expect(hole).toHaveValue('73');
  expect(changes.at(-1)?.doughnutHoleSize).toBe(73);
});

test('commits presentation value-axis bounds without corrupting intermediate input', () => {
  const changes: WorkSlideChart[] = [];

  function Fixture() {
    const [chart, setChart] = useState<WorkSlideChart>({
      type: 'column',
      categories: ['一月'],
      series: [{ name: '收入', values: [12] }],
      axes: { left: { minimum: 10, maximum: 20, majorUnit: 2 } },
    });
    return (
      <PresentationChartPanel
        chart={chart}
        onChange={(next) => {
          changes.push(next);
          setChart(next);
        }}
        onDelete={() => undefined}
        onClose={() => undefined}
      />
    );
  }

  render(<Fixture />);
  const minimum = screen.getByRole('textbox', {
    name: '演示图表纵轴最小值',
  });
  fireEvent.change(minimum, { target: { value: '-' } });
  expect(changes).toEqual([]);
  fireEvent.blur(minimum);
  expect(minimum).toHaveValue('10');
  expect(changes).toEqual([]);

  fireEvent.change(minimum, { target: { value: '12.' } });
  fireEvent.keyDown(minimum, { key: 'Enter' });
  expect(minimum).toHaveValue('12');
  expect(changes.at(-1)?.axes?.left?.minimum).toBe(12);

  fireEvent.change(minimum, { target: { value: '' } });
  fireEvent.blur(minimum);
  expect(minimum).toHaveValue('');
  expect(changes.at(-1)?.axes?.left?.minimum).toBeUndefined();
});

test('rejects presentation axis bounds that invert the visible range', () => {
  const changes: WorkSlideChart[] = [];
  const chart: WorkSlideChart = {
    type: 'column',
    categories: ['一月'],
    series: [{ name: '收入', values: [12] }],
    axes: { left: { minimum: 10, maximum: 20, majorUnit: 2 } },
  };

  render(
    <PresentationChartPanel
      chart={chart}
      onChange={(next) => changes.push(next)}
      onDelete={() => undefined}
      onClose={() => undefined}
    />,
  );

  const maximum = screen.getByRole('textbox', {
    name: '演示图表纵轴最大值',
  });
  fireEvent.change(maximum, { target: { value: '5' } });
  expect(maximum).toHaveAttribute('aria-invalid', 'true');
  fireEvent.keyDown(maximum, { key: 'Enter' });

  expect(maximum).toHaveValue('20');
  expect(maximum).not.toHaveAttribute('aria-invalid');
  expect(changes).toEqual([]);
});

test('preserves chart list drafts and rejects invalid series data', () => {
  const changes: WorkSlideChart[] = [];
  let closes = 0;

  function Fixture() {
    const [chart, setChart] = useState<WorkSlideChart>({
      type: 'column',
      categories: ['一月'],
      series: [{ name: '收入', values: [1] }],
    });
    return (
      <PresentationChartPanel
        chart={chart}
        onChange={(next) => {
          changes.push(next);
          setChart(next);
        }}
        onDelete={() => undefined}
        onClose={() => closes++}
      />
    );
  }

  render(<Fixture />);
  const categories = screen.getByRole('textbox', {
    name: '演示图表分类',
  });
  fireEvent.change(categories, { target: { value: '一月\n' } });
  expect(categories).toHaveValue('一月\n');
  expect(changes).toEqual([]);
  fireEvent.change(categories, { target: { value: '一月\n二月' } });
  fireEvent.blur(categories);
  expect(changes.at(-1)?.categories).toEqual(['一月', '二月']);

  const values = screen.getByRole('textbox', {
    name: '演示图表系列 1 数据',
  });
  const committedCount = changes.length;
  fireEvent.change(values, { target: { value: '1.5, 2.' } });
  expect(values).toHaveValue('1.5, 2.');
  expect(changes).toHaveLength(committedCount);
  fireEvent.blur(values);
  expect(values).toHaveValue('1.5, 2');
  expect(changes.at(-1)?.series[0].values).toEqual([1.5, 2]);

  const validCount = changes.length;
  fireEvent.change(values, { target: { value: '1.5, wrong' } });
  expect(values).toHaveAttribute('aria-invalid', 'true');
  fireEvent.keyDown(values, { key: 'Escape' });
  expect(values).toHaveValue('1.5, 2');
  expect(changes).toHaveLength(validCount);
  expect(closes).toBe(0);

  fireEvent.keyDown(values, { key: 'Escape' });
  expect(closes).toBe(1);
});
