import { expect, test } from '@rstest/core';
import { fireEvent, render, screen } from '@testing-library/react';
import { useState } from 'react';
import { SpreadsheetChartLayoutEditor } from '../src/internal/features/work/editors/spreadsheet-chart-layout-editor';
import { SpreadsheetChartPanel } from '../src/internal/features/work/editors/spreadsheet-chart-panel';
import { SpreadsheetChartSeriesStyleEditor } from '../src/internal/features/work/editors/spreadsheet-chart-series-style-editor';
import { SpreadsheetErrorBarEditor } from '../src/internal/features/work/editors/spreadsheet-error-bar-editor';
import { SpreadsheetTrendlineEditor } from '../src/internal/features/work/editors/spreadsheet-trendline-editor';
import type {
  WorkSpreadsheetChart,
  WorkSpreadsheetChartLayout,
  WorkSpreadsheetChartSeriesStyle,
  WorkSpreadsheetContent,
  WorkSpreadsheetErrorBars,
  WorkSpreadsheetTrendline,
} from '../src/internal/features/work/work-types';

test('commits spreadsheet chart layout numbers after a complete edit', () => {
  const changes: Array<Partial<WorkSpreadsheetChartLayout>> = [];

  function Fixture() {
    const [chart, setChart] = useState<
      WorkSpreadsheetChartLayout & {
        type: 'column';
        showLegend: boolean;
      }
    >({
      type: 'column',
      showLegend: true,
      gapWidth: 150,
      overlap: 0,
    });
    return (
      <SpreadsheetChartLayoutEditor
        chart={chart}
        onChange={(change) => {
          changes.push(change);
          setChart((current) => ({ ...current, ...change }));
        }}
      />
    );
  }

  render(<Fixture />);
  const gap = screen.getByRole('textbox', { name: '分类间距（%）' });
  fireEvent.change(gap, { target: { value: '' } });
  expect(changes).toEqual([]);
  fireEvent.blur(gap);
  expect(gap).toHaveValue('150');

  fireEvent.change(gap, { target: { value: '185.' } });
  fireEvent.keyDown(gap, { key: 'Enter' });
  expect(changes).toEqual([{ gapWidth: 185 }]);
});

test('keeps series style measurements stable during intermediate input', () => {
  const changes: Array<WorkSpreadsheetChartSeriesStyle | undefined> = [];

  function Fixture() {
    const [style, setStyle] = useState<WorkSpreadsheetChartSeriesStyle>({
      fillTransparency: 0,
      lineWidth: 2.25,
      marker: { symbol: 'circle', size: 5 },
    });
    return (
      <SpreadsheetChartSeriesStyleEditor
        seriesNumber={1}
        supportsMarkers
        value={style}
        onChange={(next) => {
          changes.push(next);
          if (next) setStyle(next);
        }}
      />
    );
  }

  render(<Fixture />);
  const lineWidth = screen.getByRole('textbox', {
    name: '系列 1 线条宽度',
  });
  fireEvent.change(lineWidth, { target: { value: '' } });
  expect(changes).toEqual([]);
  fireEvent.blur(lineWidth);
  expect(lineWidth).toHaveValue('2.25');

  fireEvent.change(lineWidth, { target: { value: '3.75' } });
  fireEvent.keyDown(lineWidth, { key: 'Enter' });
  expect(changes.at(-1)?.lineWidth).toBe(3.75);
});

test('does not corrupt trendline and error-bar drafts while typing', () => {
  const trendlineChanges: WorkSpreadsheetTrendline[][] = [];
  const errorBarChanges: WorkSpreadsheetErrorBars[][] = [];

  function TrendlineFixture() {
    const [trendlines, setTrendlines] = useState<WorkSpreadsheetTrendline[]>([
      { type: 'polynomial', order: 2, forward: 0, backward: 0 },
    ]);
    return (
      <SpreadsheetTrendlineEditor
        seriesNumber={1}
        trendlines={trendlines}
        onChange={(next) => {
          trendlineChanges.push(next);
          setTrendlines(next);
        }}
      />
    );
  }

  function ErrorBarFixture() {
    const [errorBars, setErrorBars] = useState<WorkSpreadsheetErrorBars[]>([
      {
        direction: 'y',
        barType: 'both',
        valueType: 'fixedValue',
        value: 1,
      },
    ]);
    return (
      <SpreadsheetErrorBarEditor
        chartType="column"
        seriesNumber={1}
        errorBars={errorBars}
        onChange={(next) => {
          errorBarChanges.push(next);
          setErrorBars(next);
        }}
      />
    );
  }

  const trendlineView = render(<TrendlineFixture />);
  const order = screen.getByRole('textbox', {
    name: '系列 1 趋势线 1 阶数',
  });
  fireEvent.change(order, { target: { value: '' } });
  expect(trendlineChanges).toEqual([]);
  fireEvent.blur(order);
  expect(order).toHaveValue('2');
  fireEvent.change(order, { target: { value: '4.7' } });
  fireEvent.keyDown(order, { key: 'Enter' });
  expect(trendlineChanges.at(-1)?.[0].order).toBe(5);
  trendlineView.unmount();

  render(<ErrorBarFixture />);
  const errorValue = screen.getByRole('textbox', {
    name: '系列 1 误差线 1 数值',
  });
  fireEvent.change(errorValue, { target: { value: '' } });
  expect(errorBarChanges).toEqual([]);
  fireEvent.blur(errorValue);
  expect(errorValue).toHaveValue('1');
  fireEvent.change(errorValue, { target: { value: '2.5' } });
  fireEvent.keyDown(errorValue, { key: 'Enter' });
  expect(errorBarChanges.at(-1)?.[0].value).toBe(2.5);
});

test('rejects incomplete custom error values instead of silently dropping them', () => {
  const changes: WorkSpreadsheetErrorBars[][] = [];

  function Fixture() {
    const [errorBars, setErrorBars] = useState<WorkSpreadsheetErrorBars[]>([
      {
        direction: 'y',
        barType: 'both',
        valueType: 'custom',
        plusValues: [1, 2],
      },
    ]);
    return (
      <SpreadsheetErrorBarEditor
        chartType="column"
        seriesNumber={1}
        errorBars={errorBars}
        customInput="values"
        onChange={(next) => {
          changes.push(next);
          setErrorBars(next);
        }}
      />
    );
  }

  render(<Fixture />);
  const values = screen.getByRole('textbox', {
    name: '系列 1 误差线 1 正误差值',
  });
  fireEvent.change(values, { target: { value: '1, wrong, 2' } });
  expect(values).toHaveAttribute('aria-invalid', 'true');
  expect(changes).toEqual([]);

  fireEvent.keyDown(values, { key: 'Escape' });
  expect(values).toHaveValue('1, 2');
  expect(changes).toEqual([]);

  fireEvent.change(values, { target: { value: '1.5, 2.' } });
  fireEvent.keyDown(values, { key: 'Enter' });
  expect(values).toHaveValue('1.5, 2');
  expect(changes.at(-1)?.[0].plusValues).toEqual([1.5, 2]);
});

test('keeps spreadsheet doughnut sizing local until the chart is saved', () => {
  const chart: WorkSpreadsheetChart = {
    id: 'chart-1',
    name: '销售结构',
    type: 'doughnut',
    categories: ['产品 A'],
    series: [{ name: '收入', values: [12] }],
    showLegend: true,
    doughnutHoleSize: 50,
    left: 20,
    top: 20,
    width: 420,
    height: 260,
  };
  const content: WorkSpreadsheetContent = {
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-1',
        name: '工作表 1',
        status: 1,
        data: [],
        charts: [chart],
      },
    ],
  };
  const changes: WorkSpreadsheetContent[] = [];

  render(
    <SpreadsheetChartPanel
      content={content}
      activeSheetId="sheet-1"
      onChange={(next) => changes.push(next)}
    />,
  );

  const hole = screen.getByRole('textbox', { name: '圆环孔径（%）' });
  fireEvent.change(hole, { target: { value: '' } });
  expect(changes).toEqual([]);
  fireEvent.blur(hole);
  expect(hole).toHaveValue('50');

  fireEvent.change(hole, { target: { value: '64.6' } });
  fireEvent.keyDown(hole, { key: 'Enter' });
  expect(hole).toHaveValue('65');
  expect(changes).toEqual([]);
  fireEvent.click(screen.getByRole('button', { name: '保存图表' }));
  expect(changes.at(-1)?.sheets[0].charts?.[0].doughnutHoleSize).toBe(65);
});

test('preserves a dirty chart draft across unrelated workbook updates', () => {
  const chart: WorkSpreadsheetChart = {
    id: 'chart-1',
    name: '销售图表',
    type: 'column',
    categories: ['一月'],
    series: [{ name: '收入', values: [12] }],
    showLegend: true,
    left: 20,
    top: 20,
    width: 420,
    height: 260,
  };
  const content: WorkSpreadsheetContent = {
    type: 'spreadsheet',
    sheets: [
      {
        id: 'sheet-1',
        name: '工作表 1',
        status: 1,
        data: [],
        charts: [chart],
      },
    ],
  };
  const props = {
    activeSheetId: 'sheet-1',
    onChange: () => undefined,
  };
  const { rerender } = render(
    <SpreadsheetChartPanel content={content} {...props} />,
  );

  const name = screen.getByRole('textbox', { name: '图表对象名称' });
  fireEvent.change(name, { target: { value: '尚未保存的名称' } });
  rerender(
    <SpreadsheetChartPanel
      content={{
        ...content,
        sheets: [{ ...content.sheets[0], data: [[{ v: 12 }]] }],
      }}
      {...props}
    />,
  );

  expect(name).toHaveValue('尚未保存的名称');
});

test('does not silently discard a dirty chart when another chart is selected', () => {
  const charts: WorkSpreadsheetChart[] = [
    {
      id: 'chart-1',
      name: '销售图表',
      type: 'column',
      categories: ['一月'],
      series: [{ name: '收入', values: [12] }],
      showLegend: true,
      left: 20,
      top: 20,
      width: 420,
      height: 260,
    },
    {
      id: 'chart-2',
      name: '利润图表',
      type: 'line',
      categories: ['一月'],
      series: [{ name: '利润', values: [5] }],
      showLegend: true,
      left: 60,
      top: 60,
      width: 420,
      height: 260,
    },
  ];
  const changes: WorkSpreadsheetContent[] = [];
  render(
    <SpreadsheetChartPanel
      content={{
        type: 'spreadsheet',
        sheets: [
          {
            id: 'sheet-1',
            name: '工作表 1',
            status: 1,
            data: [],
            charts,
          },
        ],
      }}
      activeSheetId="sheet-1"
      onChange={(next) => changes.push(next)}
    />,
  );

  const name = screen.getByRole('textbox', { name: '图表对象名称' });
  fireEvent.change(name, { target: { value: '尚未保存的名称' } });
  fireEvent.click(screen.getByRole('button', { name: /利润图表/ }));

  expect(name).toHaveValue('尚未保存的名称');
  expect(screen.getByRole('alert')).toHaveTextContent('请先保存或取消');
  expect(changes).toEqual([]);

  fireEvent.click(screen.getByRole('button', { name: '取消更改' }));
  expect(name).toHaveValue('销售图表');
  fireEvent.click(screen.getByRole('button', { name: /利润图表/ }));
  expect(name).toHaveValue('利润图表');
});
