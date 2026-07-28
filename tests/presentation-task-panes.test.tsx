import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { useState } from 'react';
import { PresentationChartPanel } from '../src/internal/features/work/editors/presentation-chart-panel';
import { PresentationEditor } from '../src/internal/features/work/editors/presentation-editor';
import type {
  WorkPresentationContent,
  WorkSlideChart,
} from '../src/internal/features/work/work-types';

test('keeps presentation task panes exclusive and dismisses them with Escape', () => {
  render(<PresentationHarness />);

  fireEvent.click(screen.getByRole('tab', { name: '设计' }));
  fireEvent.click(screen.getByRole('button', { name: '母版和版式' }));
  expect(screen.getByRole('region', { name: '母版与布局' })).toBeVisible();

  fireEvent.keyDown(document, { key: 'Escape' });
  expect(
    screen.queryByRole('region', { name: '母版与布局' }),
  ).not.toBeInTheDocument();

  fireEvent.click(screen.getByRole('button', { name: '母版和版式' }));
  fireEvent.click(screen.getByRole('tab', { name: '审阅' }));
  fireEvent.click(screen.getByRole('button', { name: '查看批注' }));

  expect(
    screen.queryByRole('region', { name: '母版与布局' }),
  ).not.toBeInTheDocument();
  expect(screen.getByRole('region', { name: '演示批注审阅' })).toBeVisible();

  fireEvent.keyDown(document, { key: 'Escape' });
  expect(
    screen.queryByRole('region', { name: '演示批注审阅' }),
  ).not.toBeInTheDocument();
});

test('requires trimmed text before adding a presentation comment', async () => {
  render(<PresentationHarness />);

  fireEvent.click(screen.getByRole('tab', { name: '审阅' }));
  fireEvent.click(screen.getByRole('button', { name: '新建批注' }));

  const input = screen.getByRole('textbox', { name: '批注内容' });
  const submit = screen.getByRole('button', { name: '添加批注' });
  expect(submit).toBeDisabled();

  fireEvent.change(input, { target: { value: '   ' } });
  expect(submit).toBeDisabled();
  expect(screen.getByRole('alert')).toHaveTextContent('请输入批注内容。');

  fireEvent.change(input, { target: { value: '  需要补充来源  ' } });
  expect(submit).toBeEnabled();
  fireEvent.click(submit);

  await waitFor(() =>
    expect(screen.getByRole('textbox', { name: '编辑演示批注 1' })).toHaveValue(
      '需要补充来源',
    ),
  );
});

test('selects a newly inserted chart and opens its task pane', () => {
  render(<PresentationHarness />);

  fireEvent.click(screen.getByRole('tab', { name: '插入' }));
  fireEvent.click(screen.getByRole('button', { name: '图表' }));

  expect(screen.getByRole('region', { name: '演示图表数据' })).toBeVisible();
});

test('dismisses the presentation chart task pane with Escape', () => {
  let closes = 0;
  render(
    <PresentationChartPanel
      chart={chart()}
      onChange={() => undefined}
      onDelete={() => undefined}
      onClose={() => {
        closes += 1;
      }}
    />,
  );

  fireEvent.keyDown(screen.getByRole('region', { name: '演示图表数据' }), {
    key: 'Escape',
  });
  expect(closes).toBe(1);
});

function PresentationHarness() {
  const [content, setContent] = useState<WorkPresentationContent>(() =>
    presentation(),
  );
  return (
    <PresentationEditor
      content={content}
      preview={false}
      onChange={setContent}
    />
  );
}

function presentation(): WorkPresentationContent {
  return {
    type: 'presentation',
    slides: [{ id: 'slide-1', name: 'Title slide', elements: [] }],
  };
}

function chart(): WorkSlideChart {
  return {
    type: 'column',
    title: 'Quarterly revenue',
    categories: ['Q1', 'Q2'],
    series: [{ name: 'Revenue', values: [12, 18] }],
  };
}
