import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { useRef, useState } from 'react';
import { PresentationChartPanel } from '../src/internal/features/work/editors/presentation-chart-panel';
import { PresentationCommentsPanel } from '../src/internal/features/work/editors/presentation-comments-panel';
import { PresentationDesignPanel } from '../src/internal/features/work/editors/presentation-design-panel';
import { PresentationEditor } from '../src/internal/features/work/editors/presentation-editor';
import type {
  PresentationEditorCanCommands,
  PresentationEditorCommands,
} from '../src/internal/features/work/editors/presentation-command-types';
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

  expect(screen.getByLabelText('演示图表数据')).toBeVisible();
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

  fireEvent.keyDown(screen.getByLabelText('演示图表数据'), {
    key: 'Escape',
  });
  expect(closes).toBe(1);
});

test('keeps the docked chart pane as a non-modal region', () => {
  const mediaQuery = installPresentationChartMatchMedia(false);
  const view = render(
    <PresentationChartPanel
      chart={chart()}
      onChange={() => undefined}
      onDelete={() => undefined}
      onClose={() => undefined}
    />,
  );
  try {
    const panel = screen.getByRole('region', { name: '演示图表数据' });
    expect(panel).not.toHaveAttribute('aria-modal');
  } finally {
    view.unmount();
    mediaQuery.restore();
  }
});

test('contains phone chart-pane focus and restores the selected chart', async () => {
  const mediaQuery = installPresentationChartMatchMedia(true);

  function Harness() {
    const [open, setOpen] = useState(false);
    const selectedChartRef = useRef<HTMLButtonElement>(null);
    return (
      <div>
        <button
          ref={selectedChartRef}
          type="button"
          onClick={() => setOpen(true)}
        >
          已选图表
        </button>
        {open && (
          <PresentationChartPanel
            chart={chart()}
            onChange={() => undefined}
            onDelete={() => undefined}
            restoreFocusTarget={() => selectedChartRef.current}
            onClose={() => setOpen(false)}
          />
        )}
      </div>
    );
  }

  const view = render(<Harness />);
  try {
    const selectedChart = screen.getByRole('button', { name: '已选图表' });
    fireEvent.click(selectedChart);

    const panel = screen.getByRole('dialog', { name: '演示图表数据' });
    const close = screen.getByRole('button', { name: '关闭演示图表数据' });
    const removeChart = screen.getByRole('button', { name: '删除演示图表' });
    const addSeries = screen.getByRole('button', { name: '添加图表系列' });
    expect(panel).toHaveAttribute('aria-modal', 'true');
    expect(selectedChart).toHaveAttribute('inert');
    expect(mediaQuery.queries).toContain('(max-width: 1100px)');
    await waitFor(() => expect(close).toHaveFocus());

    removeChart.focus();
    fireEvent.keyDown(removeChart, { key: 'Tab', shiftKey: true });
    expect(addSeries).toHaveFocus();
    fireEvent.keyDown(addSeries, { key: 'Tab' });
    expect(removeChart).toHaveFocus();

    fireEvent.keyDown(close, { key: 'Escape' });
    await waitFor(() => expect(panel).not.toBeInTheDocument());
    expect(selectedChart).not.toHaveAttribute('inert');
    await waitFor(() => expect(selectedChart).toHaveFocus());
  } finally {
    view.unmount();
    mediaQuery.restore();
  }
});

test('lets a chart field consume Escape before dismissing its task pane', async () => {
  render(<PresentationHarness />);

  fireEvent.click(screen.getByRole('tab', { name: '插入' }));
  fireEvent.click(screen.getByRole('button', { name: '图表' }));

  const chartPane = screen.getByLabelText('演示图表数据');
  const values = screen.getByRole('textbox', {
    name: '演示图表系列 1 数据',
  });
  fireEvent.change(values, { target: { value: '32, wrong, 61' } });
  expect(values).toHaveAttribute('aria-invalid', 'true');

  fireEvent.keyDown(values, { key: 'Escape' });
  await new Promise((resolve) => window.setTimeout(resolve, 0));
  expect(values).toHaveValue('32, 48, 61');
  expect(chartPane).toBeVisible();

  fireEvent.keyDown(values, { key: 'Escape' });
  expect(chartPane).not.toBeInTheDocument();
});

test('cancels chart title and series-name drafts before closing the task pane', () => {
  const changes: WorkSlideChart[] = [];
  let closes = 0;
  render(
    <PresentationChartPanel
      chart={chart()}
      onChange={(next) => changes.push(next)}
      onDelete={() => undefined}
      onClose={() => {
        closes += 1;
      }}
    />,
  );

  const title = screen.getByRole('textbox', { name: '演示图表标题' });
  fireEvent.change(title, { target: { value: 'Unsaved title' } });
  expect(changes).toEqual([]);
  fireEvent.keyDown(title, { key: 'Escape' });
  expect(title).toHaveValue('Quarterly revenue');
  expect(closes).toBe(0);

  const seriesName = screen.getByRole('textbox', {
    name: '演示图表系列 1 名称',
  });
  fireEvent.change(seriesName, { target: { value: 'Committed name' } });
  expect(changes).toEqual([]);
  fireEvent.blur(seriesName);
  expect(changes.at(-1)?.series[0]?.name).toBe('Committed name');
});

test('keeps design names as cancellable drafts and rejects blank names', () => {
  const calls: string[] = [];
  const commands = new Proxy(
    {},
    {
      get:
        (_target, property) =>
        (...args: unknown[]) =>
          calls.push(`${String(property)}:${args.join(',')}`),
    },
  ) as PresentationEditorCommands;
  const can = new Proxy(
    {},
    { get: () => () => true },
  ) as PresentationEditorCanCommands;
  const master = {
    id: 'master-1',
    name: 'Office master',
    background: '#ffffff',
    elements: [],
  };
  const layout = {
    id: 'layout-1',
    name: 'Title layout',
    masterId: master.id,
    elements: [],
  };
  const slide = {
    id: 'slide-1',
    name: 'Slide 1',
    background: '#ffffff',
    layoutId: layout.id,
    elements: [],
  };

  render(
    <PresentationDesignPanel
      can={can}
      commands={commands}
      content={{
        type: 'presentation',
        slides: [slide],
        masters: [master],
        layouts: [layout],
      }}
      slide={slide}
      layout={layout}
      master={master}
      mode="layout"
    />,
  );

  const name = screen.getByRole('textbox', { name: '布局名称' });
  fireEvent.change(name, { target: { value: 'Unsaved layout' } });
  expect(calls).toEqual([]);
  fireEvent.keyDown(name, { key: 'Escape' });
  expect(name).toHaveValue('Title layout');
  expect(calls).toEqual([]);

  fireEvent.change(name, { target: { value: '   ' } });
  expect(name).toHaveAttribute('aria-invalid', 'true');
  fireEvent.blur(name);
  expect(name).toHaveValue('Title layout');
  expect(calls).toEqual([]);

  fireEvent.change(name, { target: { value: '  Executive layout  ' } });
  fireEvent.blur(name);
  expect(calls).toEqual(['renamePresentationLayout:Executive layout']);
});

test('cancels comment drafts before closing and commits them on blur', () => {
  const calls: string[] = [];
  const slide = {
    id: 'slide-1',
    name: 'Slide 1',
    background: '#ffffff',
    elements: [],
    comments: [
      {
        id: 'comment-1',
        author: 'A3S',
        date: '2026-07-29T00:00:00.000Z',
        text: 'Original comment',
        x: 10,
        y: 10,
      },
    ],
  };
  const commands = new Proxy(
    {},
    {
      get:
        (_target, property) =>
        (...args: unknown[]) =>
          calls.push(`${String(property)}:${args.join(':')}`),
    },
  ) as PresentationEditorCommands;
  render(
    <PresentationCommentsPanel
      slides={[slide]}
      activeCommentId={null}
      commands={commands}
    />,
  );

  const comment = screen.getByRole('textbox', {
    name: '编辑演示批注 1',
  });
  fireEvent.change(comment, { target: { value: 'Unsaved comment' } });
  expect(calls).toEqual([]);
  fireEvent.keyDown(comment, { key: 'Escape' });
  expect(comment).toHaveValue('Original comment');
  expect(calls).toEqual([]);

  fireEvent.change(comment, { target: { value: 'Saved comment' } });
  fireEvent.blur(comment);
  expect(calls).toEqual([
    'updatePresentationComment:slide-1:comment-1:Saved comment',
  ]);
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

function installPresentationChartMatchMedia(initialMatches: boolean): {
  queries: string[];
  restore(): void;
} {
  const originalDescriptor = Object.getOwnPropertyDescriptor(
    window,
    'matchMedia',
  );
  const queries: string[] = [];
  Object.defineProperty(window, 'matchMedia', {
    configurable: true,
    value: (query: string) => {
      queries.push(query);
      return {
        matches: initialMatches,
        media: query,
        onchange: null,
        addEventListener: () => undefined,
        removeEventListener: () => undefined,
        addListener: () => undefined,
        removeListener: () => undefined,
        dispatchEvent: () => true,
      } as MediaQueryList;
    },
  });

  return {
    queries,
    restore() {
      if (originalDescriptor) {
        Object.defineProperty(window, 'matchMedia', originalDescriptor);
      } else {
        Reflect.deleteProperty(window, 'matchMedia');
      }
    },
  };
}
