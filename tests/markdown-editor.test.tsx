import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import type { MarkdownContent } from '../src/core';
import { proportionalMarkdownScrollTop } from '../src/internal/features/work/editors/markdown-workspace';
import { MarkdownEditor } from '../src/react';

test('renders the GFM compatibility surface in preview mode', async () => {
  const content: MarkdownContent = {
    type: 'markdown',
    markdown: [
      '# Release checklist',
      '',
      '- [ ] Review the API',
      '- [x] Publish the package',
      '',
      '~~Deprecated option~~',
      '',
      '| Capability | State |',
      '| --- | --- |',
      '| Tasks | Ready |',
      '',
      '<https://a3s-lab.github.io/Office/>',
    ].join('\n'),
  };

  render(
    <MarkdownEditor
      content={content}
      onChange={() => undefined}
      preview
      theme="light"
    />,
  );

  expect(await screen.findByText('Review the API')).toBeInTheDocument();
  expect(screen.getByText('Publish the package')).toBeInTheDocument();
  const tasks = screen.getAllByRole('checkbox');
  expect(tasks).toHaveLength(2);
  expect(tasks[0]).not.toBeChecked();
  expect(tasks[1]).toBeChecked();
  expect(screen.getByText('Deprecated option').closest('s')).not.toBeNull();
  expect(screen.getByRole('table')).toHaveTextContent('Tasks');
  expect(
    screen.getByRole('link', {
      name: 'https://a3s-lab.github.io/Office/',
    }),
  ).toHaveAttribute('href', 'https://a3s-lab.github.io/Office/');
});

test('coalesces source edits before rebuilding the visual Markdown tree', async () => {
  const changes: MarkdownContent[] = [];
  const content: MarkdownContent = {
    type: 'markdown',
    markdown: '# Initial title',
  };

  render(
    <MarkdownEditor
      content={content}
      onChange={(next) => changes.push(next)}
      theme="light"
    />,
  );

  const source = await screen.findByLabelText('Markdown 源码');
  const visual = screen.getByLabelText('Markdown 编辑区');
  fireEvent.change(source, { target: { value: '# Intermediate title' } });
  fireEvent.change(source, {
    target: { value: '# Final title\n\nCurrent content.' },
  });

  expect(changes.at(-1)?.markdown).toBe('# Final title\n\nCurrent content.');
  expect(visual).toHaveTextContent('Initial title');
  expect(visual).not.toHaveTextContent('Intermediate title');

  await waitFor(() => {
    expect(visual).toHaveTextContent('Final title');
    expect(visual).toHaveTextContent('Current content.');
  });
  expect(visual).not.toHaveTextContent('Intermediate title');
});

test('round-trips GFM task state from the visual editor to source', async () => {
  const changes: MarkdownContent[] = [];
  const content: MarkdownContent = {
    type: 'markdown',
    markdown: '- [ ] Ship the release',
  };

  render(
    <MarkdownEditor
      content={content}
      onChange={(next) => changes.push(next)}
      theme="light"
    />,
  );

  const taskText = await screen.findByText('Ship the release');
  const task = taskText.closest('li')?.querySelector('input[type="checkbox"]');
  expect(task).not.toBeNull();
  if (!task) throw new Error('Expected the rendered Markdown task checkbox.');
  fireEvent.click(task);

  await waitFor(() => {
    expect(changes.at(-1)?.markdown).toContain('- [x] Ship the release');
  });
  expect(
    (
      screen.getByLabelText('Markdown 源码') as HTMLTextAreaElement
    ).value.trim(),
  ).toBe('- [x] Ship the release');
});

test('applies a host-controlled Markdown replacement to both panes', async () => {
  const initial: MarkdownContent = {
    type: 'markdown',
    markdown: '# Initial title',
  };
  const { rerender } = render(
    <MarkdownEditor
      content={initial}
      onChange={() => undefined}
      theme="light"
    />,
  );

  expect(await screen.findByLabelText('Markdown 编辑区')).toHaveTextContent(
    'Initial title',
  );
  const replacement: MarkdownContent = {
    type: 'markdown',
    markdown: '# Host replacement\n\nExternal update.',
  };
  rerender(
    <MarkdownEditor
      content={replacement}
      onChange={() => undefined}
      theme="light"
    />,
  );

  await waitFor(() => {
    expect(screen.getByLabelText('Markdown 源码')).toHaveValue(
      replacement.markdown,
    );
    expect(screen.getByLabelText('Markdown 编辑区')).toHaveTextContent(
      'Host replacement',
    );
  });
});

test('maps split-pane scrolling by document progress', () => {
  expect(proportionalMarkdownScrollTop(450, 1000, 100, 2000, 200)).toBe(900);
  expect(proportionalMarkdownScrollTop(-20, 1000, 100, 2000, 200)).toBe(0);
  expect(proportionalMarkdownScrollTop(2000, 1000, 100, 2000, 200)).toBe(1800);
  expect(proportionalMarkdownScrollTop(20, 100, 100, 2000, 200)).toBe(0);
});
