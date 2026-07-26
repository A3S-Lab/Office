import { Extension } from '@tiptap/core';
import { expect, test } from '@rstest/core';
import {
  fireEvent,
  render,
  screen,
  waitFor,
  within,
} from '@testing-library/react';
import type { MarkdownContent } from '../src/core';
import { markdownTaskCheckboxLabel } from '../src/internal/features/work/editors/markdown-editor';
import { proportionalMarkdownScrollTop } from '../src/internal/features/work/editors/markdown-workspace';
import { MarkdownEditor } from '../src/react';

test('derives the Markdown task label from canonical node state', () => {
  expect(
    markdownTaskCheckboxLabel({
      attrs: { checked: true },
      textContent: 'Publish the package',
    }),
  ).toBe('已完成：Publish the package');
  expect(
    markdownTaskCheckboxLabel({
      attrs: { checked: false },
      textContent: 'Review the API',
    }),
  ).toBe('未完成：Review the API');
});

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

test('resizes and resets the Markdown split panes from the separator', async () => {
  render(
    <MarkdownEditor
      content={{ type: 'markdown', markdown: '# Split view' }}
      onChange={() => undefined}
      theme="light"
    />,
  );

  const separator = await screen.findByRole('separator', {
    name: '调整源码与编辑结果宽度',
  });
  const workspace = separator.closest('.work-markdown-workspace');
  expect(workspace).not.toBeNull();
  expect(separator).toHaveAttribute('aria-valuenow', '50');

  fireEvent.keyDown(separator, { key: 'ArrowRight' });
  expect(separator).toHaveAttribute('aria-valuenow', '55');
  expect(workspace).toHaveStyle({ '--work-markdown-source-pane': '55%' });

  fireEvent.keyDown(separator, { key: 'Home' });
  expect(separator).toHaveAttribute('aria-valuenow', '30');

  fireEvent.doubleClick(separator);
  expect(separator).toHaveAttribute('aria-valuenow', '50');
  expect(workspace).toHaveStyle({ '--work-markdown-source-pane': '50%' });
});

test('mounts host TipTap extensions in the Markdown editor', async () => {
  let shortcutCalls = 0;
  const hostShortcuts = Extension.create({
    name: 'testHostShortcuts',
    addKeyboardShortcuts() {
      return {
        F6: () => {
          shortcutCalls += 1;
          return true;
        },
      };
    },
  });

  render(
    <MarkdownEditor
      content={{ type: 'markdown', markdown: '# Extension test' }}
      extensions={[hostShortcuts]}
      onChange={() => undefined}
      theme="light"
    />,
  );

  fireEvent.keyDown(await screen.findByLabelText('Markdown 编辑区'), {
    code: 'F6',
    key: 'F6',
  });

  expect(shortcutCalls).toBe(1);
});

test('inserts Markdown links with display text from one dialog', async () => {
  const changes: MarkdownContent[] = [];
  render(
    <MarkdownEditor
      content={{ type: 'markdown', markdown: '' }}
      onChange={(content) => changes.push(content)}
      theme="light"
    />,
  );

  await screen.findByLabelText('Markdown 编辑区');
  fireEvent.click(screen.getByRole('tab', { name: '插入' }));
  fireEvent.click(screen.getByRole('button', { name: '添加链接' }));
  fireEvent.change(screen.getByRole('textbox', { name: '显示文字' }), {
    target: { value: 'A3S Office' },
  });
  fireEvent.change(screen.getByRole('textbox', { name: '链接地址' }), {
    target: { value: 'https://a3s.dev/office' },
  });
  fireEvent.click(
    within(screen.getByRole('dialog', { name: '添加链接' })).getByRole(
      'button',
      { name: '添加链接' },
    ),
  );

  await waitFor(() =>
    expect(changes.at(-1)?.markdown).toContain(
      '[A3S Office](https://a3s.dev/office)',
    ),
  );
});

test('inserts Markdown images with alternative text from one dialog', async () => {
  const changes: MarkdownContent[] = [];
  render(
    <MarkdownEditor
      content={{ type: 'markdown', markdown: '' }}
      onChange={(content) => changes.push(content)}
      theme="light"
    />,
  );

  await screen.findByLabelText('Markdown 编辑区');
  fireEvent.click(screen.getByRole('tab', { name: '插入' }));
  fireEvent.click(screen.getByRole('button', { name: '插入图片' }));
  fireEvent.change(screen.getByRole('textbox', { name: '替代文字（可选）' }), {
    target: { value: 'Office 架构图' },
  });
  fireEvent.change(screen.getByRole('textbox', { name: '图片地址' }), {
    target: { value: 'https://a3s.dev/office.png' },
  });
  fireEvent.click(
    within(screen.getByRole('dialog', { name: '插入图片' })).getByRole(
      'button',
      { name: '插入图片' },
    ),
  );

  await waitFor(() =>
    expect(changes.at(-1)?.markdown).toContain(
      '![Office 架构图](https://a3s.dev/office.png)',
    ),
  );
});
