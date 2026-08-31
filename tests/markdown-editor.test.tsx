import { type Editor, Extension } from '@tiptap/core';
import { expect, test } from '@rstest/core';
import {
  act,
  fireEvent,
  render,
  screen,
  waitFor,
  within,
} from '@testing-library/react';
import { useState } from 'react';
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

test('guides an empty Markdown source pane without adding document content', async () => {
  render(
    <MarkdownEditor
      content={{ type: 'markdown', markdown: '' }}
      onChange={() => undefined}
      theme="light"
    />,
  );

  expect(await screen.findByLabelText('Markdown 源码')).toHaveAttribute(
    'placeholder',
    '开始写 Markdown…',
  );
});

test('keeps the default split view to one editor and one read-only preview', async () => {
  render(
    <MarkdownEditor
      content={{
        type: 'markdown',
        markdown: '# Product plan\n\n- [ ] Review the plan',
      }}
      onChange={() => undefined}
      theme="light"
    />,
  );

  const source = await screen.findByLabelText('Markdown 源码');
  const preview = screen.getByLabelText('Markdown 预览');
  expect(source).not.toHaveAttribute('readonly');
  expect(preview).toHaveAttribute('contenteditable', 'false');
  expect(preview).toHaveAttribute('aria-readonly', 'true');
  expect(preview).toHaveAttribute('role', 'document');
  expect(
    screen.getByRole('checkbox', { name: '未完成：Review the plan' }),
  ).toBeDisabled();
  expect(
    within(screen.getByRole('region', { name: 'Markdown 源码窗格' })).getByText(
      '源码',
    ),
  ).toBeInTheDocument();
  expect(
    within(screen.getByRole('region', { name: 'Markdown 预览窗格' })).getByText(
      '预览',
    ),
  ).toBeInTheDocument();
  const splitWorkspace = source.closest('.work-markdown-workspace');
  const compactSwitch = screen.getByRole('group', { name: '分屏显示内容' });
  const showSource = within(compactSwitch).getByRole('button', {
    name: '显示源码窗格',
  });
  const showPreview = within(compactSwitch).getByRole('button', {
    name: '显示预览窗格',
  });
  expect(splitWorkspace).toHaveAttribute('data-compact-pane', 'source');
  expect(showSource).toHaveAttribute('aria-pressed', 'true');
  expect(showPreview).toHaveAttribute('aria-pressed', 'false');

  fireEvent.click(showPreview);
  expect(splitWorkspace).toHaveAttribute('data-compact-pane', 'preview');
  expect(showSource).toHaveAttribute('aria-pressed', 'false');
  expect(showPreview).toHaveAttribute('aria-pressed', 'true');

  fireEvent.click(await screen.findByRole('tab', { name: '视图' }));
  fireEvent.click(
    within(screen.getByRole('region', { name: '编辑方式' })).getByRole(
      'button',
      { name: '可视化编辑' },
    ),
  );

  const visualEditor = await screen.findByLabelText('Markdown 编辑区');
  expect(visualEditor).toHaveAttribute('contenteditable', 'true');
  expect(visualEditor).toHaveAttribute('aria-readonly', 'false');
  expect(visualEditor).toHaveAttribute('role', 'textbox');
  const editableTask = visualEditor.querySelector<HTMLInputElement>(
    'li[data-type="taskItem"] input[type="checkbox"]',
  );
  expect(editableTask).not.toBeNull();
  expect(editableTask).toBeEnabled();
  expect(screen.queryByLabelText('Markdown 源码')).toBeNull();
});

test('publishes only committed Chinese text from controlled visual Markdown IME', async () => {
  let editor: Editor | null = null;
  const publications: MarkdownContent[] = [];
  const captureEditor = Extension.create({
    name: 'captureControlledMarkdownCompositionEditor',
    onCreate() {
      editor = this.editor;
    },
  });

  function ControlledMarkdownEditor() {
    const [content, setContent] = useState<MarkdownContent>({
      type: 'markdown',
      markdown: '',
    });
    return (
      <MarkdownEditor
        content={content}
        extensions={[captureEditor]}
        onChange={(next) => {
          publications.push(next);
          setContent(next);
        }}
        theme="light"
      />
    );
  }

  render(<ControlledMarkdownEditor />);
  await screen.findByLabelText('Markdown 源码');
  fireEvent.click(screen.getByRole('tab', { name: '视图' }));
  fireEvent.click(
    within(screen.getByRole('region', { name: '编辑方式' })).getByRole(
      'button',
      { name: '可视化编辑' },
    ),
  );
  const surface = await screen.findByLabelText('Markdown 编辑区');
  await waitFor(() => expect(editor).not.toBeNull());
  const current = editor as Editor;

  fireEvent.compositionStart(surface, { data: 'qingwen' });
  expect(current.view.composing).toBe(true);
  act(() => {
    current.commands.insertContent('qingwen');
  });
  expect(publications).toEqual([]);

  act(() => {
    current.chain().selectAll().insertContent('请问').run();
  });
  expect(publications).toEqual([]);

  fireEvent.compositionEnd(surface, { data: '请问' });
  await waitFor(() =>
    expect(publications).toEqual([{ type: 'markdown', markdown: '请问' }]),
  );
  expect(surface).toHaveTextContent('请问');
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
  const visual = screen.getByLabelText('Markdown 预览');
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

test('undoes and redoes coalesced typing from the Markdown source surface', async () => {
  const changes: MarkdownContent[] = [];
  render(
    <MarkdownEditor
      content={{ type: 'markdown', markdown: '# Initial' }}
      onChange={(content) => changes.push(content)}
      theme="light"
    />,
  );

  const source = (await screen.findByLabelText(
    'Markdown 源码',
  )) as HTMLTextAreaElement;
  fireEvent.focus(source);
  source.setSelectionRange(source.value.length, source.value.length);
  fireEvent.select(source);
  fireEvent.change(source, {
    target: {
      value: '# Draft',
      selectionStart: 7,
      selectionEnd: 7,
      selectionDirection: 'none',
    },
  });
  fireEvent.change(source, {
    target: {
      value: '# Draft title',
      selectionStart: 13,
      selectionEnd: 13,
      selectionDirection: 'none',
    },
  });

  const undo = screen.getByRole('button', { name: '撤销' });
  const redo = screen.getByRole('button', { name: '重做' });
  await waitFor(() => expect(undo).toBeEnabled());
  fireEvent.click(undo);
  await waitFor(() => expect(source).toHaveValue('# Initial'));
  expect(source.selectionStart).toBe('# Initial'.length);
  expect(redo).toBeEnabled();

  fireEvent.click(redo);
  await waitFor(() => expect(source).toHaveValue('# Draft title'));
  expect(source.selectionStart).toBe('# Draft title'.length);
  expect(changes.at(-1)?.markdown).toBe('# Draft title');
});

test('routes source undo and redo shortcuts through the controlled history', async () => {
  render(
    <MarkdownEditor
      content={{ type: 'markdown', markdown: 'Write clearly' }}
      onChange={() => undefined}
      theme="light"
    />,
  );

  const source = (await screen.findByLabelText(
    'Markdown 源码',
  )) as HTMLTextAreaElement;
  fireEvent.focus(source);
  source.setSelectionRange(0, 5);
  fireEvent.select(source);
  fireEvent.click(screen.getByRole('button', { name: '加粗' }));
  await waitFor(() => expect(source).toHaveValue('**Write** clearly'));

  fireEvent.keyDown(source, { key: 'z', ctrlKey: true });
  await waitFor(() => expect(source).toHaveValue('Write clearly'));
  expect(source.selectionStart).toBe(0);
  expect(source.selectionEnd).toBe(5);

  fireEvent.keyDown(source, { key: 'z', ctrlKey: true, shiftKey: true });
  await waitFor(() => expect(source).toHaveValue('**Write** clearly'));
  expect(source.selectionStart).toBe(2);
  expect(source.selectionEnd).toBe(7);
});

test('rebases source history after a host-controlled Markdown replacement', async () => {
  const initial: MarkdownContent = {
    type: 'markdown',
    markdown: '# Initial',
  };
  const { rerender } = render(
    <MarkdownEditor
      content={initial}
      onChange={() => undefined}
      theme="light"
    />,
  );

  const source = (await screen.findByLabelText(
    'Markdown 源码',
  )) as HTMLTextAreaElement;
  fireEvent.focus(source);
  fireEvent.change(source, { target: { value: '# Local edit' } });
  await waitFor(() =>
    expect(screen.getByRole('button', { name: '撤销' })).toBeEnabled(),
  );

  rerender(
    <MarkdownEditor
      content={{ type: 'markdown', markdown: '# Remote edit' }}
      onChange={() => undefined}
      theme="light"
    />,
  );

  await waitFor(() => expect(source).toHaveValue('# Remote edit'));
  expect(screen.getByRole('button', { name: '撤销' })).toBeDisabled();
  fireEvent.keyDown(source, { key: 'z', ctrlKey: true });
  expect(source).toHaveValue('# Remote edit');
});

test('routes ribbon formatting to the active Markdown source selection', async () => {
  const changes: MarkdownContent[] = [];
  render(
    <MarkdownEditor
      content={{ type: 'markdown', markdown: 'Write clearly' }}
      onChange={(content) => changes.push(content)}
      theme="light"
    />,
  );

  const source = (await screen.findByLabelText(
    'Markdown 源码',
  )) as HTMLTextAreaElement;
  fireEvent.focus(source);
  source.setSelectionRange(0, 5);
  fireEvent.select(source);
  fireEvent.click(screen.getByRole('button', { name: '加粗' }));

  await waitFor(() => expect(source).toHaveValue('**Write** clearly'));
  await waitFor(() =>
    expect(screen.getByRole('button', { name: '加粗' })).toHaveAttribute(
      'aria-pressed',
      'true',
    ),
  );
  expect(changes.at(-1)?.markdown).toBe('**Write** clearly');
  expect(source.selectionStart).toBe(2);
  expect(source.selectionEnd).toBe(7);

  fireEvent.keyDown(source, { key: 'i', ctrlKey: true });
  await waitFor(() => expect(source).toHaveValue('***Write*** clearly'));
});

test('applies block formatting to selected Markdown source lines', async () => {
  render(
    <MarkdownEditor
      content={{ type: 'markdown', markdown: 'First item\nSecond item' }}
      onChange={() => undefined}
      theme="light"
    />,
  );

  const source = (await screen.findByLabelText(
    'Markdown 源码',
  )) as HTMLTextAreaElement;
  fireEvent.focus(source);
  source.setSelectionRange(0, source.value.length);
  fireEvent.select(source);
  fireEvent.click(screen.getByRole('button', { name: '项目列表' }));

  await waitFor(() =>
    expect(source).toHaveValue('- First item\n- Second item'),
  );
});

test('lets the host replace the selected-text menu in Markdown source', async () => {
  const content: MarkdownContent = {
    type: 'markdown',
    markdown: '# Plan\n\nShip the release today.\n\n## Notes',
  };
  const snapshots: import('../src/core').MarkdownSelectionSnapshot[] = [];
  const actions: import('../src/core').MarkdownSelectionContext[] = [];
  render(
    <MarkdownEditor
      content={content}
      getSelectionMenuItems={(snapshot) => {
        snapshots.push(snapshot);
        return [
          {
            id: 'polish',
            label: '润色',
            icon: 'wand',
            onSelect: (context) => actions.push(context),
          },
        ];
      }}
      onChange={() => undefined}
      theme="light"
    />,
  );

  const source = (await screen.findByLabelText(
    'Markdown 源码',
  )) as HTMLTextAreaElement;
  const start = content.markdown.indexOf('Ship the release today.');
  source.setSelectionRange(start, start + 'Ship the release today.'.length);
  fireEvent.select(source);
  fireEvent.contextMenu(source, { clientX: 120, clientY: 180 });

  const menu = await screen.findByRole('menu', { name: '选中文本操作' });
  fireEvent.click(within(menu).getByRole('menuitem', { name: '润色' }));

  expect(snapshots).toHaveLength(1);
  expect(snapshots[0]?.selection.surface).toBe('source');
  expect(snapshots[0]?.selection.text).toBe('Ship the release today.');
  expect(snapshots[0]?.selection.beforeText).toContain('# Plan');
  expect(snapshots[0]?.selection.afterText).toContain('## Notes');
  expect(snapshots[0]?.document.markdown).toBe(content.markdown);
  expect(actions).toHaveLength(1);

  expect(actions[0]?.commands.replaceText('Ship today.')).toEqual({
    applied: true,
  });
  await waitFor(() =>
    expect(source).toHaveValue('# Plan\n\nShip today.\n\n## Notes'),
  );
});

test('keeps the host selection menu available in the Markdown preview', async () => {
  const changes: MarkdownContent[] = [];
  const surfaces: string[] = [];
  const commandResults: import('../src/core').MarkdownSelectionCommandResult[] =
    [];
  render(
    <MarkdownEditor
      content={{ type: 'markdown', markdown: 'Review this sentence.' }}
      getSelectionMenuItems={(snapshot) => {
        surfaces.push(snapshot.selection.surface);
        return [
          {
            id: 'replace',
            label: '替换',
            onSelect: ({ commands }) => {
              commandResults.push(commands.replaceText('Keep this sentence.'));
            },
          },
        ];
      }}
      onChange={(content) => changes.push(content)}
      theme="light"
    />,
  );

  const visual = await screen.findByLabelText('Markdown 预览');
  selectDomText(visual, 'Review this sentence.');
  await new Promise((resolve) => setTimeout(resolve, 0));
  fireEvent.contextMenu(visual, { clientX: 120, clientY: 180 });
  fireEvent.click(
    within(await screen.findByRole('menu', { name: '选中文本操作' })).getByRole(
      'menuitem',
      { name: '替换' },
    ),
  );

  expect(surfaces).toEqual(['visual']);
  expect(commandResults).toEqual([{ applied: false, reason: 'read-only' }]);
  expect(changes).toEqual([]);
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

  fireEvent.click(await screen.findByRole('tab', { name: '视图' }));
  fireEvent.click(
    within(screen.getByRole('region', { name: '编辑方式' })).getByRole(
      'button',
      { name: '可视化编辑' },
    ),
  );
  const visualEditor = await screen.findByLabelText('Markdown 编辑区');
  await waitFor(() => expect(visualEditor).toHaveFocus());

  const taskText = await screen.findByText('Ship the release');
  const task = taskText.closest('li')?.querySelector('input[type="checkbox"]');
  expect(task).not.toBeNull();
  if (!task) throw new Error('Expected the rendered Markdown task checkbox.');
  expect(task).toBeEnabled();
  expect(task).not.toBeChecked();
  fireEvent.change(task, { target: { checked: true } });

  await waitFor(() => {
    expect(changes.at(-1)?.markdown).toContain('- [x] Ship the release');
  });
  fireEvent.click(
    within(screen.getByRole('region', { name: '编辑方式' })).getByRole(
      'button',
      { name: '分屏编辑' },
    ),
  );
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

  expect(await screen.findByLabelText('Markdown 预览')).toHaveTextContent(
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
    expect(screen.getByLabelText('Markdown 预览')).toHaveTextContent(
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
    name: '调整编辑与预览宽度',
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

  fireEvent.click(screen.getByRole('tab', { name: '视图' }));
  fireEvent.click(
    within(screen.getByRole('region', { name: '编辑方式' })).getByRole(
      'button',
      { name: '可视化编辑' },
    ),
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

  await screen.findByLabelText('Markdown 源码');
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
      { name: '添加' },
    ),
  );

  await waitFor(() =>
    expect(changes.at(-1)?.markdown).toContain(
      '[A3S Office](https://a3s.dev/office)',
    ),
  );
});

test('rejects unsafe Markdown link and image sources with one concise error', async () => {
  const changes: MarkdownContent[] = [];
  render(
    <MarkdownEditor
      content={{ type: 'markdown', markdown: '' }}
      onChange={(content) => changes.push(content)}
      theme="light"
    />,
  );

  await screen.findByLabelText('Markdown 源码');
  fireEvent.click(screen.getByRole('tab', { name: '插入' }));
  fireEvent.click(screen.getByRole('button', { name: '添加链接' }));
  const linkDialog = screen.getByRole('dialog', { name: '添加链接' });
  const linkSource = within(linkDialog).getByRole('textbox', {
    name: '链接地址',
  });
  fireEvent.change(linkSource, {
    target: { value: 'javascript:alert(1)' },
  });

  expect(linkSource).toHaveAttribute('aria-invalid', 'true');
  expect(within(linkDialog).getByRole('alert')).toHaveTextContent(
    '请输入完整的 http、https、mailto 或 # 文档内地址。',
  );
  expect(
    within(linkDialog).getByRole('button', { name: '添加' }),
  ).toBeDisabled();
  expect(changes).toHaveLength(0);

  fireEvent.click(within(linkDialog).getByRole('button', { name: '取消' }));
  fireEvent.click(screen.getByRole('button', { name: '插入图片' }));
  const imageDialog = screen.getByRole('dialog', { name: '插入图片' });
  const imageSource = within(imageDialog).getByRole('textbox', {
    name: '图片地址',
  });
  fireEvent.change(imageSource, {
    target: { value: 'data:text/html,<script>alert(1)</script>' },
  });

  expect(imageSource).toHaveAttribute('aria-invalid', 'true');
  expect(within(imageDialog).getByRole('alert')).toHaveTextContent(
    '请输入完整的 http、https 或相对图片地址。',
  );
  expect(
    within(imageDialog).getByRole('button', { name: '插入' }),
  ).toBeDisabled();
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

  await screen.findByLabelText('Markdown 源码');
  fireEvent.click(screen.getByRole('tab', { name: '插入' }));
  fireEvent.click(screen.getByRole('button', { name: '插入图片' }));
  fireEvent.change(screen.getByRole('textbox', { name: '替代文字（可选）' }), {
    target: { value: 'Office 架构图' },
  });
  fireEvent.change(screen.getByRole('textbox', { name: '图片地址' }), {
    target: { value: '../assets/office diagram.png' },
  });
  fireEvent.click(
    within(screen.getByRole('dialog', { name: '插入图片' })).getByRole(
      'button',
      { name: '插入' },
    ),
  );

  await waitFor(() =>
    expect(changes.at(-1)?.markdown).toContain(
      '![Office 架构图](../assets/office%20diagram.png)',
    ),
  );
});

function selectDomText(root: HTMLElement, text: string): void {
  const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
  let node = walker.nextNode();
  while (node) {
    const value = node.textContent ?? '';
    const offset = value.indexOf(text);
    if (offset >= 0) {
      const range = document.createRange();
      range.setStart(node, offset);
      range.setEnd(node, offset + text.length);
      const selection = window.getSelection();
      selection?.removeAllRanges();
      selection?.addRange(range);
      root.focus();
      document.dispatchEvent(new Event('selectionchange', { bubbles: true }));
      return;
    }
    node = walker.nextNode();
  }
  throw new Error(`Unable to select "${text}" in the editor.`);
}
