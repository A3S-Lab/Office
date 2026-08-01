import { Editor } from '@tiptap/core';
import { afterEach, expect, test } from '@rstest/core';
import { fireEvent, render, screen, within } from '@testing-library/react';
import { MarkdownToolbar } from '../src/internal/features/work/editors/markdown-toolbar';
import type { MarkdownSourceCommand } from '../src/internal/features/work/editors/markdown-source-commands';
import { createWorkMarkdownExtensions } from '../src/internal/features/work/work-markdown-extensions';

let editor: Editor | null = null;

afterEach(() => {
  editor?.destroy();
  editor = null;
});

test('keeps failed source commands out of the hidden visual editor', () => {
  editor = new Editor({
    extensions: createWorkMarkdownExtensions(),
    content: '<p>Visual editor content</p>',
  });
  const sourceCommands: MarkdownSourceCommand[] = [];

  render(
    <MarkdownToolbar
      editor={editor}
      sourceEditing
      canSourceRedo={false}
      canSourceUndo={false}
      viewMode="split"
      getSourceFocusTarget={() => null}
      getSourceSelection={() => ({
        markdown: 'Source editor content',
        selection: { start: 0, end: 6, direction: 'none' },
        text: 'Source',
      })}
      onSourceCommand={(command) => {
        sourceCommands.push(command);
        return false;
      }}
      onSourceRedo={() => false}
      onSourceReplace={() => false}
      onSourceUndo={() => false}
      onViewModeChange={() => undefined}
    />,
  );

  fireEvent.click(screen.getByRole('button', { name: '加粗' }));

  expect(sourceCommands).toEqual(['bold']);
  expect(editor.isActive('bold')).toBe(false);
  expect(editor.getHTML()).toBe('<p>Visual editor content</p>');
});

test('does not leak hidden visual heading or link state into source menus', () => {
  editor = new Editor({
    extensions: createWorkMarkdownExtensions(),
    content: '<h1><a href="https://a3s.dev">Visual link</a></h1>',
  });
  editor.commands.setTextSelection(2);

  render(
    <MarkdownToolbar
      editor={editor}
      sourceEditing
      canSourceRedo={false}
      canSourceUndo={false}
      viewMode="source"
      getSourceFocusTarget={() => null}
      getSourceSelection={() => ({
        markdown: 'Plain source paragraph',
        selection: { start: 4, end: 4, direction: 'none' },
        text: '',
      })}
      onSourceCommand={() => true}
      onSourceRedo={() => false}
      onSourceReplace={() => false}
      onSourceUndo={() => false}
      onViewModeChange={() => undefined}
    />,
  );

  expect(screen.getByRole('combobox', { name: '段落样式' })).toHaveTextContent(
    '正文',
  );
  fireEvent.click(screen.getByRole('tab', { name: '插入' }));
  expect(screen.getByRole('button', { name: '添加链接' })).toHaveAttribute(
    'aria-pressed',
    'false',
  );
  expect(screen.getByRole('button', { name: '移除链接' })).toBeDisabled();
});

test('edits and removes the exact source link targeted by the current selection', () => {
  editor = new Editor({
    extensions: createWorkMarkdownExtensions(),
    content: '<p>Hidden visual content</p>',
  });
  const markdown = '[Office](https://a3s.dev) after';
  const replacements: Array<{
    replacement: string;
    selectedRange?: { start: number; end: number };
    target?: {
      markdown: string;
      selection: { start: number; end: number; direction: string };
      text: string;
    };
  }> = [];

  render(
    <MarkdownToolbar
      editor={editor}
      sourceEditing
      canSourceRedo={false}
      canSourceUndo={false}
      viewMode="source"
      getSourceFocusTarget={() => null}
      getSourceSelection={() => ({
        markdown,
        selection: { start: 3, end: 3, direction: 'none' },
        text: '',
      })}
      onSourceCommand={() => true}
      onSourceRedo={() => false}
      onSourceReplace={(replacement, selectedRange, target) => {
        replacements.push({ replacement, selectedRange, target });
        return true;
      }}
      onSourceUndo={() => false}
      onViewModeChange={() => undefined}
    />,
  );

  fireEvent.click(screen.getByRole('tab', { name: '插入' }));
  expect(screen.getByRole('button', { name: '编辑链接' })).toHaveAttribute(
    'aria-pressed',
    'true',
  );
  expect(screen.getByRole('button', { name: '移除链接' })).toBeEnabled();

  fireEvent.click(screen.getByRole('button', { name: '编辑链接' }));
  expect(screen.getByRole('textbox', { name: '显示文字' })).toHaveValue(
    'Office',
  );
  expect(screen.getByRole('textbox', { name: '链接地址' })).toHaveValue(
    'https://a3s.dev',
  );
  fireEvent.change(screen.getByRole('textbox', { name: '显示文字' }), {
    target: { value: 'Office docs' },
  });
  fireEvent.change(screen.getByRole('textbox', { name: '链接地址' }), {
    target: { value: 'https://a3s.dev/office' },
  });
  fireEvent.click(screen.getByRole('button', { name: '保存' }));

  expect(replacements).toEqual([
    {
      replacement: '[Office docs](https://a3s.dev/office)',
      selectedRange: { start: 1, end: 12 },
      target: {
        markdown,
        selection: { start: 0, end: 25, direction: 'none' },
        text: '[Office](https://a3s.dev)',
      },
    },
  ]);

  fireEvent.click(screen.getByRole('button', { name: '移除链接' }));
  expect(replacements.at(-1)).toEqual({
    replacement: 'Office',
    selectedRange: { start: 0, end: 6 },
    target: {
      markdown,
      selection: { start: 0, end: 25, direction: 'none' },
      text: '[Office](https://a3s.dev)',
    },
  });
});

test('advertises only the keyboard shortcuts implemented by the editor', () => {
  editor = new Editor({
    extensions: createWorkMarkdownExtensions(),
    content: '<p>Markdown shortcuts</p>',
  });

  render(
    <MarkdownToolbar
      editor={editor}
      sourceEditing={false}
      canSourceRedo={false}
      canSourceUndo={false}
      viewMode="visual"
      getSourceFocusTarget={() => null}
      getSourceSelection={() => null}
      onSourceCommand={() => false}
      onSourceRedo={() => false}
      onSourceReplace={() => false}
      onSourceUndo={() => false}
      onViewModeChange={() => undefined}
    />,
  );

  expect(screen.getByRole('button', { name: '撤销' })).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Z Meta+Z',
  );
  expect(screen.getByRole('button', { name: '重做' })).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+Z Meta+Shift+Z Control+Y Meta+Y',
  );
  expect(screen.getByRole('button', { name: '加粗' })).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+B Meta+B',
  );
  expect(screen.getByRole('button', { name: '斜体' })).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+I Meta+I',
  );
});

test('exposes concise and unambiguous Markdown view modes', () => {
  editor = new Editor({
    extensions: createWorkMarkdownExtensions(),
    content: '<p>Markdown view modes</p>',
  });

  render(
    <MarkdownToolbar
      editor={editor}
      sourceEditing
      canSourceRedo={false}
      canSourceUndo={false}
      viewMode="split"
      getSourceFocusTarget={() => null}
      getSourceSelection={() => null}
      onSourceCommand={() => false}
      onSourceRedo={() => false}
      onSourceReplace={() => false}
      onSourceUndo={() => false}
      onViewModeChange={() => undefined}
    />,
  );

  fireEvent.click(screen.getByRole('tab', { name: '视图' }));
  const modeGroup = screen.getByRole('region', { name: '编辑方式' });
  const visual = within(modeGroup).getByRole('button', {
    name: '可视化编辑',
  });
  const source = within(modeGroup).getByRole('button', { name: '源码编辑' });
  const split = within(modeGroup).getByRole('button', { name: '分屏编辑' });

  expect(visual).toHaveTextContent('可视化');
  expect(source).toHaveTextContent('源码');
  expect(split).toHaveTextContent('分屏');
  expect(visual).toHaveAttribute('aria-pressed', 'false');
  expect(source).toHaveAttribute('aria-pressed', 'false');
  expect(split).toHaveAttribute('aria-pressed', 'true');
  expect(within(modeGroup).queryByRole('button', { name: '编辑' })).toBeNull();
});
