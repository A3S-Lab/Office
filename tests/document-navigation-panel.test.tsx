import { Editor } from '@tiptap/core';
import { expect, test } from '@rstest/core';
import {
  act,
  fireEvent,
  render,
  screen,
  waitFor,
  within,
} from '@testing-library/react';
import { DocumentNavigationPanel } from '../src/internal/features/work/editors/document-navigation-panel';
import {
  collectWorkDocumentOutline,
  currentWorkDocumentOutlineItem,
  visibleWorkDocumentOutlineItems,
} from '../src/internal/features/work/work-document-outline';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';

const documentHtml = [
  '<h1>项目方案</h1>',
  '<p>项目简介。</p>',
  '<h2>背景与目标</h2>',
  '<p>背景内容。</p>',
  '<h3>成功标准</h3>',
  '<p>成功标准内容。</p>',
  '<h2>工作范围</h2>',
  '<p>工作范围内容。</p>',
  '<h1>交付计划</h1>',
  '<p>交付计划内容。</p>',
].join('');

test('derives a typed hierarchy and the heading for the current position', () => {
  const { editor, element } = createEditor();

  try {
    const outline = collectWorkDocumentOutline(editor.state.doc);
    expect(
      outline.map(({ text, level, depth, hasChildren, parentId }) => ({
        text,
        level,
        depth,
        hasChildren,
        parent: outline.find((item) => item.id === parentId)?.text,
      })),
    ).toEqual([
      {
        text: '项目方案',
        level: 1,
        depth: 0,
        hasChildren: true,
        parent: undefined,
      },
      {
        text: '背景与目标',
        level: 2,
        depth: 1,
        hasChildren: true,
        parent: '项目方案',
      },
      {
        text: '成功标准',
        level: 3,
        depth: 2,
        hasChildren: false,
        parent: '背景与目标',
      },
      {
        text: '工作范围',
        level: 2,
        depth: 1,
        hasChildren: false,
        parent: '项目方案',
      },
      {
        text: '交付计划',
        level: 1,
        depth: 0,
        hasChildren: false,
        parent: undefined,
      },
    ]);

    const success = outline.find((item) => item.text === '成功标准');
    if (!success) throw new Error('Expected the success heading.');
    expect(
      currentWorkDocumentOutlineItem(outline, success.from + 2)?.text,
    ).toBe('成功标准');
    expect(
      visibleWorkDocumentOutlineItems(
        outline,
        new Set([outline[0]?.id ?? '']),
        '',
      ).map((item) => item.text),
    ).toEqual(['项目方案', '交付计划']);
    expect(
      visibleWorkDocumentOutlineItems(outline, new Set(), '目标').map(
        (item) => item.text,
      ),
    ).toEqual(['背景与目标']);
  } finally {
    editor.destroy();
    element.remove();
  }
});

test('filters, collapses, keyboard-navigates, and jumps through headings', async () => {
  const { editor, element } = createEditor();

  try {
    render(
      <DocumentNavigationPanel editor={editor} onClose={() => undefined} />,
    );
    const pane = screen.getByRole('complementary', { name: '文档导航' });
    const search = within(pane).getByRole('searchbox', { name: '搜索标题' });
    expect(search).toHaveFocus();
    expect(within(pane).getByText('5 个标题')).toBeVisible();

    fireEvent.click(
      within(pane).getByRole('button', { name: '折叠 项目方案' }),
    );
    expect(
      within(pane).queryByRole('button', { name: '背景与目标' }),
    ).not.toBeInTheDocument();

    const project = within(pane).getByRole('button', { name: '项目方案' });
    project.focus();
    fireEvent.keyDown(project, { key: 'ArrowRight' });
    const background = within(pane).getByRole('button', {
      name: '背景与目标',
    });
    expect(background).toBeVisible();
    fireEvent.keyDown(project, { key: 'ArrowDown' });
    await waitFor(() => expect(background).toHaveFocus());

    fireEvent.change(search, { target: { value: '交付' } });
    expect(within(pane).getByText('1 个匹配')).toBeVisible();
    expect(
      within(pane).queryByRole('button', { name: '项目方案' }),
    ).not.toBeInTheDocument();
    const delivery = within(pane).getByRole('button', { name: '交付计划' });
    fireEvent.click(delivery);
    expect(
      editor.state.doc.textBetween(
        editor.state.selection.from,
        editor.state.selection.from + '交付计划'.length,
      ),
    ).toBe('交付计划');
    expect(delivery).toHaveAttribute('aria-current', 'location');
  } finally {
    editor.destroy();
    element.remove();
  }
});

test('refreshes the outline after a controlled document replacement', () => {
  const { editor, element } = createEditor();

  try {
    render(
      <DocumentNavigationPanel editor={editor} onClose={() => undefined} />,
    );
    expect(screen.getByRole('button', { name: '项目方案' })).toBeVisible();
    act(() => {
      editor.commands.setContent(
        '<h1>替换后的文档</h1><h2>新的章节</h2><p>正文。</p>',
        { emitUpdate: false },
      );
    });
    expect(
      screen.queryByRole('button', { name: '项目方案' }),
    ).not.toBeInTheDocument();
    expect(screen.getByRole('button', { name: '替换后的文档' })).toBeVisible();
    expect(screen.getByRole('button', { name: '新的章节' })).toBeVisible();
  } finally {
    editor.destroy();
    element.remove();
  }
});

function createEditor(): {
  editor: Editor;
  element: HTMLDivElement;
} {
  const element = document.createElement('div');
  document.body.append(element);
  return {
    editor: new Editor({
      element,
      extensions: createWorkDocumentExtensions(),
      content: documentHtml,
    }),
    element,
  };
}
