import { Editor } from '@tiptap/core';
import { expect, test } from '@rstest/core';
import { useState } from 'react';
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

test('collapses, keyboard-navigates, and jumps through headings', async () => {
  const { editor, element } = createEditor();

  try {
    render(
      <DocumentNavigationPanel editor={editor} onClose={() => undefined} />,
    );
    const pane = screen.getByRole('complementary', { name: '文档导航' });
    const search = within(pane).getByRole('searchbox', { name: '搜索文档' });
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

test('searches body text, previews context, and jumps to a selected result', async () => {
  const { editor, element } = createEditor();

  try {
    render(
      <DocumentNavigationPanel editor={editor} onClose={() => undefined} />,
    );
    const pane = screen.getByRole('complementary', { name: '文档导航' });
    const search = within(pane).getByRole('searchbox', { name: '搜索文档' });

    fireEvent.change(search, { target: { value: '成功标准内容' } });
    expect(within(pane).getByText('1 个匹配')).toBeVisible();
    expect(
      within(pane).getByRole('navigation', { name: '文档搜索结果' }),
    ).toBeVisible();
    const result = within(pane).getByRole('button', {
      name: '第 1 个匹配：成功标准内容',
    });
    expect(result).toHaveTextContent('成功标准内容');
    expect(result).toHaveTextContent('成功标准');

    fireEvent.click(result);
    expect(
      editor.state.doc.textBetween(
        editor.state.selection.from,
        editor.state.selection.to,
      ),
    ).toBe('成功标准内容');
    expect(result).toHaveAttribute('aria-current', 'location');

    fireEvent.change(search, { target: { value: '内容' } });
    expect(within(pane).getByText('4 个匹配')).toBeVisible();
    fireEvent.keyDown(search, { key: 'ArrowDown' });
    const first = within(pane).getByRole('button', {
      name: '第 1 个匹配：内容',
    });
    const second = within(pane).getByRole('button', {
      name: '第 2 个匹配：内容',
    });
    await waitFor(() => expect(first).toHaveFocus());
    fireEvent.keyDown(first, { key: 'ArrowDown' });
    await waitFor(() => expect(second).toHaveFocus());

    fireEvent.change(search, { target: { value: '' } });
    expect(
      within(pane).getByRole('navigation', { name: '文档标题' }),
    ).toBeVisible();
    expect(within(pane).getByText('5 个标题')).toBeVisible();
  } finally {
    editor.destroy();
    element.remove();
  }
});

test('switches to page previews and jumps to the selected page', async () => {
  const { editor, element } = createEditor();
  const outline = collectWorkDocumentOutline(editor.state.doc);
  const delivery = outline.find((item) => item.text === '交付计划');
  if (!delivery) throw new Error('Expected the delivery heading.');

  try {
    render(
      <DocumentNavigationPanel
        currentPage={1}
        editor={editor}
        pages={[
          {
            physicalPage: 1,
            pageNumber: 1,
            orientation: 'portrait',
            previewText: '项目方案 背景与目标 工作范围',
            selectionPosition: 1,
          },
          {
            physicalPage: 2,
            pageNumber: 2,
            orientation: 'portrait',
            previewText: '交付计划 交付计划内容',
            selectionPosition: delivery.from,
          },
        ]}
        onClose={() => undefined}
      />,
    );

    fireEvent.click(screen.getByRole('tab', { name: '页面' }));
    const pages = screen.getByRole('navigation', { name: '文档页面' });
    const firstPage = within(pages).getByRole('button', { name: '第 1 页' });
    const secondPage = within(pages).getByRole('button', { name: '第 2 页' });
    expect(firstPage).toHaveAttribute('aria-current', 'page');
    expect(secondPage).toHaveTextContent('交付计划');

    fireEvent.click(secondPage);
    expect(secondPage).toHaveAttribute('aria-current', 'page');
    expect(
      editor.state.doc.textBetween(
        editor.state.selection.from,
        editor.state.selection.from + '交付计划'.length,
      ),
    ).toBe('交付计划');

    secondPage.focus();
    fireEvent.keyDown(secondPage, { key: 'Home' });
    await waitFor(() => expect(firstPage).toHaveFocus());
  } finally {
    editor.destroy();
    element.remove();
  }
});

test('closes compact navigation before restoring the body selection', async () => {
  const { editor, element } = createEditor();

  function CompactNavigation() {
    const [open, setOpen] = useState(true);
    return open ? (
      <DocumentNavigationPanel
        editor={editor}
        modal
        onClose={() => setOpen(false)}
      />
    ) : null;
  }

  try {
    render(<CompactNavigation />);
    const search = screen.getByRole('searchbox', { name: '搜索文档' });
    fireEvent.change(search, { target: { value: '成功标准内容' } });
    fireEvent.click(
      screen.getByRole('button', { name: '第 1 个匹配：成功标准内容' }),
    );

    await waitFor(() =>
      expect(
        screen.queryByRole('complementary', { name: '文档导航' }),
      ).not.toBeInTheDocument(),
    );
    await waitFor(() => expect(editor.view.dom).toHaveFocus());
    expect(
      editor.state.doc.textBetween(
        editor.state.selection.from,
        editor.state.selection.to,
      ),
    ).toBe('成功标准内容');
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
