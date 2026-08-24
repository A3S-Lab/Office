import { expect, test } from '@rstest/core';
import {
  fireEvent,
  render,
  screen,
  waitFor,
  within,
} from '@testing-library/react';
import { Editor } from '@tiptap/core';
import { useRef } from 'react';
import { useDocumentInsertCommands } from '../src/internal/features/work/editors/use-document-insert-commands';
import { editorDocumentCaptionTargets } from '../src/internal/features/work/work-document-caption-nodes';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import type { WorkDocumentContent } from '../src/internal/features/work/work-types';

test('uses a focused caption dialog and returns focus to the document after insert', () => {
  const { editor, element } = createEditor();

  try {
    render(<InsertDialogHarness editor={editor} />);
    fireEvent.click(screen.getByRole('button', { name: '打开图片题注' }));

    const field = screen.getByRole('textbox', { name: '题注文字' });
    expect(field).toHaveFocus();
    fireEvent.change(field, { target: { value: '系统架构' } });
    fireEvent.click(screen.getByRole('button', { name: '插入题注' }));

    expect(editorDocumentCaptionTargets(editor)[0]).toMatchObject({
      display: '图 1',
      title: '系统架构',
    });
    expect(editor.view.dom).toHaveFocus();
  } finally {
    editor.destroy();
    element.remove();
  }
});

test('chooses a cross-reference from existing captions instead of free text', () => {
  const { editor, element } = createEditor();
  editor.commands.setTextSelection(6);
  editor.commands.insertDocumentCaption('figure', '系统架构');

  try {
    render(<InsertDialogHarness editor={editor} />);
    fireEvent.click(screen.getByRole('button', { name: '打开交叉引用' }));

    const target = screen.getByRole('radio', { name: '图 1 系统架构' });
    expect(target).toBeChecked();
    fireEvent.click(screen.getByRole('button', { name: '插入引用' }));

    expect(editor.getHTML()).toContain('data-document-cross-reference="true"');
  } finally {
    editor.destroy();
    element.remove();
  }
});

test('moves through cross-reference targets with arrow keys', async () => {
  const { editor, element } = createEditor();
  editor.commands.setTextSelection(6);
  editor.commands.insertDocumentCaption('figure', '系统架构');
  editor.commands.setTextSelection(editor.state.doc.content.size - 1);
  editor.commands.insertDocumentCaption('table', '能力矩阵');

  try {
    render(<InsertDialogHarness editor={editor} />);
    fireEvent.click(screen.getByRole('button', { name: '打开交叉引用' }));

    const first = screen.getByRole('radio', { name: '图 1 系统架构' });
    const second = screen.getByRole('radio', { name: '表 1 能力矩阵' });
    await waitFor(() => expect(first).toHaveFocus());
    fireEvent.keyDown(first, { key: 'ArrowDown' });
    await waitFor(() => expect(second).toHaveFocus());
    expect(second).toBeChecked();
  } finally {
    editor.destroy();
    element.remove();
  }
});

test('offers a body bookmark as a cross-reference target', () => {
  const { editor, element } = createEditor();
  editor.commands.setTextSelection(textRange(editor, 'Alpha'));
  editor.commands.insertDocumentBookmark('Alpha_target');

  try {
    render(<InsertDialogHarness editor={editor} />);
    fireEvent.click(screen.getByRole('button', { name: '打开交叉引用' }));

    const target = screen.getByRole('radio', {
      name: '书签 Alpha_target Alpha',
    });
    expect(target).toBeChecked();
    fireEvent.click(screen.getByRole('button', { name: '插入引用' }));

    expect(editor.getHTML()).toContain('data-reference-target-type="bookmark"');
    expect(editor.getHTML()).toContain(
      'data-reference-target-name="Alpha_target"',
    );
  } finally {
    editor.destroy();
    element.remove();
  }
});

test('inserts a configured table of contents and restores document focus', async () => {
  const { editor, element } = createEditor();
  editor.commands.setTextSelection(textRange(editor, 'Alpha').from);

  try {
    render(<InsertDialogHarness editor={editor} />);
    fireEvent.click(screen.getByRole('button', { name: '打开目录' }));

    const dialog = screen.getByRole('dialog', { name: '插入目录' });
    expect(
      within(dialog).getByRole('combobox', { name: '起始标题级别' }),
    ).toHaveFocus();
    fireEvent.click(
      within(dialog).getByRole('combobox', { name: '目录前导符' }),
    );
    fireEvent.click(screen.getByRole('option', { name: '短横线（----）' }));
    fireEvent.click(within(dialog).getByRole('button', { name: '插入目录' }));

    await waitFor(() =>
      expect(screen.queryByRole('dialog', { name: '插入目录' })).toBeNull(),
    );
    const html = editor.getHTML();
    expect(html).toContain('data-document-table-of-contents="true"');
    expect(html).toContain('data-toc-leader="dash"');
    expect(html).toContain('Alpha heading');
    expect(html).toContain('Beta heading');
    expect(editor.view.dom).toHaveFocus();
  } finally {
    editor.destroy();
    element.remove();
  }
});

test('marks the selected text as a configured index entry and restores focus', async () => {
  const { editor, element } = createEditor();
  editor.commands.setTextSelection(textRange(editor, 'Alpha'));

  try {
    render(<InsertDialogHarness editor={editor} />);
    fireEvent.click(screen.getByRole('button', { name: '打开索引项' }));

    const dialog = screen.getByRole('dialog', { name: '标记索引项' });
    const mainEntry = within(dialog).getByRole('textbox', {
      name: '主索引项',
    });
    expect(mainEntry).toHaveFocus();
    expect(mainEntry).toHaveValue('Alpha');
    fireEvent.change(
      within(dialog).getByRole('textbox', { name: '次索引项' }),
      { target: { value: 'Runtime' } },
    );
    fireEvent.click(within(dialog).getByRole('checkbox', { name: '页码加粗' }));
    fireEvent.click(within(dialog).getByRole('button', { name: '标记' }));

    await waitFor(() =>
      expect(screen.queryByRole('dialog', { name: '标记索引项' })).toBeNull(),
    );
    expect(editor.getHTML()).toContain('data-document-index-entry="true"');
    expect(editor.getHTML()).toContain('data-index-main-entry="Alpha"');
    expect(editor.getHTML()).toContain('data-index-sub-entry="Runtime"');
    expect(editor.getHTML()).toContain('data-index-page-bold="true"');
    expect(editor.view.dom).toHaveFocus();
  } finally {
    editor.destroy();
    element.remove();
  }
});

test('inserts a configured native index from marked entries', async () => {
  const { editor, element } = createEditor();
  editor.commands.setTextSelection(textRange(editor, 'Alpha'));
  editor.commands.markDocumentIndexEntry({
    mainEntry: 'Alpha',
    subEntry: '',
    crossReference: '',
    pageBold: false,
    pageItalic: false,
  });
  editor.commands.setTextSelection(textRange(editor, 'Beta').from);

  try {
    render(<InsertDialogHarness editor={editor} />);
    fireEvent.click(screen.getByRole('button', { name: '打开索引' }));

    const dialog = screen.getByRole('dialog', { name: '插入索引' });
    expect(
      within(dialog).getByRole('combobox', { name: '索引栏数' }),
    ).toHaveFocus();
    fireEvent.click(within(dialog).getByRole('combobox', { name: '索引栏数' }));
    fireEvent.click(screen.getByRole('option', { name: '两栏' }));
    fireEvent.click(
      within(dialog).getByRole('combobox', { name: '索引前导符' }),
    );
    fireEvent.click(screen.getByRole('option', { name: '短横线（----）' }));
    fireEvent.click(within(dialog).getByRole('button', { name: '插入索引' }));

    await waitFor(() =>
      expect(screen.queryByRole('dialog', { name: '插入索引' })).toBeNull(),
    );
    expect(editor.getHTML()).toContain('data-document-index="true"');
    expect(editor.getHTML()).toContain('data-index-columns="2"');
    expect(editor.getHTML()).toContain('data-index-leader="dash"');
    expect(editor.getHTML()).toContain('Alpha');
    expect(editor.view.dom).toHaveFocus();
  } finally {
    editor.destroy();
    element.remove();
  }
});

function InsertDialogHarness({ editor }: { editor: Editor }) {
  const contentRef = useRef<WorkDocumentContent>({
    type: 'document',
    html: editor.getHTML(),
  });
  const commands = useDocumentInsertCommands({ contentRef, editor });
  return (
    <>
      <button type="button" onClick={() => commands.insertCaption('figure')}>
        打开图片题注
      </button>
      <button type="button" onClick={commands.insertCrossReference}>
        打开交叉引用
      </button>
      <button type="button" onClick={commands.openTableOfContents}>
        打开目录
      </button>
      <button type="button" onClick={commands.openIndexEntry}>
        打开索引项
      </button>
      <button type="button" onClick={commands.openIndex}>
        打开索引
      </button>
      {commands.dialog}
    </>
  );
}

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
      content: [
        '<section data-document-section="true" data-section-id="section-1">',
        '<p>Alpha</p><p>Beta</p><h1>Alpha heading</h1><h2>Beta heading</h2>',
        '</section>',
      ].join(''),
    }),
    element,
  };
}

function textRange(editor: Editor, text: string): { from: number; to: number } {
  let range: { from: number; to: number } | null = null;
  editor.state.doc.descendants((node, position) => {
    if (range || !node.isText || !node.text) return;
    const offset = node.text.indexOf(text);
    if (offset < 0) return;
    range = {
      from: position + offset,
      to: position + offset + text.length,
    };
  });
  if (!range) throw new Error(`Unable to find "${text}".`);
  return range;
}
