import { expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
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
        '<p>Alpha</p><p>Beta</p>',
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
