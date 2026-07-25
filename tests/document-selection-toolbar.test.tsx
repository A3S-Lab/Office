import { Editor } from '@tiptap/core';
import { EditorContent } from '@tiptap/react';
import { afterEach, expect, test } from '@rstest/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { DocumentSelectionToolbar } from '../src/internal/features/work/editors/document-selection-toolbar';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';

let editor: Editor | null = null;

afterEach(() => {
  editor?.destroy();
  editor = null;
});

test('formats the current text selection without losing it', async () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content:
      '<section data-document-section="true"><p>Format this selection.</p></section>',
  });
  const comments: boolean[] = [];
  const view = render(
    <>
      <EditorContent editor={editor} />
      <DocumentSelectionToolbar
        editor={editor}
        canInsertComment
        onInsertComment={() => comments.push(true)}
      />
    </>,
  );
  const range = textRange(editor, 'this selection');

  editor.chain().focus().setTextSelection(range).run();
  view.rerender(
    <>
      <EditorContent editor={editor} />
      <DocumentSelectionToolbar
        editor={editor}
        canInsertComment
        onInsertComment={() => comments.push(true)}
      />
    </>,
  );

  const toolbar = await screen.findByRole('toolbar', {
    name: '文本快捷工具栏',
  });
  await waitFor(() => expect(toolbar).toBeVisible());
  fireEvent.click(screen.getByRole('button', { name: '加粗' }));

  expect(editor.isActive('bold')).toBe(true);
  expect(editor.state.selection.empty).toBe(false);

  fireEvent.click(screen.getByRole('button', { name: '添加批注' }));
  expect(comments).toEqual([true]);
  expect(editor.state.selection.toJSON()).toEqual({
    anchor: range.from,
    head: range.to,
    type: 'text',
  });
});

test('hides for a caret and disables unavailable comment insertion', async () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content:
      '<section data-document-section="true"><p>Selection state.</p></section>',
  });
  const view = render(
    <>
      <EditorContent editor={editor} />
      <DocumentSelectionToolbar
        editor={editor}
        canInsertComment={false}
        onInsertComment={() => undefined}
      />
    </>,
  );
  const range = textRange(editor, 'Selection');

  editor.chain().focus().setTextSelection(range).run();
  view.rerender(
    <>
      <EditorContent editor={editor} />
      <DocumentSelectionToolbar
        editor={editor}
        canInsertComment={false}
        onInsertComment={() => undefined}
      />
    </>,
  );

  const toolbar = await screen.findByRole('toolbar', {
    name: '文本快捷工具栏',
  });
  await waitFor(() => expect(toolbar).toBeVisible());
  expect(screen.getByRole('button', { name: '添加批注' })).toBeDisabled();

  editor.commands.setTextSelection(range.from);
  view.rerender(
    <>
      <EditorContent editor={editor} />
      <DocumentSelectionToolbar
        editor={editor}
        canInsertComment={false}
        onInsertComment={() => undefined}
      />
    </>,
  );
  await waitFor(() => expect(toolbar).not.toBeVisible());
});

test('hides when focus moves outside the editor and toolbar', async () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content:
      '<section data-document-section="true"><p>Selection focus.</p></section>',
  });
  const view = render(
    <>
      <button type="button">审阅</button>
      <EditorContent editor={editor} />
      <DocumentSelectionToolbar
        editor={editor}
        canInsertComment
        onInsertComment={() => undefined}
      />
    </>,
  );
  const range = textRange(editor, 'Selection');

  editor.chain().focus().setTextSelection(range).run();
  view.rerender(
    <>
      <button type="button">审阅</button>
      <EditorContent editor={editor} />
      <DocumentSelectionToolbar
        editor={editor}
        canInsertComment
        onInsertComment={() => undefined}
      />
    </>,
  );

  const toolbar = await screen.findByRole('toolbar', {
    name: '文本快捷工具栏',
  });
  await waitFor(() => expect(toolbar).toBeVisible());
  fireEvent.mouseDown(screen.getByRole('button', { name: '加粗' }));
  screen.getByRole('button', { name: '审阅' }).focus();
  await waitFor(() => expect(toolbar).not.toBeVisible());
});

function textRange(current: Editor, text: string) {
  let range: { from: number; to: number } | null = null;
  current.state.doc.descendants((node, position) => {
    if (!node.isText || !node.text || range) return;
    const offset = node.text.indexOf(text);
    if (offset < 0) return;
    range = {
      from: position + offset,
      to: position + offset + text.length,
    };
  });
  if (!range) throw new Error(`Expected document text: ${text}`);
  return range;
}
