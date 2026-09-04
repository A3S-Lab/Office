import { type Editor, Extension } from '@tiptap/core';
import { expect, test } from '@rstest/core';
import {
  act,
  fireEvent,
  render,
  screen,
  waitFor,
} from '@testing-library/react';
import { useState } from 'react';
import type { DocumentContent } from '../src/core';
import { DocumentEditor } from '../src/react';
import { applyDocumentFontDialogPatch } from '../src/internal/features/work/editors/document-font-dialog-model';

test('publishes only committed Chinese text from a controlled IME composition', async () => {
  let editor: Editor | null = null;
  const publications: DocumentContent[] = [];
  const captureEditor = Extension.create({
    name: 'captureControlledCompositionEditor',
    onCreate() {
      editor = this.editor;
    },
  });
  const initial: DocumentContent = {
    type: 'document',
    html: '<p></p>',
    pageSize: 'a4',
  };

  function ControlledDocumentEditor() {
    const [content, setContent] = useState(initial);
    return (
      <DocumentEditor
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

  render(<ControlledDocumentEditor />);
  const surface = await screen.findByRole('textbox', { name: '文档正文' });
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
  await waitFor(() => expect(publications).toHaveLength(1));
  expect(publications[0]?.html).toContain('请问');
  expect(publications[0]?.html).not.toContain('qingwen');
  expect(surface).toHaveTextContent('请问');
});

test('defers an authoritative controlled replacement until IME composition settles', async () => {
  let editor: Editor | null = null;
  const publications: DocumentContent[] = [];
  const captureEditor = Extension.create({
    name: 'captureDeferredControlledCompositionEditor',
    onCreate() {
      editor = this.editor;
    },
  });
  const initial: DocumentContent = {
    type: 'document',
    html: '<p>Initial</p>',
    pageSize: 'a4',
  };
  const replacement: DocumentContent = {
    ...initial,
    html: '<p>Authoritative host update</p>',
  };
  const properties = {
    extensions: [captureEditor],
    onChange: (next: DocumentContent) => publications.push(next),
    theme: 'light' as const,
  };
  const view = render(<DocumentEditor {...properties} content={initial} />);
  const surface = await screen.findByRole('textbox', { name: '文档正文' });
  await waitFor(() => expect(editor).not.toBeNull());
  const current = editor as Editor;

  fireEvent.compositionStart(surface, { data: 'qingwen' });
  act(() => {
    current.chain().selectAll().insertContent('qingwen').run();
  });
  view.rerender(<DocumentEditor {...properties} content={replacement} />);

  expect(surface).toHaveTextContent('qingwen');
  expect(surface).not.toHaveTextContent('Authoritative host update');
  expect(publications).toEqual([]);

  act(() => {
    current.chain().selectAll().insertContent('请问').run();
  });
  fireEvent.compositionEnd(surface, { data: '请问' });

  await waitFor(() => expect(publications).toHaveLength(1));
  expect(publications[0]?.html).toContain('请问');
  expect(publications[0]?.html).not.toContain('qingwen');
  await waitFor(() =>
    expect(surface).toHaveTextContent('Authoritative host update'),
  );
});

test('normalizes a WebKit commit that arrives after compositionend', async () => {
  let editor: Editor | null = null;
  const publications: DocumentContent[] = [];
  const captureEditor = Extension.create({
    name: 'captureLateControlledCompositionEditor',
    onCreate() {
      editor = this.editor;
    },
  });
  const initial: DocumentContent = {
    type: 'document',
    html: '<p></p>',
    pageSize: 'a4',
  };

  function ControlledDocumentEditor() {
    const [content, setContent] = useState(initial);
    return (
      <DocumentEditor
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

  render(<ControlledDocumentEditor />);
  const surface = await screen.findByRole('textbox', { name: '文档正文' });
  await waitFor(() => expect(editor).not.toBeNull());
  const current = editor as Editor;

  fireEvent.compositionStart(surface, { data: 'ni hao' });
  act(() => {
    current.commands.insertContent('ni hao');
  });
  expect(surface).toHaveTextContent('ni hao');
  expect(publications).toEqual([]);

  fireEvent.compositionEnd(surface, { data: '你好' });
  // WebKit's late insertFromComposition path can append the committed text
  // after compositionend has already reached the editor.
  const committedInput = new Event('beforeinput', { bubbles: true });
  Object.defineProperties(committedInput, {
    data: { value: '你好' },
    inputType: { value: 'insertFromComposition' },
  });
  surface.dispatchEvent(committedInput);
  // The browser's `input` event still precedes ProseMirror's mutation
  // observer flush. The tracker must not freeze the range until that flush
  // has produced its transaction.
  const committedInputEvent = new Event('input', { bubbles: true });
  Object.defineProperties(committedInputEvent, {
    data: { value: '你好' },
    inputType: { value: 'insertFromComposition' },
  });
  surface.dispatchEvent(committedInputEvent);
  act(() => {
    current.commands.insertContent('你好');
  });

  await waitFor(() => expect(publications).toHaveLength(1), { timeout: 1_000 });
  await waitFor(() => expect(surface).toHaveTextContent('你好'));
  expect(surface).not.toHaveTextContent('ni hao');
  expect(current.getText()).not.toContain('ni hao');
  expect(current.getText()).toContain('你好');
  expect(publications).toHaveLength(1);
  expect(publications[0]?.html).toContain('你好');
  expect(publications[0]?.html).not.toContain('ni hao');
});

test('folds a duplicate WebKit commit when the range already contains Chinese', async () => {
  let editor: Editor | null = null;
  const publications: DocumentContent[] = [];
  const captureEditor = Extension.create({
    name: 'captureAlreadyCommittedCompositionEditor',
    onCreate() {
      editor = this.editor;
    },
  });
  const initial: DocumentContent = {
    type: 'document',
    html: '<p></p>',
    pageSize: 'a4',
  };

  function ControlledDocumentEditor() {
    const [content, setContent] = useState(initial);
    return (
      <DocumentEditor
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

  render(<ControlledDocumentEditor />);
  const surface = await screen.findByRole('textbox', { name: '文档正文' });
  await waitFor(() => expect(editor).not.toBeNull());
  const current = editor as Editor;

  fireEvent.compositionStart(surface, { data: 'ni hao' });
  act(() => {
    current.commands.insertContent('你好');
  });
  fireEvent.compositionEnd(surface, { data: '你好' });

  // Some WebKit builds expose the committed text once, then report the same
  // commit again after compositionend. The second transaction must be folded
  // into the tracked range rather than leaving "你好你好" in the document.
  const committedInput = new Event('input', { bubbles: true });
  Object.defineProperties(committedInput, {
    data: { value: '你好' },
    inputType: { value: 'insertFromComposition' },
  });
  surface.dispatchEvent(committedInput);
  act(() => {
    current.commands.insertContent('你好');
  });

  await waitFor(() => expect(publications).toHaveLength(1), { timeout: 1_000 });
  expect(current.getText()).toContain('你好');
  expect(current.getText()).not.toContain('你好你好');
  expect(publications[0]?.html).toContain('你好');
  expect(publications[0]?.html).not.toContain('你好你好');
});

test('keeps a local formatting publication until the controlled host acknowledges it', async () => {
  let editor: Editor | null = null;
  let published: DocumentContent | null = null;
  let publicationCount = 0;
  const initial: DocumentContent = {
    type: 'document',
    html: '<p>Controlled formatting</p>',
    pageSize: 'a4',
  };
  const captureEditor = Extension.create({
    name: 'captureControlledPublicationEditor',
    onCreate() {
      editor = this.editor;
    },
  });
  const properties = {
    content: initial,
    extensions: [captureEditor],
    onChange: (next: DocumentContent) => {
      publicationCount += 1;
      published = next;
    },
    theme: 'light' as const,
  };
  const view = render(<DocumentEditor {...properties} />);
  const surface = await screen.findByRole('textbox', { name: '文档正文' });
  await waitFor(() => expect(editor).not.toBeNull());
  await waitFor(() =>
    expect(surface).toHaveAttribute('data-document-editor-mount-ms'),
  );
  let textSelection: { from: number; to: number } | null = null;
  editor?.state.doc.descendants((node, position) => {
    if (
      textSelection ||
      !node.isText ||
      node.text !== 'Controlled formatting'
    ) {
      return;
    }
    textSelection = {
      from: position,
      to: position + node.nodeSize,
    };
  });
  const selectedRange = textSelection;
  if (!selectedRange) throw new Error('Expected the controlled text range.');

  act(() => {
    editor?.commands.setTextSelection(selectedRange);
  });
  const selection = editor?.state.selection;
  if (!editor || !selection) throw new Error('Expected a mounted editor.');
  expect(editor.schema.marks.textStyle?.spec.attrs).toHaveProperty(
    'characterScalePercent',
  );
  let updateCount = 0;
  editor.on('update', () => {
    updateCount += 1;
  });

  act(() => {
    expect(
      applyDocumentFontDialogPatch(
        editor as Editor,
        { from: selection.from, to: selection.to },
        { characterScalePercent: 80, emphasisMark: 'circle' },
      ),
    ).toBe(true);
    // Closing the real font dialog restores its saved selection. That local
    // selection render must not make the still-old controlled prop authoritative.
    editor?.commands.setTextSelection({
      from: selection.from,
      to: selection.to,
    });
  });

  await waitFor(() =>
    expect({ publicationCount, updateCount }).toMatchObject({
      publicationCount: 1,
      updateCount: 1,
    }),
  );
  expect(
    surface.querySelector(
      'span[data-office-character-scale-percent="80"][data-office-emphasis-mark="circle"]',
    ),
  ).toHaveTextContent('Controlled formatting');

  const acknowledged = published;
  if (!acknowledged) throw new Error('Expected a published document snapshot.');
  view.rerender(
    <DocumentEditor
      {...properties}
      content={acknowledged as DocumentContent}
    />,
  );

  await waitFor(() =>
    expect(
      surface.querySelector(
        'span[data-office-character-scale-percent="80"][data-office-emphasis-mark="circle"]',
      ),
    ).toHaveTextContent('Controlled formatting'),
  );
  expect(publicationCount).toBe(1);
});
