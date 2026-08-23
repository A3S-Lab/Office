import { type Editor, Extension } from '@tiptap/core';
import { expect, test } from '@rstest/core';
import { act, render, screen, waitFor } from '@testing-library/react';
import type { DocumentContent } from '../src/core';
import { DocumentEditor } from '../src/react';
import { applyDocumentFontDialogPatch } from '../src/internal/features/work/editors/document-font-dialog-model';

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
