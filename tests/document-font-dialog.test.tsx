import { afterEach, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { useState } from 'react';
import { DocumentFontDialog } from '../src/internal/features/work/editors/document-font-dialog';
import {
  applyDocumentFontDialogPatch,
  createDocumentFontDialogDraft,
  documentFontDialogDraftError,
  documentFontDialogSource,
} from '../src/internal/features/work/editors/document-font-dialog-model';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';

let editor: Editor | null = null;

afterEach(() => {
  editor?.destroy();
  editor = null;
});

test('validates the exact native character-spacing range', () => {
  const draft = createDocumentFontDialogDraft({
    characterSpacing: { mixed: false, value: null },
    fontFamily: null,
    fontSize: null,
    previewText: 'A3S Office',
    selectedCharacters: 0,
  });
  expect(
    documentFontDialogDraftError({
      ...draft,
      characterSpacingMode: 'expanded',
      characterSpacingPoints: '0.04',
    }),
  ).not.toBeNull();
  expect(
    documentFontDialogDraftError({
      ...draft,
      characterSpacingMode: 'expanded',
      characterSpacingPoints: '0.05',
    }),
  ).toBeNull();
  expect(
    documentFontDialogDraftError({
      ...draft,
      characterSpacingMode: 'condensed',
      characterSpacingPoints: '1584',
    }),
  ).toBeNull();
  expect(
    documentFontDialogDraftError({
      ...draft,
      characterSpacingMode: 'condensed',
      characterSpacingPoints: '1584.01',
    }),
  ).not.toBeNull();
});

test('leaves mixed character spacing untouched until the user chooses a mode', async () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content:
      '<p><span data-office-character-spacing-twips="20" style="letter-spacing: 1pt">Wide</span> <span data-office-character-spacing-twips="-20" style="letter-spacing: -1pt">Tight</span></p>',
  });
  editor.commands.selectAll();
  const source = documentFontDialogSource(editor);
  expect(source.characterSpacing).toEqual({ mixed: true, value: null });
  const patches: unknown[] = [];

  render(
    <DocumentFontDialog
      source={source}
      restoreFocusTarget={() => editor?.view.dom ?? null}
      onApply={(patch) => {
        patches.push(patch);
        return true;
      }}
      onClose={() => undefined}
    />,
  );

  expect(screen.getByRole('button', { name: '应用' })).toBeDisabled();
  expect(screen.getByText(/当前选区包含多种字符间距/)).toHaveTextContent(
    '当前选区包含多种字符间距',
  );

  fireEvent.click(screen.getByRole('combobox', { name: '字符间距' }));
  fireEvent.click(await screen.findByRole('option', { name: '加宽' }));
  const amount = screen.getByRole('textbox', { name: '间距值（磅）' });
  fireEvent.change(amount, { target: { value: '2.5' } });
  fireEvent.click(screen.getByRole('button', { name: '应用' }));

  expect(patches).toEqual([{ characterSpacingTwips: 50 }]);
});

test('restores the saved selection and applies explicit normal spacing in one undo step', async () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content:
      '<p><span data-office-character-spacing-twips="40" style="letter-spacing: 2pt">Selected text</span> remains</p>',
  });
  document.body.append(editor.view.dom);
  const selection = textRange(editor, 'Selected text');
  editor.commands.setTextSelection(selection);
  const source = documentFontDialogSource(editor);

  render(
    <FontDialogHarness editor={editor} selection={selection} source={source} />,
  );
  editor.commands.setTextSelection(textRange(editor, 'remains'));

  fireEvent.click(screen.getByRole('combobox', { name: '字符间距' }));
  fireEvent.click(await screen.findByRole('option', { name: '标准' }));
  fireEvent.click(screen.getByRole('button', { name: '应用' }));

  await waitFor(() =>
    expect(screen.queryByRole('dialog', { name: '字体高级设置' })).toBeNull(),
  );

  expect(editor.state.selection.toJSON()).toEqual({
    type: 'text',
    anchor: selection.from,
    head: selection.to,
  });
  expect(editor.getAttributes('textStyle').characterSpacingTwips).toBe(0);
  expect(editor.getHTML()).toContain('data-office-character-spacing-twips="0"');

  expect(editor.commands.undo()).toBe(true);
  editor.commands.setTextSelection(selection);
  expect(editor.getAttributes('textStyle').characterSpacingTwips).toBe(40);
  await waitFor(() => expect(editor?.view.dom).toHaveFocus());
});

function FontDialogHarness({
  editor,
  selection,
  source,
}: {
  editor: Editor;
  selection: { from: number; to: number };
  source: ReturnType<typeof documentFontDialogSource>;
}) {
  const [open, setOpen] = useState(true);
  return open ? (
    <DocumentFontDialog
      source={source}
      restoreFocusTarget={() => editor.view.dom}
      onApply={(patch) =>
        applyDocumentFontDialogPatch(editor, selection, patch)
      }
      onClose={() => setOpen(false)}
    />
  ) : null;
}

function textRange(
  currentEditor: Editor,
  text: string,
): { from: number; to: number } {
  let range: { from: number; to: number } | null = null;
  currentEditor.state.doc.descendants((node, position) => {
    if (range || !node.isText || !node.text) return;
    const offset = node.text.indexOf(text);
    if (offset < 0) return;
    range = {
      from: position + offset,
      to: position + offset + text.length,
    };
  });
  if (!range) throw new Error(`Text "${text}" was not found.`);
  return range;
}
