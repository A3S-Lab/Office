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

test('validates exact native scale, spacing, kerning, and position ranges', () => {
  const draft = createDocumentFontDialogDraft({
    characterScale: { mixed: false, value: null },
    characterPosition: { mixed: false, value: null },
    characterSpacing: { mixed: false, value: null },
    kerningThreshold: { mixed: false, value: null },
    fontFamily: null,
    fontSize: null,
    previewText: 'A3S Office',
    selectedCharacters: 0,
  });
  expect(
    documentFontDialogDraftError({
      ...draft,
      characterScalePercent: '0',
    }),
  ).not.toBeNull();
  expect(
    documentFontDialogDraftError({
      ...draft,
      characterScalePercent: '1',
    }),
  ).toBeNull();
  expect(
    documentFontDialogDraftError({
      ...draft,
      characterScalePercent: '600',
    }),
  ).toBeNull();
  expect(
    documentFontDialogDraftError({
      ...draft,
      characterScalePercent: '601',
    }),
  ).not.toBeNull();
  expect(
    documentFontDialogDraftError({
      ...draft,
      characterScalePercent: '100.5',
    }),
  ).not.toBeNull();
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
  expect(
    documentFontDialogDraftError({
      ...draft,
      characterPositionMode: 'raised',
      characterPositionPoints: '0.4',
    }),
  ).not.toBeNull();
  expect(
    documentFontDialogDraftError({
      ...draft,
      characterPositionMode: 'raised',
      characterPositionPoints: '0.5',
    }),
  ).toBeNull();
  expect(
    documentFontDialogDraftError({
      ...draft,
      characterPositionMode: 'lowered',
      characterPositionPoints: '1.25',
    }),
  ).not.toBeNull();
  expect(
    documentFontDialogDraftError({
      ...draft,
      characterPositionMode: 'lowered',
      characterPositionPoints: '1584',
    }),
  ).toBeNull();
  expect(
    documentFontDialogDraftError({
      ...draft,
      characterPositionMode: 'lowered',
      characterPositionPoints: '1584.5',
    }),
  ).not.toBeNull();
  expect(
    documentFontDialogDraftError({
      ...draft,
      kerningEnabled: true,
      kerningThresholdPoints: '-0.5',
    }),
  ).not.toBeNull();
  expect(
    documentFontDialogDraftError({
      ...draft,
      kerningEnabled: true,
      kerningThresholdPoints: '0',
    }),
  ).toBeNull();
  expect(
    documentFontDialogDraftError({
      ...draft,
      kerningEnabled: true,
      kerningThresholdPoints: '1638.5',
    }),
  ).toBeNull();
  expect(
    documentFontDialogDraftError({
      ...draft,
      kerningEnabled: true,
      kerningThresholdPoints: '1639',
    }),
  ).not.toBeNull();
  expect(
    documentFontDialogDraftError({
      ...draft,
      kerningEnabled: true,
      kerningThresholdPoints: '12.25',
    }),
  ).not.toBeNull();
});

test('keeps independently mixed scale, spacing, kerning, and position untouched until edited', async () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content:
      '<p><span data-office-character-scale-percent="80" data-office-character-position-half-points="2" data-office-character-spacing-twips="20" data-office-kerning-threshold-half-points="24" style="font-stretch: 80%; --work-document-character-position: 1pt; letter-spacing: 1pt; font-kerning: none">Wide</span> <span data-office-character-scale-percent="125" data-office-character-position-half-points="-2" data-office-character-spacing-twips="-20" data-office-kerning-threshold-half-points="0" style="font-stretch: 125%; --work-document-character-position: -1pt; letter-spacing: -1pt; font-kerning: normal">Tight</span></p>',
  });
  editor.commands.selectAll();
  const source = documentFontDialogSource(editor);
  expect(source.characterScale).toEqual({ mixed: true, value: null });
  expect(source.characterSpacing).toEqual({ mixed: true, value: null });
  expect(source.characterPosition).toEqual({ mixed: true, value: null });
  expect(source.kerningThreshold).toEqual({ mixed: true, value: null });
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
  expect(screen.getByText(/当前选区包含多种字符缩放比例/)).toHaveTextContent(
    '当前选区包含多种字符缩放比例',
  );
  expect(
    screen.getByRole('textbox', { name: '字符缩放比例（%）' }),
  ).toHaveAttribute('placeholder', '混合');
  expect(screen.getByText(/当前选区包含多种字符间距/)).toHaveTextContent(
    '当前选区包含多种字符间距',
  );
  expect(screen.getByText(/当前选区包含多种字符位置/)).toHaveTextContent(
    '当前选区包含多种字符位置',
  );
  expect(screen.getByText(/当前选区包含不同的字距调整设置/)).toHaveTextContent(
    '当前选区包含不同的字距调整设置',
  );
  expect(
    screen.getByRole('checkbox', {
      name: '为字号达到以下值的字体调整字距',
    }),
  ).toHaveAttribute('aria-checked', 'mixed');

  fireEvent.click(screen.getByRole('combobox', { name: '字符间距' }));
  fireEvent.click(await screen.findByRole('option', { name: '加宽' }));
  const amount = screen.getByRole('textbox', { name: '间距值（磅）' });
  fireEvent.change(amount, { target: { value: '2.5' } });
  fireEvent.click(screen.getByRole('button', { name: '应用' }));

  expect(patches).toEqual([{ characterSpacingTwips: 50 }]);
});

test('applies scale, spacing, kerning, and position to the saved selection in one undo step', async () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content:
      '<p><span data-office-character-scale-percent="125" data-office-character-position-half-points="4" data-office-character-spacing-twips="40" data-office-kerning-threshold-half-points="24" style="font-stretch: 125%; --work-document-character-position: 2pt; letter-spacing: 2pt; font-kerning: none">Selected text</span> remains</p>',
  });
  document.body.append(editor.view.dom);
  const selection = textRange(editor, 'Selected text');
  editor.commands.setTextSelection(selection);
  const source = documentFontDialogSource(editor);

  render(
    <FontDialogHarness editor={editor} selection={selection} source={source} />,
  );
  editor.commands.setTextSelection(textRange(editor, 'remains'));

  fireEvent.change(screen.getByRole('textbox', { name: '字符缩放比例（%）' }), {
    target: { value: '80' },
  });
  fireEvent.click(screen.getByRole('combobox', { name: '字符间距' }));
  fireEvent.click(await screen.findByRole('option', { name: '标准' }));
  fireEvent.click(screen.getByRole('combobox', { name: '字符位置' }));
  fireEvent.click(await screen.findByRole('option', { name: '降低' }));
  fireEvent.change(screen.getByRole('textbox', { name: '位置值（磅）' }), {
    target: { value: '1.5' },
  });
  fireEvent.change(
    screen.getByRole('textbox', { name: '字距调整阈值（磅）' }),
    { target: { value: '0' } },
  );
  const preview = screen.getByLabelText('字符高级格式预览');
  await waitFor(() => {
    expect(preview.querySelector('output')?.getAttribute('style')).toContain(
      'font-stretch: 80%',
    );
    expect(preview.querySelector('output')?.getAttribute('style')).toContain(
      'letter-spacing: 0pt',
    );
    expect(preview.querySelector('output')?.getAttribute('style')).toContain(
      'font-kerning: normal',
    );
    expect(
      preview.querySelector('output > span')?.getAttribute('style'),
    ).toContain('vertical-align: -1.5pt');
  });
  fireEvent.click(screen.getByRole('button', { name: '应用' }));

  await waitFor(() =>
    expect(screen.queryByRole('dialog', { name: '字体高级设置' })).toBeNull(),
  );

  expect(editor.state.selection.toJSON()).toEqual({
    type: 'text',
    anchor: selection.from,
    head: selection.to,
  });
  expect(editor.getAttributes('textStyle').characterScalePercent).toBe(80);
  expect(editor.getAttributes('textStyle').characterSpacingTwips).toBe(0);
  expect(editor.getAttributes('textStyle').characterPositionHalfPoints).toBe(
    -3,
  );
  expect(editor.getAttributes('textStyle').kerningThresholdHalfPoints).toBe(0);
  expect(editor.getHTML()).toContain('data-office-character-spacing-twips="0"');
  expect(editor.getHTML()).toContain(
    'data-office-character-scale-percent="80"',
  );
  expect(editor.getHTML()).toContain(
    'data-office-character-position-half-points="-3"',
  );
  expect(editor.getHTML()).toContain(
    'data-office-kerning-threshold-half-points="0"',
  );

  expect(editor.commands.undo()).toBe(true);
  editor.commands.setTextSelection(selection);
  expect(editor.getAttributes('textStyle').characterScalePercent).toBe(125);
  expect(editor.getAttributes('textStyle').characterSpacingTwips).toBe(40);
  expect(editor.getAttributes('textStyle').characterPositionHalfPoints).toBe(4);
  expect(editor.getAttributes('textStyle').kerningThresholdHalfPoints).toBe(24);
  await waitFor(() => expect(editor?.view.dom).toHaveFocus());
});

test('clears direct kerning from the saved selection and restores it with one undo', async () => {
  editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content:
      '<p><span data-office-kerning-threshold-half-points="24" style="font-kerning: none">Direct kerning</span> remains</p>',
  });
  document.body.append(editor.view.dom);
  const selection = textRange(editor, 'Direct kerning');
  editor.commands.setTextSelection(selection);
  const source = documentFontDialogSource(editor);

  render(
    <FontDialogHarness editor={editor} selection={selection} source={source} />,
  );

  const kerning = screen.getByRole('checkbox', {
    name: '为字号达到以下值的字体调整字距',
  });
  expect(kerning).toBeChecked();
  fireEvent.click(kerning);
  expect(
    screen.getByRole('textbox', { name: '字距调整阈值（磅）' }),
  ).toBeDisabled();
  expect(
    screen
      .getByLabelText('字符高级格式预览')
      .querySelector('output')
      ?.getAttribute('style'),
  ).toContain('font-kerning: none');

  fireEvent.click(screen.getByRole('button', { name: '应用' }));
  await waitFor(() =>
    expect(screen.queryByRole('dialog', { name: '字体高级设置' })).toBeNull(),
  );

  editor.commands.setTextSelection(selection);
  expect(
    editor.getAttributes('textStyle').kerningThresholdHalfPoints,
  ).toBeUndefined();
  expect(editor.getHTML()).not.toContain(
    'data-office-kerning-threshold-half-points',
  );

  expect(editor.commands.undo()).toBe(true);
  editor.commands.setTextSelection(selection);
  expect(editor.getAttributes('textStyle').kerningThresholdHalfPoints).toBe(24);
  expect(editor.getHTML()).toContain(
    'data-office-kerning-threshold-half-points="24"',
  );
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
