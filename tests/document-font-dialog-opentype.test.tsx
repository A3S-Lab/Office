import { afterEach, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import { fireEvent, render, screen, waitFor } from '@testing-library/react';
import { useState } from 'react';
import { DocumentFontDialog } from '../src/internal/features/work/editors/document-font-dialog';
import {
  applyDocumentFontDialogPatch,
  documentFontDialogSource,
  type DocumentFontDialogPatch,
} from '../src/internal/features/work/editors/document-font-dialog-model';
import { collectDocumentChanges } from '../src/internal/features/work/work-document-changes';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import {
  documentOpenTypeDomAttributes,
  parseDocumentOpenTypeFeatures,
  type WorkDocumentOpenTypeFeatures,
} from '../src/internal/features/work/work-document-opentype';

let editor: Editor | null = null;

afterEach(() => {
  editor?.destroy();
  editor = null;
});

test('reports independently mixed OpenType fields and emits only the touched field', async () => {
  editor = createEditor(
    `<p>${openTypeSpan(
      {
        ligatures: 'standard',
        numberForm: 'oldStyle',
        numberSpacing: 'tabular',
        stylisticSets: [1],
        contextualAlternates: true,
      },
      'First',
    )}${openTypeSpan(
      {
        ligatures: 'historical',
        numberForm: 'lining',
        numberSpacing: 'proportional',
        stylisticSets: [2],
        contextualAlternates: false,
      },
      'Second',
    )}</p>`,
  );
  editor.commands.selectAll();
  const source = documentFontDialogSource(editor);

  expect(source.openTypeLigatures).toEqual({ mixed: true, value: null });
  expect(source.openTypeNumberForm).toEqual({ mixed: true, value: null });
  expect(source.openTypeNumberSpacing).toEqual({ mixed: true, value: null });
  expect(source.openTypeStylisticSets).toEqual({ mixed: true, value: null });
  expect(source.openTypeContextualAlternates).toEqual({
    mixed: true,
    value: null,
  });
  const patches: DocumentFontDialogPatch[] = [];

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
  expect(screen.getByText(/不同的 OpenType 排版设置/)).toBeVisible();

  await selectOption('OpenType 连字', '全部');
  fireEvent.click(screen.getByRole('button', { name: '应用' }));

  expect(patches).toEqual([{ openTypeLigatures: 'all' }]);
});

test('preserves every untouched per-run feature and restores the edit with one undo', async () => {
  const first: WorkDocumentOpenTypeFeatures = {
    ligatures: 'standard',
    numberForm: 'oldStyle',
    numberSpacing: 'tabular',
    stylisticSets: [4, 1],
    contextualAlternates: true,
  };
  const second: WorkDocumentOpenTypeFeatures = {
    ligatures: 'historical',
    numberForm: 'lining',
    numberSpacing: 'proportional',
    stylisticSets: [20],
    contextualAlternates: false,
  };
  editor = createEditor(
    `<p>${openTypeSpan(first, 'First')}${openTypeSpan(second, 'Second')}</p>`,
  );
  document.body.append(editor.view.dom);
  editor.commands.selectAll();
  const selection = {
    from: editor.state.selection.from,
    to: editor.state.selection.to,
  };

  render(
    <FontDialogHarness
      editor={editor}
      selection={selection}
      source={documentFontDialogSource(editor)}
    />,
  );

  await selectOption('OpenType 连字', '全部');
  fireEvent.click(screen.getByRole('button', { name: '应用' }));
  await waitFor(() =>
    expect(screen.queryByRole('dialog', { name: '字体高级设置' })).toBeNull(),
  );

  expect(openTypeRuns(editor)).toEqual([
    { ...first, ligatures: 'all' },
    { ...second, ligatures: 'all' },
  ]);
  expect(editor.commands.undo()).toBe(true);
  expect(openTypeRuns(editor)).toEqual([first, second]);
  expect(editor.commands.undo()).toBe(false);
  await waitFor(() => expect(editor?.view.dom).toHaveFocus());
});

test('offers every native control and applies one exact typography patch', async () => {
  editor = createEditor('<p>Office 0123</p>');
  editor.commands.selectAll();
  const patches: DocumentFontDialogPatch[] = [];

  render(
    <DocumentFontDialog
      source={documentFontDialogSource(editor)}
      restoreFocusTarget={() => editor?.view.dom ?? null}
      onApply={(patch) => {
        patches.push(patch);
        return true;
      }}
      onClose={() => undefined}
    />,
  );

  await selectOption('OpenType 连字', '标准和上下文');
  await selectOption('OpenType 数字字形', '旧式数字');
  await selectOption('OpenType 数字间距', '等宽');
  await selectOption('OpenType 样式集', '样式集 4');
  await selectOption('OpenType 上下文替代', '禁用');
  fireEvent.click(screen.getByRole('button', { name: '应用' }));

  expect(patches).toEqual([
    {
      openTypeLigatures: 'standardContextual',
      openTypeNumberForm: 'oldStyle',
      openTypeNumberSpacing: 'tabular',
      openTypeStylisticSets: [4],
      openTypeContextualAlternates: false,
    },
  ]);
});

test('tracks and rejects an exact OpenType formatting change', () => {
  const original: WorkDocumentOpenTypeFeatures = {
    ligatures: 'standard',
    numberForm: 'oldStyle',
    numberSpacing: 'tabular',
    stylisticSets: [4],
    contextualAlternates: true,
  };
  editor = new Editor({
    extensions: createWorkDocumentExtensions({
      isTracking: () => true,
      createChange: () => ({
        id: 'formatting-opentype',
        author: 'Reviewer',
        date: '2026-09-01T01:00:00.000Z',
      }),
    }),
    content: `<p>${openTypeSpan(original, 'Tracked typography')}</p>`,
  });
  editor.commands.selectAll();
  const selection = {
    from: editor.state.selection.from,
    to: editor.state.selection.to,
  };

  expect(
    applyDocumentFontDialogPatch(editor, selection, {
      openTypeLigatures: 'all',
      openTypeContextualAlternates: false,
    }),
  ).toBe(true);
  expect(collectDocumentChanges(editor.state.doc)).toEqual([
    expect.objectContaining({
      id: 'formatting-opentype',
      kind: 'formatting',
      text: 'Tracked typography',
    }),
  ]);
  expect(editor.getHTML()).toContain('&quot;openTypeFeatures&quot;');

  expect(editor.commands.rejectDocumentChange('formatting-opentype')).toBe(
    true,
  );
  editor.commands.selectAll();
  expect(
    parseDocumentOpenTypeFeatures(
      editor.getAttributes('textStyle').openTypeFeatures,
    ),
  ).toEqual(original);
});

function FontDialogHarness({
  editor: currentEditor,
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
      restoreFocusTarget={() => currentEditor.view.dom}
      onApply={(patch) =>
        applyDocumentFontDialogPatch(currentEditor, selection, patch)
      }
      onClose={() => setOpen(false)}
    />
  ) : null;
}

async function selectOption(ariaLabel: string, option: string): Promise<void> {
  fireEvent.click(screen.getByRole('combobox', { name: ariaLabel }));
  fireEvent.click(await screen.findByRole('option', { name: option }));
}

function createEditor(content: string): Editor {
  return new Editor({
    extensions: createWorkDocumentExtensions(),
    content,
  });
}

function openTypeSpan(
  features: WorkDocumentOpenTypeFeatures,
  text: string,
): string {
  const span = document.createElement('span');
  for (const [name, value] of Object.entries(
    documentOpenTypeDomAttributes(features),
  )) {
    span.setAttribute(name, value);
  }
  span.textContent = text;
  return span.outerHTML;
}

function openTypeRuns(currentEditor: Editor): WorkDocumentOpenTypeFeatures[] {
  const currentDocument = new DOMParser().parseFromString(
    currentEditor.getHTML(),
    'text/html',
  );
  return Array.from(
    currentDocument.querySelectorAll<HTMLElement>(
      '[data-office-opentype-features]',
    ),
  ).map((element) => {
    const features = parseDocumentOpenTypeFeatures(
      element.dataset.officeOpentypeFeatures,
    );
    if (!features) throw new Error('Expected valid OpenType features.');
    return features;
  });
}
