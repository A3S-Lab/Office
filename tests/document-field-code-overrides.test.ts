import { expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import { NodeSelection } from '@tiptap/pm/state';
import {
  documentFieldShowsCode,
  selectedDocumentFieldIds,
  syncDocumentFieldCodeClasses,
  toggleDocumentFieldCodeOverrides,
} from '../src/internal/features/work/editors/document-field-code-overrides';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';

test('toggles XOR field-code overrides for selected fields', () => {
  expect(documentFieldShowsCode('a', false, new Set())).toBe(false);
  expect(documentFieldShowsCode('a', false, new Set(['a']))).toBe(true);
  expect(documentFieldShowsCode('a', true, new Set())).toBe(true);
  expect(documentFieldShowsCode('a', true, new Set(['a']))).toBe(false);

  const toggled = toggleDocumentFieldCodeOverrides(new Set(['a']), ['a', 'b']);
  expect([...toggled].sort()).toEqual(['b']);
});

test('collects selected document field ids for Shift+F9', () => {
  const editor = new Editor({
    extensions: createWorkDocumentExtensions(),
    content: '<p>Before</p>',
  });
  try {
    expect(editor.commands.insertDocumentField('page')).toBe(true);
    const fieldPos = findFieldPos(editor);
    expect(fieldPos).not.toBeNull();
    editor.view.dispatch(
      editor.state.tr.setSelection(
        NodeSelection.create(editor.state.doc, fieldPos!),
      ),
    );
    const ids = selectedDocumentFieldIds(editor);
    expect(ids).toHaveLength(1);
    expect(ids[0]).toMatch(/^field/);

    const root = document.createElement('div');
    root.innerHTML = `<span class="work-document-field" data-field-id="${ids[0]}" data-field-code="{ PAGE }"></span><span class="work-document-field" data-field-id="other" data-field-code="{ NUMPAGES }"></span>`;
    const synced = syncDocumentFieldCodeClasses(root, false, new Set(ids));
    expect(synced).toEqual({ applied: 2, showing: 1 });
    expect(
      root
        .querySelector(`[data-field-id="${ids[0]}"]`)
        ?.classList.contains('show-field-code'),
    ).toBe(true);
    expect(
      root
        .querySelector('[data-field-id="other"]')
        ?.classList.contains('show-field-code'),
    ).toBe(false);
  } finally {
    editor.destroy();
  }
});

function findFieldPos(editor: Editor): number | null {
  let found: number | null = null;
  editor.state.doc.descendants((node, pos) => {
    if (node.type.name === 'documentField') {
      found = pos;
      return false;
    }
    return true;
  });
  return found;
}
