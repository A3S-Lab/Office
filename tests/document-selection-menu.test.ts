import { Editor } from '@tiptap/core';
import { closeHistory } from '@tiptap/pm/history';
import { afterEach, describe, expect, test } from '@rstest/core';
import { collectDocumentChanges } from '../src/internal/features/work/work-document-changes';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import {
  createWorkDocumentSelectionAction,
  createWorkDocumentSelectionSnapshot,
} from '../src/internal/features/work/work-document-selection-menu';
import type { WorkDocumentContent } from '../src/internal/features/work/work-types';

const content: WorkDocumentContent = {
  type: 'document',
  html: [
    '<section data-document-section="true">',
    '<p>Before target text after.</p>',
    '<p>Second paragraph.</p>',
    '</section>',
  ].join(''),
  pageSize: 'a4',
};

let editor: Editor | null = null;

afterEach(() => {
  editor?.destroy();
  editor = null;
});

describe('document selection menu context', () => {
  test('provides the selection, nearby text, structured fragment, and full document', () => {
    editor = createEditor();
    editor.commands.setTextSelection(textRange(editor, 'target text'));

    const snapshot = createWorkDocumentSelectionSnapshot(editor, content);

    expect(snapshot).not.toBeNull();
    expect(snapshot?.selection).toMatchObject({
      text: 'target text',
      rawText: 'target text',
      beforeText: expect.stringContaining('Before '),
      afterText: expect.stringContaining(' after.'),
      model: {
        type: 'doc',
        content: expect.any(Array),
      },
    });
    expect(snapshot?.document.content).toBe(content);
    expect(snapshot?.document.text).toContain('Second paragraph.');
    expect(snapshot?.document.html).toContain(
      '<p>Before target text after.</p>',
    );
  });

  test('tracks the target while an asynchronous host action changes text before it', () => {
    editor = createEditor();
    const range = textRange(editor, 'target text');
    editor.commands.setTextSelection(range);
    const snapshot = requiredSnapshot(editor);
    const action = createWorkDocumentSelectionAction(
      editor,
      snapshot,
      () => false,
    );

    expect(editor.commands.insertContentAt(range.from, 'new ')).toBe(true);
    editor.view.dispatch(closeHistory(editor.state.tr));
    expect(action.context.commands.replaceText('replacement')).toEqual({
      applied: true,
    });
    expect(editor.getText()).toContain('Before new replacement after.');

    expect(editor.commands.undo()).toBe(true);
    expect(editor.getText()).toContain('Before new target text after.');
    action.dispose();
  });

  test('rejects an asynchronous replacement after the selected text changes', () => {
    editor = createEditor();
    const range = textRange(editor, 'target text');
    editor.commands.setTextSelection(range);
    const action = createWorkDocumentSelectionAction(
      editor,
      requiredSnapshot(editor),
      () => false,
    );

    expect(editor.commands.insertContentAt(range, 'changed text')).toBe(true);
    expect(action.context.commands.replaceText('replacement')).toEqual({
      applied: false,
      reason: 'stale-selection',
    });
    expect(editor.getText()).toContain('Before changed text after.');
    action.dispose();
  });

  test('inserts text on either side without losing the live target', () => {
    editor = createEditor();
    editor.commands.setTextSelection(textRange(editor, 'target text'));
    const action = createWorkDocumentSelectionAction(
      editor,
      requiredSnapshot(editor),
      () => false,
    );
    let updates = 0;
    editor.on('update', () => {
      updates += 1;
    });

    expect(action.context.commands.insertBefore('expanded ')).toEqual({
      applied: true,
    });
    expect(action.context.commands.insertAfter(' polished')).toEqual({
      applied: true,
    });
    expect(editor.getText()).toContain(
      'Before expanded target text polished after.',
    );
    expect(updates).toBe(2);
    action.dispose();
  });

  test('keeps host replacement in one undo step and honors tracked changes', () => {
    let sequence = 0;
    editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
        createChange: (kind) => ({
          id: `${kind}-${++sequence}`,
          author: 'Host integration',
          date: '2026-07-26T00:00:00.000Z',
        }),
      }),
      content: content.html,
    });
    editor.commands.setTextSelection(textRange(editor, 'target text'));
    const action = createWorkDocumentSelectionAction(
      editor,
      requiredSnapshot(editor),
      () => true,
    );

    expect(action.context.commands.replaceText('polished text')).toEqual({
      applied: true,
    });
    expect(collectDocumentChanges(editor.state.doc)).toHaveLength(2);
    expect(editor.commands.undo()).toBe(true);
    expect(editor.getText()).toContain('Before target text after.');
    expect(collectDocumentChanges(editor.state.doc)).toEqual([]);
    action.dispose();
  });

  test('does not allow edit commands in read-only mode', () => {
    editor = createEditor();
    editor.commands.setTextSelection(textRange(editor, 'target text'));
    const action = createWorkDocumentSelectionAction(
      editor,
      requiredSnapshot(editor),
      () => false,
    );
    editor.setEditable(false);

    expect(action.context.commands.insertAfter('More text')).toEqual({
      applied: false,
      reason: 'read-only',
    });
    action.dispose();
  });
});

function createEditor(): Editor {
  return new Editor({
    extensions: createWorkDocumentExtensions(),
    content: content.html,
  });
}

function requiredSnapshot(editor: Editor) {
  const snapshot = createWorkDocumentSelectionSnapshot(editor, content);
  if (!snapshot) throw new Error('Expected a document selection snapshot.');
  return snapshot;
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
  if (!range) throw new Error(`Unable to find "${text}" in the document.`);
  return range;
}
