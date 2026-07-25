import { Editor } from '@tiptap/core';
import { describe, expect, test } from '@rstest/core';
import {
  collectDocumentChanges,
  type WorkDocumentChangeKind,
} from '../src/internal/features/work/work-document-changes';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';

describe('document change commands', () => {
  test('owns the controlled tracking setting in the change extension', () => {
    let tracking = false;
    const changes: boolean[] = [];
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => tracking,
        onTrackingChange: (enabled) => {
          tracking = enabled;
          changes.push(enabled);
        },
      }),
      content:
        '<section data-document-section="true"><p>Tracked text</p></section>',
    });

    expect(editor.commands.toggleDocumentTrackChanges()).toBe(true);
    expect(tracking).toBe(true);
    expect(editor.commands.setDocumentTrackChanges(false)).toBe(true);
    expect(changes).toEqual([true, false]);

    editor.destroy();
  });

  test('creates and resolves tracked replacements through TipTap commands', () => {
    let sequence = 0;
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
        createChange: (kind: WorkDocumentChangeKind) => ({
          id: `${kind}-${++sequence}`,
          author: 'Reviewer',
          date: '2026-07-25T00:00:00.000Z',
        }),
      }),
      content:
        '<section data-document-section="true"><p>Alpha beta</p></section>',
    });
    const range = textRange(editor, 'Alpha');

    expect(
      editor.commands.replaceDocumentTextWithTrackedChange(
        range.from,
        range.to,
        'Omega',
      ),
    ).toBe(true);
    expect(collectDocumentChanges(editor.state.doc)).toHaveLength(2);

    expect(editor.commands.acceptAllDocumentChanges()).toBe(true);
    expect(collectDocumentChanges(editor.state.doc)).toEqual([]);
    expect(editor.getText()).toContain('Omega beta');
    expect(editor.getText()).not.toContain('Alpha');

    editor.destroy();
  });
});

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
