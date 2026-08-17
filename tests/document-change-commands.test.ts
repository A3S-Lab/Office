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

  test('records one character-formatting revision and restores mixed formatting on reject', () => {
    let sequence = 0;
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
        createChange: (kind: WorkDocumentChangeKind) => ({
          id: `${kind}-${++sequence}`,
          author: 'Reviewer',
          date: '2026-08-17T14:00:00.000Z',
        }),
      }),
      content:
        '<section data-document-section="true"><p><em>Alpha</em> beta</p></section>',
    });
    const alpha = textRange(editor, 'Alpha');
    const beta = textRange(editor, 'beta');
    const range = { from: alpha.from, to: beta.to };

    expect(editor.chain().setTextSelection(range).toggleBold().run()).toBe(
      true,
    );
    expect(collectDocumentChanges(editor.state.doc)).toEqual([
      expect.objectContaining({
        author: 'Reviewer',
        kind: 'formatting',
        text: 'Alpha beta',
      }),
    ]);
    expect(editor.getHTML()).toContain('data-change-kind="formatting"');
    expect(editor.getHTML()).toContain('<strong>');

    const change = collectDocumentChanges(editor.state.doc)[0];
    if (!change) throw new Error('Expected a formatting revision.');
    expect(editor.commands.rejectDocumentChange(change.id)).toBe(true);
    expect(collectDocumentChanges(editor.state.doc)).toEqual([]);
    expect(editor.getHTML()).toContain('<em>Alpha</em> beta');
    expect(editor.getHTML()).not.toContain('<strong>');

    editor.destroy();
  });

  test('keeps accepted formatting and groups revision metadata into one undo step', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
        createChange: () => ({
          id: 'formatting-bold',
          author: 'Reviewer',
          date: '2026-08-17T14:05:00.000Z',
        }),
      }),
      content:
        '<section data-document-section="true"><p>Tracked text</p></section>',
    });
    const range = textRange(editor, 'Tracked');

    expect(editor.chain().setTextSelection(range).toggleBold().run()).toBe(
      true,
    );
    expect(collectDocumentChanges(editor.state.doc)).toHaveLength(1);
    expect(editor.commands.undo()).toBe(true);
    expect(collectDocumentChanges(editor.state.doc)).toEqual([]);
    expect(editor.getHTML()).not.toContain('<strong>');
    expect(editor.commands.redo()).toBe(true);
    expect(collectDocumentChanges(editor.state.doc)).toHaveLength(1);

    expect(editor.commands.acceptDocumentChange('formatting-bold')).toBe(true);
    expect(collectDocumentChanges(editor.state.doc)).toEqual([]);
    expect(editor.getHTML()).toContain('<strong>Tracked</strong>');

    editor.destroy();
  });

  test('treats formatting on inserted text as part of the insertion revision', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({ isTracking: () => true }),
      content:
        '<section data-document-section="true"><p><ins data-document-change="true" data-change-kind="insertion" data-change-id="inserted" data-change-author="Reviewer" data-change-date="2026-08-17T14:10:00.000Z">Draft</ins></p></section>',
    });
    const range = textRange(editor, 'Draft');

    expect(editor.chain().setTextSelection(range).toggleItalic().run()).toBe(
      true,
    );
    expect(collectDocumentChanges(editor.state.doc)).toEqual([
      expect.objectContaining({ id: 'inserted', kind: 'insertion' }),
    ]);
    expect(editor.getHTML()).toContain('<em>Draft</em>');

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
