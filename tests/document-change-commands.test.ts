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

  test('records paragraph formatting and restores the exact previous attributes on reject', () => {
    let sequence = 0;
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
        createChange: (kind: WorkDocumentChangeKind) => ({
          id: `${kind}-${++sequence}`,
          actorId: 'reviewer-1',
          author: 'Reviewer',
          date: '2026-08-18T09:00:00.000Z',
        }),
      }),
      content:
        '<section data-document-section="true"><p data-office-space-before="6" data-office-indent-right="12">Paragraph review</p></section>',
    });
    const range = textRange(editor, 'Paragraph review');

    expect(
      editor
        .chain()
        .setTextSelection(range)
        .setTextAlign('center')
        .setDocumentParagraphSpacing({
          before: 18,
          after: 9,
          lineHeight: '1.5',
          lineRule: 'auto',
        })
        .setDocumentParagraphIndent(
          { left: 48, right: 24, firstLine: -12 },
          { restoreFocus: false },
        )
        .run(),
    ).toBe(true);

    const changes = collectDocumentChanges(editor.state.doc);
    expect(changes).toEqual([
      expect.objectContaining({
        id: 'paragraph-formatting-1',
        actorId: 'reviewer-1',
        kind: 'paragraph-formatting',
        text: 'Paragraph review',
      }),
    ]);
    expect(editor.getHTML()).toContain(
      'data-change-kind="paragraph-formatting"',
    );
    expect(paragraphAttributes(editor)).toEqual(
      expect.objectContaining({
        textAlign: 'center',
        indentLevel: 2,
        rightIndent: 24,
        firstLineIndent: -12,
        spaceBefore: 18,
        spaceAfter: 9,
        lineHeight: '1.5',
        lineRule: 'auto',
      }),
    );

    expect(editor.commands.rejectDocumentChange(changes[0]?.id ?? '')).toBe(
      true,
    );
    expect(collectDocumentChanges(editor.state.doc)).toEqual([]);
    expect(paragraphAttributes(editor)).toEqual(
      expect.objectContaining({
        textAlign: null,
        indentLevel: 0,
        rightIndent: 12,
        firstLineIndent: 0,
        spaceBefore: 6,
        spaceAfter: null,
        lineHeight: null,
        lineRule: null,
      }),
    );
    expect(editor.getText()).toContain('Paragraph review');

    editor.destroy();
  });

  test('keeps accepted paragraph formatting and groups tracking into one undo record', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
        createChange: () => ({
          id: 'paragraph-alignment',
          author: 'Reviewer',
          date: '2026-08-18T09:05:00.000Z',
        }),
      }),
      content:
        '<section data-document-section="true"><p>Tracked paragraph</p></section>',
    });
    const range = textRange(editor, 'Tracked paragraph');

    expect(
      editor.chain().setTextSelection(range).setTextAlign('right').run(),
    ).toBe(true);
    expect(collectDocumentChanges(editor.state.doc)).toHaveLength(1);
    expect(editor.commands.undo()).toBe(true);
    expect(collectDocumentChanges(editor.state.doc)).toEqual([]);
    expect(paragraphAttributes(editor).textAlign).toBeNull();
    expect(editor.commands.redo()).toBe(true);
    expect(collectDocumentChanges(editor.state.doc)).toHaveLength(1);

    expect(editor.commands.acceptDocumentChange('paragraph-alignment')).toBe(
      true,
    );
    expect(collectDocumentChanges(editor.state.doc)).toEqual([]);
    expect(paragraphAttributes(editor).textAlign).toBe('right');

    editor.destroy();
  });

  test('retains the original paragraph snapshot and identity across later edits', () => {
    let sequence = 0;
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
        createChange: (kind: WorkDocumentChangeKind) => ({
          id: `${kind}-${++sequence}`,
          author: 'Reviewer',
          date: '2026-08-18T09:10:00.000Z',
        }),
      }),
      content:
        '<section data-document-section="true"><p>Iterated paragraph</p></section>',
    });
    const range = textRange(editor, 'Iterated paragraph');

    editor.chain().setTextSelection(range).setTextAlign('center').run();
    editor.chain().setTextSelection(range).setTextAlign('justify').run();

    expect(collectDocumentChanges(editor.state.doc)).toEqual([
      expect.objectContaining({
        id: 'paragraph-formatting-1',
        kind: 'paragraph-formatting',
      }),
    ]);
    expect(editor.commands.rejectDocumentChange('paragraph-formatting-1')).toBe(
      true,
    );
    expect(paragraphAttributes(editor).textAlign).toBeNull();

    editor.destroy();
  });

  test('resolves a shared paragraph revision across multiple paragraphs atomically', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        isTracking: () => true,
        createChange: () => ({
          id: 'multi-paragraph',
          author: 'Reviewer',
          date: '2026-08-18T09:15:00.000Z',
        }),
      }),
      content:
        '<section data-document-section="true"><p>First paragraph</p><p>Second paragraph</p></section>',
    });
    const first = textRange(editor, 'First paragraph');
    const second = textRange(editor, 'Second paragraph');

    expect(
      editor
        .chain()
        .setTextSelection({ from: first.from, to: second.to })
        .setTextAlign('center')
        .run(),
    ).toBe(true);
    expect(collectDocumentChanges(editor.state.doc)).toEqual([
      expect.objectContaining({
        id: 'multi-paragraph',
        kind: 'paragraph-formatting',
        text: 'First paragraph\nSecond paragraph',
      }),
    ]);
    expect(editor.commands.rejectDocumentChange('multi-paragraph')).toBe(true);
    expect(collectDocumentChanges(editor.state.doc)).toEqual([]);
    expect(paragraphAttributeValues(editor, 'textAlign')).toEqual([null, null]);
    expect(editor.commands.undo()).toBe(true);
    expect(paragraphAttributeValues(editor, 'textAlign')).toEqual([
      'center',
      'center',
    ]);

    editor.destroy();
  });

  test('fails closed when a paragraph revision snapshot is malformed', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({ isTracking: () => false }),
      content:
        '<section data-document-section="true"><p data-document-change="true" data-change-kind="paragraph-formatting" data-change-id="malformed" data-change-author="Reviewer" data-change-date="2026-08-18T09:20:00.000Z" data-change-before="{}" style="text-align: center">Malformed snapshot</p></section>',
    });

    expect(collectDocumentChanges(editor.state.doc)).toEqual([
      expect.objectContaining({
        id: 'malformed',
        kind: 'paragraph-formatting',
      }),
    ]);
    expect(editor.commands.rejectDocumentChange('malformed')).toBe(false);
    expect(editor.commands.acceptDocumentChange('malformed')).toBe(false);
    expect(collectDocumentChanges(editor.state.doc)).toHaveLength(1);
    expect(paragraphAttributes(editor).textAlign).toBe('center');

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

function paragraphAttributes(editor: Editor): Record<string, unknown> {
  let attributes: Record<string, unknown> | null = null;
  editor.state.doc.descendants((node) => {
    if (!attributes && node.type.name === 'paragraph') attributes = node.attrs;
  });
  if (!attributes) throw new Error('Unable to find a paragraph.');
  return attributes;
}

function paragraphAttributeValues(
  editor: Editor,
  attribute: string,
): unknown[] {
  const values: unknown[] = [];
  editor.state.doc.descendants((node) => {
    if (node.type.name === 'paragraph') values.push(node.attrs[attribute]);
  });
  return values;
}
