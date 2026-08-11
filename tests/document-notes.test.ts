import { describe, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import { closeHistory } from '@tiptap/pm/history';
import JSZip from 'jszip';
import {
  createArtifact,
  createArtifactBlob,
  importOfficeFile,
} from '../src/core';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import { normalizeDocumentNotesHtml } from '../src/internal/features/work/work-document-notes';

describe('document notes', () => {
  test('renumbers notes by reference order and removes a deleted pair atomically', () => {
    const editor = createEditor();
    try {
      insertFootnote(editor, 'Beta', 'Later note');
      insertFootnote(editor, 'Alpha', 'Earlier note');

      const before = noteGraph(editor);
      expect(before.references.map(({ number }) => number)).toEqual([1, 2]);
      expect(definitionNumbersById(before)).toEqual(
        new Map(before.references.map(({ id, number }) => [id, number])),
      );

      const first = before.references[0];
      if (!first) throw new Error('Expected the first note reference.');
      editor.view.dispatch(
        editor.state.tr.delete(first.position, first.position + first.nodeSize),
      );

      const afterDelete = noteGraph(editor);
      expect(afterDelete.references).toHaveLength(1);
      expect(afterDelete.references[0]).toMatchObject({ number: 1 });
      expect(afterDelete.definitions).toEqual([
        expect.objectContaining({
          id: afterDelete.references[0]?.id,
          number: 1,
          text: 'Later note',
        }),
      ]);

      expect(editor.commands.undo()).toBe(true);
      expect(noteGraph(editor).references.map(({ number }) => number)).toEqual([
        1, 2,
      ]);
      expect(noteGraph(editor).definitions).toHaveLength(2);
    } finally {
      editor.destroy();
    }
  });

  test('gives a copied reference an independent definition and identity', () => {
    const editor = createEditor();
    try {
      insertFootnote(editor, 'Alpha', 'Shared note text');
      const original = noteGraph(editor).references[0];
      if (!original) throw new Error('Expected an original note reference.');
      const copy = editor.state.doc.slice(
        original.position,
        original.position + original.nodeSize,
        false,
      );
      editor.view.dispatch(
        closeHistory(editor.state.tr).insert(
          textRange(editor, 'Beta').to,
          copy.content,
        ),
      );

      const copied = noteGraph(editor);
      expect(copied.references).toHaveLength(2);
      expect(new Set(copied.references.map(({ id }) => id)).size).toBe(2);
      expect(copied.references[0]?.id).toBe(original.id);
      expect(copied.references.map(({ number }) => number)).toEqual([1, 2]);
      expect(copied.definitions).toHaveLength(2);
      expect(copied.definitions.map(({ text }) => text)).toEqual([
        'Shared note text',
        'Shared note text',
      ]);

      expect(editor.commands.undo()).toBe(true);
      expect(noteGraph(editor).references).toEqual([
        expect.objectContaining({ id: original.id, number: 1 }),
      ]);
      expect(noteGraph(editor).definitions).toEqual([
        expect.objectContaining({ id: original.id, number: 1 }),
      ]);
      expect(editor.commands.redo()).toBe(true);
      expect(
        new Set(noteGraph(editor).references.map(({ id }) => id)).size,
      ).toBe(2);
    } finally {
      editor.destroy();
    }
  });

  test('removes a reference when its complete definition is deleted', () => {
    const editor = createEditor();
    try {
      insertFootnote(editor, 'Alpha', 'Disposable note');
      const definition = noteGraph(editor).definitions[0];
      if (!definition) throw new Error('Expected a note definition.');
      editor.view.dispatch(
        closeHistory(editor.state.tr).delete(
          definition.position,
          definition.position + definition.nodeSize,
        ),
      );

      expect(noteGraph(editor)).toMatchObject({
        references: [],
        definitions: [],
      });
      expect(editor.commands.undo()).toBe(true);
      expect(noteGraph(editor).references).toHaveLength(1);
      expect(noteGraph(editor).definitions).toHaveLength(1);
    } finally {
      editor.destroy();
    }
  });

  test('places endnote definitions in the final document section', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: [
        '<section data-document-section="true" data-section-id="section-1">',
        '<p>Alpha</p></section>',
        '<section data-document-section="true" data-section-id="section-2">',
        '<p>Omega</p></section>',
      ].join(''),
    });
    try {
      editor.commands.setTextSelection(textRange(editor, 'Alpha').to);
      expect(editor.commands.insertDocumentNote('endnote')).toBe(true);
      expect(editor.commands.insertContent('Final note')).toBe(true);

      expect(noteGraph(editor).references).toEqual([
        expect.objectContaining({
          kind: 'endnote',
          number: 1,
          sectionId: 'section-1',
        }),
      ]);
      expect(noteGraph(editor).definitions).toEqual([
        expect.objectContaining({
          kind: 'endnote',
          number: 1,
          sectionId: 'section-2',
          text: 'Final note',
        }),
      ]);
    } finally {
      editor.destroy();
    }
  });

  test('removes a reference nested inside note content', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: [
        '<section data-document-section="true" data-section-id="section-1">',
        '<p>Alpha<sup data-document-note-reference="true" data-note-kind="footnote" data-note-id="outer">1</sup></p>',
        '<aside data-document-note="true" data-note-kind="footnote" data-note-id="outer">',
        '<p>Note text<sup data-document-note-reference="true" data-note-kind="footnote" data-note-id="nested">1</sup></p>',
        '</aside></section>',
      ].join(''),
    });
    try {
      expect(noteGraph(editor).references).toEqual([
        expect.objectContaining({ id: 'outer', number: 1 }),
      ]);
      expect(noteGraph(editor).definitions).toEqual([
        expect.objectContaining({ id: 'outer', number: 1, text: 'Note text' }),
      ]);
    } finally {
      editor.destroy();
    }
  });

  test('round-trips independently numbered footnotes and endnotes in native DOCX', async () => {
    const editor = createEditor();
    try {
      insertFootnote(editor, 'Beta', 'Later footnote');
      insertFootnote(editor, 'Alpha', 'Earlier footnote');
      editor.commands.setTextSelection(textRange(editor, 'Alpha').to);
      expect(editor.commands.insertDocumentNote('endnote')).toBe(true);
      expect(editor.commands.insertContent('Document endnote')).toBe(true);

      const artifact = createArtifact('blank-document');
      if (artifact.content.type !== 'document') {
        throw new Error('Expected a document artifact.');
      }
      artifact.content.html = editor.getHTML();
      const first = await createArtifactBlob(artifact);
      await assertNativeNotes(first);

      const imported = await importOfficeFile(
        new File([first], 'notes.docx', { type: first.type }),
      );
      if (imported.content.type !== 'document') {
        throw new Error('Expected an imported document artifact.');
      }
      const document = new DOMParser().parseFromString(
        imported.content.html,
        'text/html',
      );
      expect(
        document.body.querySelectorAll(
          'sup[data-document-note-reference][data-note-kind="footnote"]',
        ),
      ).toHaveLength(2);
      expect(
        document.body.querySelectorAll(
          'sup[data-document-note-reference][data-note-kind="endnote"]',
        ),
      ).toHaveLength(1);
      expect(document.body.textContent).toContain('Earlier footnote');
      expect(document.body.textContent).toContain('Later footnote');
      expect(document.body.textContent).toContain('Document endnote');
      expect(imported.compatibility.issues).toContainEqual(
        expect.objectContaining({ code: 'docx.notes', severity: 'info' }),
      );

      await assertNativeNotes(await createArtifactBlob(imported));
    } finally {
      editor.destroy();
    }
  });

  test('normalizes duplicated controlled HTML references into independent pairs', () => {
    const normalized = normalizeDocumentNotesHtml(
      [
        '<section data-document-section="true">',
        '<p>Alpha<sup data-document-note-reference="true" data-note-kind="footnote" data-note-id="shared">1</sup></p>',
        '<p>Beta<sup data-document-note-reference="true" data-note-kind="footnote" data-note-id="shared">1</sup></p>',
        '<aside data-document-note="true" data-note-kind="footnote" data-note-id="shared"><p>Shared text</p></aside>',
        '</section>',
      ].join(''),
    );
    const document = new DOMParser().parseFromString(normalized, 'text/html');
    const references = Array.from(
      document.body.querySelectorAll<HTMLElement>(
        'sup[data-document-note-reference]',
      ),
    );
    const definitions = Array.from(
      document.body.querySelectorAll<HTMLElement>('aside[data-document-note]'),
    );

    expect(references.map(({ dataset }) => dataset.noteNumber)).toEqual([
      '1',
      '2',
    ]);
    expect(new Set(references.map(({ dataset }) => dataset.noteId)).size).toBe(
      2,
    );
    expect(definitions.map(({ dataset }) => dataset.noteId)).toEqual(
      references.map(({ dataset }) => dataset.noteId),
    );
    expect(definitions.map(({ textContent }) => textContent)).toEqual([
      'Shared text',
      'Shared text',
    ]);
  });
});

interface NoteAtPosition {
  id: string;
  kind: string;
  number: number;
  position: number;
  nodeSize: number;
  sectionId: string;
  text: string;
}

function createEditor(): Editor {
  return new Editor({
    extensions: createWorkDocumentExtensions(),
    content: [
      '<section data-document-section="true" data-section-id="section-1">',
      '<p>Alpha</p><p>Beta</p>',
      '</section>',
    ].join(''),
  });
}

function insertFootnote(editor: Editor, target: string, text: string): void {
  editor.commands.setTextSelection(textRange(editor, target).to);
  expect(editor.commands.insertDocumentNote('footnote')).toBe(true);
  expect(editor.commands.insertContent(text)).toBe(true);
}

function noteGraph(editor: Editor): {
  references: NoteAtPosition[];
  definitions: NoteAtPosition[];
} {
  const references: NoteAtPosition[] = [];
  const definitions: NoteAtPosition[] = [];
  editor.state.doc.descendants((node, position) => {
    if (
      node.type.name !== 'documentNoteReference' &&
      node.type.name !== 'documentNote'
    ) {
      return;
    }
    const item = {
      id: String(node.attrs.id),
      kind: String(node.attrs.kind),
      number: Number(node.attrs.number),
      position,
      nodeSize: node.nodeSize,
      sectionId: sectionIdAt(editor, position),
      text: node.textContent,
    };
    if (node.type.name === 'documentNoteReference') references.push(item);
    else definitions.push(item);
  });
  return { references, definitions };
}

function sectionIdAt(editor: Editor, position: number): string {
  const resolved = editor.state.doc.resolve(position);
  for (let depth = resolved.depth; depth > 0; depth -= 1) {
    const node = resolved.node(depth);
    if (node.type.name === 'documentSection') return String(node.attrs.id);
  }
  return '';
}

function definitionNumbersById(graph: ReturnType<typeof noteGraph>) {
  return new Map(graph.definitions.map(({ id, number }) => [id, number]));
}

async function assertNativeNotes(blob: Blob): Promise<void> {
  const archive = await JSZip.loadAsync(await blob.arrayBuffer());
  const [documentXml = '', footnotesXml = '', endnotesXml = ''] =
    await Promise.all([
      archive.file('word/document.xml')?.async('string'),
      archive.file('word/footnotes.xml')?.async('string'),
      archive.file('word/endnotes.xml')?.async('string'),
    ]);
  expect(
    Array.from(
      documentXml.matchAll(
        /<w:footnoteReference\b(?=[^>]*w:id="(\d+)")[^>]*>/g,
      ),
      ([, id]) => id,
    ),
  ).toEqual(['1', '2']);
  expect(documentXml).toMatch(/<w:endnoteReference\b(?=[^>]*w:id="1")[^>]*>/);
  expect(noteXmlContaining(footnotesXml, 'Earlier footnote')).toContain(
    'w:id="1"',
  );
  expect(noteXmlContaining(footnotesXml, 'Later footnote')).toContain(
    'w:id="2"',
  );
  expect(noteXmlContaining(endnotesXml, 'Document endnote')).toContain(
    'w:id="1"',
  );
}

function noteXmlContaining(source: string, text: string): string {
  return (
    Array.from(
      source.matchAll(
        /<w:(?:footnote|endnote)\b[\s\S]*?<\/w:(?:footnote|endnote)>/g,
      ),
      ([note]) => note,
    ).find((note) => note.includes(text)) ?? ''
  );
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
  if (!range) throw new Error(`Unable to find "${text}".`);
  return range;
}
