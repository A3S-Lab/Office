import { Editor } from '@tiptap/core';
import { describe, expect, test } from '@rstest/core';
import { editorDocumentCaptionTargets } from '../src/internal/features/work/work-document-caption-nodes';
import { documentCitationCount } from '../src/internal/features/work/work-document-citation-nodes';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import type {
  WorkDocumentBibliography,
  WorkDocumentCitationSource,
  WorkDocumentContent,
} from '../src/internal/features/work/work-types';

const source: WorkDocumentCitationSource = {
  id: 'source-1',
  tag: 'doe2026',
  sourceType: 'Book',
  title: 'Modular Editors',
  year: '2026',
  contributors: {
    Author: {
      people: [{ first: 'Jane', last: 'Doe' }],
    },
  },
};

const bibliography: WorkDocumentBibliography = {
  style: 'apa',
  sources: [source],
};

describe('document reference commands', () => {
  test('owns controlled bibliography updates in the citation extension', () => {
    let content: WorkDocumentContent = {
      type: 'document',
      pageSize: 'a4',
      html: '<section data-document-section="true"><p>References</p></section>',
    };
    const updates: WorkDocumentContent[] = [];
    const editor = new Editor({
      extensions: createWorkDocumentExtensions({
        getContent: () => content,
        onContentChange: (next) => {
          content = next;
          updates.push(next);
        },
      }),
      content: content.html,
    });

    expect(editor.can().setDocumentBibliography(bibliography)).toBe(true);
    expect(updates).toEqual([]);
    expect(editor.commands.setDocumentBibliography(bibliography)).toBe(true);
    expect(content.bibliography).toEqual(bibliography);
    expect(updates).toHaveLength(1);

    editor.destroy();
  });

  test('owns captions, notes, fields, and citations in TipTap extensions', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: [
        '<section data-document-section="true" data-section-id="section-1">',
        '<p>Alpha</p><p>Beta</p>',
        '</section>',
      ].join(''),
    });

    try {
      editor.commands.setTextSelection(textRange(editor, 'Alpha').to);
      const beforeCan = editor.getJSON();
      expect(editor.can().insertDocumentCaption('figure', 'Architecture')).toBe(
        true,
      );
      expect(editor.getJSON()).toEqual(beforeCan);
      expect(
        editor.commands.insertDocumentCaption('figure', 'Architecture'),
      ).toBe(true);

      const targets = editorDocumentCaptionTargets(editor);
      expect(targets).toHaveLength(1);
      expect(targets[0]).toMatchObject({
        kind: 'figure',
        number: 1,
        title: 'Architecture',
      });

      editor.commands.setTextSelection(textRange(editor, 'Beta').to);
      expect(editor.commands.insertDocumentCrossReference(targets[0])).toBe(
        true,
      );
      expect(nodeCount(editor, 'documentCrossReference')).toBe(1);

      editor.commands.setTextSelection(textRange(editor, 'Alpha').to);
      expect(editor.commands.insertDocumentNote('footnote')).toBe(true);
      expect(nodeCount(editor, 'documentNoteReference')).toBe(1);
      expect(nodeCount(editor, 'documentNote')).toBe(1);

      editor.commands.setTextSelection(textRange(editor, 'Alpha').to);
      expect(editor.commands.insertDocumentField('page')).toBe(true);
      expect(nodeCount(editor, 'documentField')).toBe(1);
      expect(
        editor.commands.refreshDocumentFields(
          documentContent(editor, bibliography),
        ),
      ).toBe(false);

      editor.commands.setTextSelection(textRange(editor, 'Beta').from);
      expect(editor.commands.insertDocumentCitation(source, bibliography)).toBe(
        true,
      );
      expect(documentCitationCount(editor)).toBe(1);

      expect(
        editor.commands.renameDocumentCitationTag('doe2026', 'doe2027'),
      ).toBe(true);
      const renamedSource = { ...source, tag: 'doe2027' };
      const renamedBibliography = {
        ...bibliography,
        sources: [renamedSource],
      };
      expect(
        editor.commands.refreshDocumentCitations(
          documentContent(editor, renamedBibliography),
        ),
      ).toBe(false);
      expect(editor.getHTML()).toContain('data-citation-tags="doe2027"');
      expect(editor.getHTML()).not.toContain('缺失引文');

      editor.commands.setTextSelection(textRange(editor, 'Beta').to);
      expect(
        editor.commands.insertDocumentBibliography(renamedBibliography),
      ).toBe(true);
      expect(nodeCount(editor, 'documentBibliography')).toBe(1);
    } finally {
      editor.destroy();
    }
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

function nodeCount(editor: Editor, type: string): number {
  let count = 0;
  editor.state.doc.descendants((node) => {
    if (node.type.name === type) count += 1;
  });
  return count;
}

function documentContent(
  editor: Editor,
  value: WorkDocumentBibliography,
): WorkDocumentContent {
  return {
    type: 'document',
    pageSize: 'a4',
    html: editor.getHTML(),
    bibliography: value,
  };
}
