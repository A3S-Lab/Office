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

  test('keeps caption numbering and cross-reference state live after deletion', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: [
        '<section data-document-section="true" data-section-id="section-1">',
        '<p>Alpha</p><p>Beta</p><p>References</p>',
        '</section>',
      ].join(''),
    });

    try {
      editor.commands.setTextSelection(textRange(editor, 'Alpha').to);
      expect(
        editor.commands.insertDocumentCaption('figure', 'Architecture'),
      ).toBe(true);
      editor.commands.setTextSelection(textRange(editor, 'Beta').to);
      expect(editor.commands.insertDocumentCaption('figure', 'Data flow')).toBe(
        true,
      );

      const [architecture, dataFlow] = editorDocumentCaptionTargets(editor);
      expect(architecture).toMatchObject({ number: 1, title: 'Architecture' });
      expect(dataFlow).toMatchObject({ number: 2, title: 'Data flow' });
      if (!architecture || !dataFlow)
        throw new Error('Expected both caption targets.');
      expect(
        editor.view.dom
          .querySelector<HTMLElement>(`[data-caption-id="${architecture.id}"]`)
          ?.getAttribute('aria-label'),
      ).toBe('图 1 Architecture');

      editor.commands.setTextSelection(textRange(editor, 'References').to);
      expect(editor.commands.insertDocumentCrossReference(architecture)).toBe(
        true,
      );
      expect(editor.commands.insertDocumentCrossReference(dataFlow)).toBe(true);

      const firstCaption = nodePosition(
        editor,
        'documentCaption',
        architecture.id,
      );
      editor.view.dispatch(
        editor.state.tr.delete(
          firstCaption.position,
          firstCaption.position + firstCaption.nodeSize,
        ),
      );

      expect(editorDocumentCaptionTargets(editor)).toEqual([
        expect.objectContaining({
          id: dataFlow.id,
          number: 1,
          title: 'Data flow',
        }),
      ]);
      expect(
        editor.view.dom
          .querySelector<HTMLElement>(`[data-caption-id="${dataFlow.id}"]`)
          ?.getAttribute('aria-label'),
      ).toBe('图 1 Data flow');
      expect(referenceStates(editor)).toEqual([
        {
          targetId: architecture.id,
          number: 1,
          orphaned: true,
          text: '引用缺失',
        },
        {
          targetId: dataFlow.id,
          number: 1,
          orphaned: false,
          text: '图 1',
        },
      ]);

      expect(editor.commands.undo()).toBe(true);
      expect(
        editorDocumentCaptionTargets(editor).map(({ number }) => number),
      ).toEqual([1, 2]);
      expect(referenceStates(editor)).toEqual([
        {
          targetId: architecture.id,
          number: 1,
          orphaned: false,
          text: '图 1',
        },
        {
          targetId: dataFlow.id,
          number: 2,
          orphaned: false,
          text: '图 2',
        },
      ]);

      const restoredCaption = nodePosition(
        editor,
        'documentCaption',
        architecture.id,
      );
      editor.commands.setTextSelection({
        from: restoredCaption.position + 1,
        to: restoredCaption.position + restoredCaption.nodeSize - 1,
      });
      expect(editor.commands.deleteSelection()).toBe(true);
      expect(
        editor.view.dom
          .querySelector<HTMLElement>(`[data-caption-id="${architecture.id}"]`)
          ?.getAttribute('aria-label'),
      ).toBe('图 1');
      expect(editor.commands.keyboardShortcut('Backspace')).toBe(true);
      expect(editorDocumentCaptionTargets(editor)).toEqual([
        expect.objectContaining({ id: dataFlow.id, number: 1 }),
      ]);
      expect(referenceStates(editor)[0]).toMatchObject({
        targetId: architecture.id,
        orphaned: true,
        text: '引用缺失',
      });
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

function nodePosition(
  editor: Editor,
  type: string,
  id: string,
): { position: number; nodeSize: number } {
  let match: { position: number; nodeSize: number } | null = null;
  editor.state.doc.descendants((node, position) => {
    if (match || node.type.name !== type || node.attrs.id !== id) return;
    match = { position, nodeSize: node.nodeSize };
  });
  if (!match) throw new Error(`Unable to find ${type} ${id}.`);
  return match;
}

function referenceStates(editor: Editor): Array<{
  targetId: string;
  number: number;
  orphaned: boolean;
  text: string;
}> {
  const references: Array<{
    targetId: string;
    number: number;
    orphaned: boolean;
    text: string;
  }> = [];
  editor.state.doc.descendants((node) => {
    if (node.type.name !== 'documentCrossReference') return;
    references.push({
      targetId: String(node.attrs.targetId),
      number: Number(node.attrs.number),
      orphaned: Boolean(node.attrs.orphaned),
      text: node.attrs.orphaned ? '引用缺失' : `图 ${node.attrs.number}`,
    });
  });
  return references;
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
