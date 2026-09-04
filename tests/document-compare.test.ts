import { describe, expect, test } from '@rstest/core';
import { type Content, Editor } from '@tiptap/core';
import JSZip from 'jszip';
import {
  createArtifact,
  createArtifactBlob,
  importOfficeFile,
} from '../src/core';
import { collectDocumentChanges } from '../src/internal/features/work/work-document-changes';
import {
  applyDocumentComparison,
  type DocumentComparisonApplyResult,
} from '../src/internal/features/work/work-document-compare';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';

const comparisonIdentity = {
  author: 'Morgan',
  date: '2026-08-25T10:00:00.000Z',
  sourceName: 'reviewed-contract.docx',
};

describe('document compare and combine', () => {
  test('creates deterministic reviewable text and block revisions in one undo record', () => {
    const original = documentHtml('<p>Alpha beta.</p><p>Stable paragraph.</p>');
    const revised = documentHtml(
      '<p>Alpha gamma.</p><p>Inserted paragraph.</p><p>Stable paragraph.</p>',
    );
    const first = createEditor(original);
    const second = createEditor(original);
    const originalSnapshot = first.getHTML();

    const firstResult = applyCompare(first, revised);
    const secondResult = applyCompare(second, revised);

    expect(firstResult.status).toBe('applied');
    expect(secondResult.status).toBe('applied');
    expect(firstResult.summary).toEqual({
      deletions: 1,
      formatting: 0,
      insertions: 2,
      paragraphFormatting: 0,
    });
    expect(collectDocumentChanges(first.state.doc)).toEqual(
      collectDocumentChanges(second.state.doc),
    );
    expect(
      collectDocumentChanges(first.state.doc).map((change) => change.kind),
    ).toEqual(['deletion', 'insertion', 'insertion']);

    expect(first.commands.undo()).toBe(true);
    expect(first.getHTML()).toBe(originalSnapshot);
    expect(first.commands.undo()).toBe(false);
    expect(first.commands.redo()).toBe(true);
    expect(collectDocumentChanges(first.state.doc)).toHaveLength(3);

    expect(first.commands.rejectAllDocumentChanges()).toBe(true);
    expect(normalizedText(first)).toBe('Alpha beta. Stable paragraph.');
    expect(paragraphTexts(first)).toEqual(['Alpha beta.', 'Stable paragraph.']);

    expect(second.commands.acceptAllDocumentChanges()).toBe(true);
    expect(normalizedText(second)).toBe(
      'Alpha gamma. Inserted paragraph. Stable paragraph.',
    );
    expect(paragraphTexts(second)).toEqual([
      'Alpha gamma.',
      'Inserted paragraph.',
      'Stable paragraph.',
    ]);

    first.destroy();
    second.destroy();
  });

  test('collects structural block revision text exactly once', () => {
    const editor = createEditor(documentHtml('<p>Stable paragraph.</p>'));

    const result = applyCompare(
      editor,
      documentHtml('<p>Inserted paragraph.</p><p>Stable paragraph.</p>'),
    );

    expect(result.status).toBe('applied');
    expect(collectDocumentChanges(editor.state.doc)).toEqual([
      expect.objectContaining({
        kind: 'insertion',
        text: 'Inserted paragraph.',
      }),
    ]);

    editor.destroy();
  });

  test('records character and paragraph formatting differences with exact rejection', () => {
    const editor = createEditor(
      documentHtml('<p><strong>Formatted text</strong></p>'),
    );
    const originalHtml = editor.getHTML();

    const result = applyCompare(
      editor,
      documentHtml('<p style="text-align: center"><em>Formatted text</em></p>'),
    );

    expect(result.status).toBe('applied');
    expect(result.summary).toEqual({
      deletions: 0,
      formatting: 1,
      insertions: 0,
      paragraphFormatting: 1,
    });
    expect(collectDocumentChanges(editor.state.doc)).toEqual([
      expect.objectContaining({ kind: 'paragraph-formatting' }),
      expect.objectContaining({ kind: 'formatting', text: 'Formatted text' }),
    ]);
    expect(editor.getHTML()).toContain('<em>Formatted text</em>');
    expect(editor.getHTML()).toContain('text-align: center');

    expect(editor.commands.rejectAllDocumentChanges()).toBe(true);
    expect(editor.getHTML()).toBe(originalHtml);

    editor.destroy();
  });

  test('infers a bounded text move and round-trips native move revisions', async () => {
    const original = documentHtml('<p>Alpha beta gamma.</p>');
    const revised = documentHtml('<p>Alpha gamma beta.</p>');
    const editor = createEditor(original);
    const result = applyCompare(editor, revised);

    expect(result.status).toBe('applied');
    expect(result.summary).toEqual({
      deletions: 0,
      formatting: 0,
      insertions: 0,
      moves: 1,
      paragraphFormatting: 0,
    });
    const move = collectDocumentChanges(editor.state.doc).find(
      (change) => change.kind === 'move',
    );
    expect(move).toEqual(
      expect.objectContaining({
        author: 'Morgan',
        text: ' beta',
      }),
    );
    expect(
      collectDocumentChanges(editor.state.doc).filter(
        (change) => change.kind === 'move',
      ),
    ).toHaveLength(1);

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a Writer artifact.');
    }
    artifact.content.html = editor.getHTML();
    artifact.content.trackChanges = true;
    const exported = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await exported.arrayBuffer());
    const documentXml =
      (await archive.file('word/document.xml')?.async('text')) ?? '';
    expect(documentXml).toContain('<w:moveFrom');
    expect(documentXml).toContain('<w:moveTo');
    expect(documentXml).not.toContain('data-change-move-role');

    const reopened = await importOfficeFile(
      new File([exported], 'comparison-move.docx', {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened Writer artifact.');
    }
    const reopenedEditor = createEditor(reopened.content.html);
    expect(
      collectDocumentChanges(reopenedEditor.state.doc).filter(
        (change) => change.kind === 'move',
      ),
    ).toEqual([
      expect.objectContaining({
        author: 'Morgan',
        kind: 'move',
        text: ' beta',
      }),
    ]);

    expect(reopenedEditor.commands.rejectAllDocumentChanges()).toBe(true);
    expect(normalizedText(reopenedEditor)).toBe('Alpha beta gamma.');
    expect(reopenedEditor.commands.undo()).toBe(true);
    expect(reopenedEditor.commands.acceptAllDocumentChanges()).toBe(true);
    expect(normalizedText(reopenedEditor)).toBe('Alpha gamma beta.');

    editor.destroy();
    reopenedEditor.destroy();
  });

  test('infers a bounded move between simple paragraphs atomically', async () => {
    const original = documentHtml(
      '<p>Intro move phrase remains.</p><p>Destination tail remains.</p>',
    );
    const revised = documentHtml(
      '<p>Intro remains.</p><p>Destination move phrase tail remains.</p>',
    );
    const editor = createEditor(original);

    const result = applyCompare(editor, revised);

    expect(result.status).toBe('applied');
    expect(result.summary).toEqual({
      deletions: 0,
      formatting: 0,
      insertions: 0,
      moves: 1,
      paragraphFormatting: 0,
    });
    expect(collectDocumentChanges(editor.state.doc)).toEqual([
      expect.objectContaining({ kind: 'move', text: ' move phrase' }),
    ]);

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a Writer artifact.');
    }
    artifact.content.html = editor.getHTML();
    artifact.content.trackChanges = true;
    const exported = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await exported.arrayBuffer());
    const documentXml =
      (await archive.file('word/document.xml')?.async('text')) ?? '';
    expect(documentXml).toContain('<w:moveFrom');
    expect(documentXml).toContain('<w:moveTo');
    expect(documentXml).not.toContain('data-change-move-role');

    const reopened = await importOfficeFile(
      new File([exported], 'cross-paragraph-move.docx', {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened Writer artifact.');
    }
    const reopenedEditor = createEditor(reopened.content.html);
    expect(
      collectDocumentChanges(reopenedEditor.state.doc).filter(
        (change) => change.kind === 'move',
      ),
    ).toEqual([
      expect.objectContaining({
        author: 'Morgan',
        kind: 'move',
        text: ' move phrase',
      }),
    ]);

    expect(editor.commands.rejectAllDocumentChanges()).toBe(true);
    expect(paragraphTexts(editor)).toEqual([
      'Intro move phrase remains.',
      'Destination tail remains.',
    ]);
    expect(editor.commands.undo()).toBe(true);
    expect(editor.commands.acceptAllDocumentChanges()).toBe(true);
    expect(paragraphTexts(editor)).toEqual([
      'Intro remains.',
      'Destination move phrase tail remains.',
    ]);

    expect(reopenedEditor.commands.rejectAllDocumentChanges()).toBe(true);
    expect(paragraphTexts(reopenedEditor)).toEqual([
      'Intro move phrase remains.',
      'Destination tail remains.',
    ]);
    expect(reopenedEditor.commands.undo()).toBe(true);
    expect(reopenedEditor.commands.acceptAllDocumentChanges()).toBe(true);
    expect(paragraphTexts(reopenedEditor)).toEqual([
      'Intro remains.',
      'Destination move phrase tail remains.',
    ]);

    editor.destroy();
    reopenedEditor.destroy();
  });

  test('leaves ambiguous duplicate cross-paragraph ranges as ordinary revisions', () => {
    const original = documentHtml(
      '<p>Source A move phrase.</p><p>Source B move phrase.</p><p>Destination tail.</p>',
    );
    const revised = documentHtml(
      '<p>Source A.</p><p>Source B.</p><p>Destination move phrase tail.</p>',
    );
    const editor = createEditor(original);

    const result = applyCompare(editor, revised);

    expect(result.status).toBe('applied');
    expect(result.summary).toEqual({
      deletions: 2,
      formatting: 0,
      insertions: 1,
      paragraphFormatting: 0,
    });
    expect(
      collectDocumentChanges(editor.state.doc).some(
        (change) => change.kind === 'move',
      ),
    ).toBe(false);

    expect(editor.commands.rejectAllDocumentChanges()).toBe(true);
    expect(paragraphTexts(editor)).toEqual([
      'Source A move phrase.',
      'Source B move phrase.',
      'Destination tail.',
    ]);
    expect(editor.commands.undo()).toBe(true);
    expect(editor.commands.acceptAllDocumentChanges()).toBe(true);
    expect(paragraphTexts(editor)).toEqual([
      'Source A.',
      'Source B.',
      'Destination move phrase tail.',
    ]);

    editor.destroy();
  });

  test('does not pair move ranges across section boundaries', () => {
    const original =
      '<section data-document-section="true"><p>Intro move phrase remains.</p></section>' +
      '<section data-document-section="true"><p>Destination tail remains.</p></section>';
    const revised =
      '<section data-document-section="true"><p>Intro remains.</p></section>' +
      '<section data-document-section="true"><p>Destination move phrase tail remains.</p></section>';
    const editor = createEditor(original);

    const result = applyCompare(editor, revised);

    expect(result.status).toBe('applied');
    expect(result.summary).toEqual({
      deletions: 1,
      formatting: 0,
      insertions: 1,
      paragraphFormatting: 0,
    });
    expect(
      collectDocumentChanges(editor.state.doc).some(
        (change) => change.kind === 'move',
      ),
    ).toBe(false);

    editor.destroy();
  });

  test('fails closed for changed complex structures and leaves the document untouched', () => {
    const original = documentHtml(
      '<table><tbody><tr><td><p>Original cell</p></td></tr></tbody></table>',
    );
    const revised = documentHtml(
      '<table><tbody><tr><td><p>Revised cell</p></td></tr></tbody></table>',
    );
    const editor = createEditor(original);
    const revisedEditor = createEditor(revised);
    const originalSnapshot = editor.getHTML();

    const result = applyCompare(editor, revisedEditor.getJSON());

    expect(result.status).toBe('unsupported');
    expect(result.diagnostics).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ code: 'changed-complex-structure' }),
      ]),
    );
    expect(editor.getHTML()).toBe(originalSnapshot);
    expect(collectDocumentChanges(editor.state.doc)).toEqual([]);
    expect(editor.commands.undo()).toBe(false);

    editor.destroy();
    revisedEditor.destroy();
  });

  test('returns unchanged without publishing a transaction', () => {
    const original = documentHtml('<p>Identical content.</p>');
    const editor = createEditor(original);

    const result = applyCompare(editor, original);

    expect(result.status).toBe('unchanged');
    expect(result.summary).toEqual({
      deletions: 0,
      formatting: 0,
      insertions: 0,
      paragraphFormatting: 0,
    });
    expect(editor.commands.undo()).toBe(false);

    editor.destroy();
  });

  test('combines a reviewed copy only when its rejected baseline matches current content', () => {
    const original = documentHtml(
      '<p data-office-paragraph-id="00001001" data-office-paragraph-text-id="00002001">Alpha beta.</p>',
    );
    const reviewed = documentHtml(
      '<p data-office-paragraph-id="00001002" data-office-paragraph-text-id="00003002">Alpha <del data-document-change="true" data-change-kind="deletion" data-change-id="review-delete" data-change-author="Riley" data-change-date="2026-08-24T09:00:00.000Z">beta</del><ins data-document-change="true" data-change-kind="insertion" data-change-id="review-insert" data-change-author="Riley" data-change-date="2026-08-24T09:01:00.000Z">gamma</ins>.</p>',
    );
    const editor = createEditor(original);
    const originalSnapshot = editor.getHTML();

    const result = applyDocumentComparison(editor, reviewed, {
      ...comparisonIdentity,
      mode: 'combine',
    });

    expect(result.status).toBe('applied');
    expect(result.summary).toEqual({
      deletions: 1,
      formatting: 0,
      insertions: 1,
      paragraphFormatting: 0,
    });
    expect(collectDocumentChanges(editor.state.doc)).toEqual([
      expect.objectContaining({
        author: 'Riley',
        id: 'review-delete',
        kind: 'deletion',
      }),
      expect.objectContaining({
        author: 'Riley',
        id: 'review-insert',
        kind: 'insertion',
      }),
    ]);
    expect(editor.getHTML()).toContain('00001001');
    expect(editor.getHTML()).not.toContain('00001002');
    expect(editor.getHTML()).not.toContain('00003002');

    expect(editor.commands.undo()).toBe(true);
    expect(editor.getHTML()).toBe(originalSnapshot);
    expect(editor.commands.redo()).toBe(true);
    expect(editor.commands.acceptAllDocumentChanges()).toBe(true);
    expect(normalizedText(editor)).toBe('Alpha gamma.');

    editor.destroy();
  });

  test('rejects combine baseline mismatches and copies without revisions', () => {
    const original = documentHtml('<p>Canonical baseline.</p>');
    const editor = createEditor(original);
    const originalSnapshot = editor.getHTML();

    const mismatch = applyDocumentComparison(
      editor,
      documentHtml(
        '<p>Different <ins data-document-change="true" data-change-kind="insertion" data-change-id="change" data-change-author="Riley" data-change-date="2026-08-24T09:00:00.000Z">proposal</ins>.</p>',
      ),
      { ...comparisonIdentity, mode: 'combine' },
    );
    expect(mismatch.status).toBe('unsupported');
    expect(mismatch.diagnostics).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ code: 'combine-baseline-mismatch' }),
      ]),
    );
    expect(editor.getHTML()).toBe(originalSnapshot);

    const noRevisions = applyDocumentComparison(editor, original, {
      ...comparisonIdentity,
      mode: 'combine',
    });
    expect(noRevisions.status).toBe('unsupported');
    expect(noRevisions.diagnostics).toEqual(
      expect.arrayContaining([
        expect.objectContaining({ code: 'combine-without-revisions' }),
      ]),
    );
    expect(editor.commands.undo()).toBe(false);

    editor.destroy();
  });

  test.each([
    {
      kind: 'insertion',
      current: '<p>Stable paragraph.</p>',
      reviewed:
        '<p data-document-block-change="true" data-block-change-kind="insertion" data-block-change-id="block-insert" data-block-change-author="Riley" data-block-change-date="2026-08-24T09:00:00.000Z"><ins data-document-change="true" data-change-kind="insertion" data-change-id="block-insert" data-change-author="Riley" data-change-date="2026-08-24T09:00:00.000Z">Inserted paragraph.</ins></p><p>Stable paragraph.</p>',
    },
    {
      kind: 'deletion',
      current: '<p>Deleted paragraph.</p><p>Stable paragraph.</p>',
      reviewed:
        '<p data-document-block-change="true" data-block-change-kind="deletion" data-block-change-id="block-delete" data-block-change-author="Riley" data-block-change-date="2026-08-24T09:00:00.000Z"><del data-document-change="true" data-change-kind="deletion" data-change-id="block-delete" data-change-author="Riley" data-change-date="2026-08-24T09:00:00.000Z">Deleted paragraph.</del></p><p>Stable paragraph.</p>',
    },
  ])('rejects $kind structural block revisions during combine', ({
    current,
    reviewed,
  }) => {
    const editor = createEditor(documentHtml(current));
    const originalSnapshot = editor.getHTML();

    const result = applyDocumentComparison(editor, documentHtml(reviewed), {
      ...comparisonIdentity,
      mode: 'combine',
    });

    expect(result.status).toBe('unsupported');
    expect(result.diagnostics).toEqual([
      expect.objectContaining({ code: 'combine-structural-revisions' }),
    ]);
    expect(editor.getHTML()).toBe(originalSnapshot);
    expect(editor.commands.undo()).toBe(false);

    editor.destroy();
  });

  test('does not overwrite unresolved current revisions during compare or combine', () => {
    const current = documentHtml(
      '<p>Alpha <ins data-document-change="true" data-change-kind="insertion" data-change-id="current-change" data-change-author="Morgan" data-change-date="2026-08-25T09:00:00.000Z">draft</ins>.</p>',
    );
    const editor = createEditor(current);
    const currentSnapshot = editor.getHTML();

    for (const mode of ['compare', 'combine'] as const) {
      const result = applyDocumentComparison(
        editor,
        documentHtml('<p>Alpha revised.</p>'),
        { ...comparisonIdentity, mode },
      );
      expect(result.status).toBe('unsupported');
      expect(result.diagnostics).toEqual(
        expect.arrayContaining([
          expect.objectContaining({ code: 'current-revisions-present' }),
        ]),
      );
      expect(editor.getHTML()).toBe(currentSnapshot);
    }

    editor.destroy();
  });

  test('preserves generated inline and formatting revisions through a native DOCX reopen', async () => {
    const editor = createEditor(
      documentHtml('<p><strong>Alpha</strong> beta.</p>'),
    );
    const result = applyCompare(
      editor,
      documentHtml('<p><em>Alpha</em> gamma.</p>'),
    );
    expect(result.status).toBe('applied');
    expect(result.summary).toEqual({
      deletions: 1,
      formatting: 1,
      insertions: 1,
      paragraphFormatting: 0,
    });

    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a Writer artifact.');
    }
    artifact.content.html = editor.getHTML();
    const exported = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await exported.arrayBuffer());
    const documentXml = await archive.file('word/document.xml')?.async('text');
    expect(documentXml).toContain('<w:del');
    expect(documentXml).toContain('<w:ins');
    expect(documentXml).toContain('<w:rPrChange');

    const reopened = await importOfficeFile(
      new File([exported], 'reopened-comparison.docx', {
        type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      }),
    );
    if (reopened.content.type !== 'document') {
      throw new Error('Expected a reopened Writer artifact.');
    }
    const reopenedEditor = createEditor(reopened.content.html);
    expect(
      collectDocumentChanges(reopenedEditor.state.doc).map((change) => ({
        author: change.author,
        kind: change.kind,
        text: change.text,
      })),
    ).toEqual([
      { author: 'Morgan', kind: 'formatting', text: 'Alpha' },
      { author: 'Morgan', kind: 'deletion', text: 'beta' },
      { author: 'Morgan', kind: 'insertion', text: 'gamma' },
    ]);

    editor.destroy();
    reopenedEditor.destroy();
  });
});

function applyCompare(
  editor: Editor,
  revised: Content,
): DocumentComparisonApplyResult {
  return applyDocumentComparison(editor, revised, {
    ...comparisonIdentity,
    mode: 'compare',
  });
}

function createEditor(content: string): Editor {
  return new Editor({
    extensions: createWorkDocumentExtensions({ isTracking: () => false }),
    content,
  });
}

function documentHtml(body: string): string {
  return `<section data-document-section="true">${body}</section>`;
}

function normalizedText(editor: Editor): string {
  return editor.getText().replaceAll(/\s+/g, ' ').trim();
}

function paragraphTexts(editor: Editor): string[] {
  const paragraphs: string[] = [];
  editor.state.doc.descendants((node) => {
    if (node.type.name === 'paragraph') paragraphs.push(node.textContent);
  });
  return paragraphs;
}
