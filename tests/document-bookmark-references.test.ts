import { describe, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import { closeHistory } from '@tiptap/pm/history';
import JSZip from 'jszip';
import {
  createArtifact,
  createArtifactBlob,
  importOfficeFile,
} from '../src/core';
import {
  documentBookmarkReferenceInstruction,
  normalizeDocumentBookmarkReferencesHtml,
  supportedDocxBookmarkReferenceInstruction,
} from '../src/internal/features/work/work-document-bookmark-references';
import {
  editorDocumentBookmarkReferenceTargets,
  editorDocumentBookmarks,
} from '../src/internal/features/work/work-document-bookmarks';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';

describe('document bookmark references', () => {
  test('recognizes the supported REF switch subset and normalizes the rest', () => {
    expect(
      supportedDocxBookmarkReferenceInstruction(
        'REF Architecture \\h \\* MERGEFORMAT',
      ),
    ).toBe(true);
    expect(
      supportedDocxBookmarkReferenceInstruction('REF Architecture \\p'),
    ).toBe(false);
    expect(
      supportedDocxBookmarkReferenceInstruction('REF Architecture \\# "0"'),
    ).toBe(false);
    expect(
      documentBookmarkReferenceInstruction(
        'Architecture_2',
        'REF Architecture \\p',
      ),
    ).toBe('REF Architecture_2');
  });

  test('updates bookmark REF display and repairs a missing target on undo', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: [
        '<section data-document-section="true">',
        '<p>Architecture</p><p>target</p><p>Reference</p>',
        '</section>',
      ].join(''),
    });
    try {
      const architecture = textRange(editor, 'Architecture');
      const targetText = textRange(editor, 'target');
      editor.commands.setTextSelection({
        from: architecture.from,
        to: targetText.to,
      });
      expect(editor.commands.insertDocumentBookmark('Architecture')).toBe(true);
      const bookmark = editorDocumentBookmarks(editor)[0];
      if (!bookmark) throw new Error('Expected a bookmark target.');
      const referenceTarget = editorDocumentBookmarkReferenceTargets(editor)[0];
      if (!referenceTarget) throw new Error('Expected a reference target.');

      editor.commands.setTextSelection(textRange(editor, 'Reference').to);
      expect(
        editor.commands.insertDocumentCrossReference(referenceTarget),
      ).toBe(true);
      expect(bookmarkReferenceStates(editor)).toEqual([
        {
          targetId: bookmark.id,
          targetName: 'Architecture',
          display: 'Architecture target',
          orphaned: false,
        },
      ]);

      editor.view.dispatch(
        editor.state.tr.insertText(
          'Updated ',
          textRange(editor, 'Architecture').from,
        ),
      );
      expect(bookmarkReferenceStates(editor)[0]?.display).toBe(
        'Updated Architecture target',
      );

      expect(editor.commands.deleteDocumentBookmark(bookmark.id)).toBe(true);
      expect(bookmarkReferenceStates(editor)[0]).toMatchObject({
        display: 'Updated Architecture target',
        orphaned: true,
      });
      expect(editor.commands.undo()).toBe(true);
      expect(bookmarkReferenceStates(editor)[0]).toMatchObject({
        display: 'Updated Architecture target',
        orphaned: false,
      });
    } finally {
      editor.destroy();
    }
  });

  test('normalizes a stale cross-paragraph HTML reference display', () => {
    const source = [
      '<p><span data-document-bookmark-boundary="true"',
      ' data-bookmark-kind="start" data-bookmark-id="bookmark-target"',
      ' data-bookmark-name="Architecture"></span>Architecture</p>',
      '<p>target<span data-document-bookmark-boundary="true"',
      ' data-bookmark-kind="end" data-bookmark-id="bookmark-target"',
      ' data-bookmark-name="Architecture"></span></p>',
      '<p><span data-document-cross-reference="true"',
      ' data-reference-target-type="bookmark"',
      ' data-reference-target-id="bookmark-target"',
      ' data-reference-target-name="Architecture"',
      ' data-reference-display="Stale">Stale</span></p>',
    ].join('');

    const document = new DOMParser().parseFromString(
      normalizeDocumentBookmarkReferencesHtml(source),
      'text/html',
    );
    const reference = document.body.querySelector<HTMLElement>(
      'span[data-reference-target-type="bookmark"]',
    );
    expect(reference?.dataset.referenceDisplay).toBe('Architecture target');
    expect(reference?.textContent).toBe('Architecture target');
    expect(reference?.dataset.referenceOrphaned).toBeUndefined();
  });

  test('retargets a copied self-reference to the copied bookmark identity', () => {
    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content:
        '<section data-document-section="true"><p>Architecture</p></section>',
    });
    try {
      editor.commands.setTextSelection(textRange(editor, 'Architecture'));
      expect(editor.commands.insertDocumentBookmark('Architecture')).toBe(true);
      const target = editorDocumentBookmarkReferenceTargets(editor)[0];
      const original = editorDocumentBookmarks(editor)[0];
      if (!target || !original) throw new Error('Expected a bookmark target.');
      editor.commands.setTextSelection(original.to);
      expect(editor.commands.insertDocumentCrossReference(target)).toBe(true);

      const complete = editorDocumentBookmarks(editor)[0];
      if (!complete) throw new Error('Expected a complete bookmark.');
      const copy = editor.state.doc.slice(
        complete.from,
        complete.to + 1,
        false,
      );
      editor.view.dispatch(
        closeHistory(editor.state.tr).insert(complete.from, copy.content),
      );

      const bookmarks = editorDocumentBookmarks(editor);
      const references = bookmarkReferenceStates(editor);
      expect(bookmarks.map(({ name }) => name)).toEqual([
        'Architecture_2',
        'Architecture',
      ]);
      expect(references).toHaveLength(2);
      expect(references[0]).toMatchObject({
        targetId: bookmarks[0]?.id,
        targetName: 'Architecture_2',
        orphaned: false,
      });
      expect(references[1]).toMatchObject({
        targetId: bookmarks[1]?.id,
        targetName: 'Architecture',
        orphaned: false,
      });
    } finally {
      editor.destroy();
    }
  });

  test('round-trips a native Word REF field against a body bookmark', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = [
      '<section data-document-section="true">',
      '<p><span data-document-bookmark-boundary="true"',
      ' data-bookmark-kind="start" data-bookmark-id="bookmark-target"',
      ' data-bookmark-name="Architecture" data-office-bookmark-id="17"></span>',
      'Architecture target<span data-document-bookmark-boundary="true"',
      ' data-bookmark-kind="end" data-bookmark-id="bookmark-target"',
      ' data-bookmark-name="Architecture" data-office-bookmark-id="17"></span></p>',
      '<p><span data-document-cross-reference="true"',
      ' data-reference-target-type="bookmark"',
      ' data-reference-target-id="bookmark-target"',
      ' data-reference-target-name="Architecture"',
      ' data-reference-instruction="REF Architecture \\h"',
      ' data-reference-display="Architecture target">Architecture target</span></p>',
      '</section>',
    ].join('');

    const first = await createArtifactBlob(artifact);
    await assertNativeBookmarkReference(first);

    const imported = await importOfficeFile(
      new File([first], 'bookmark-reference.docx', { type: first.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    expect(imported.content.html).toContain(
      'data-reference-target-type="bookmark"',
    );
    expect(imported.content.html).toContain(
      'data-reference-target-name="Architecture"',
    );
    expect(imported.content.html).toContain(
      'data-reference-display="Architecture target"',
    );
    expect(imported.content.html).not.toContain('data-reference-orphaned');

    await assertNativeBookmarkReference(await createArtifactBlob(imported));
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
  if (!range) throw new Error(`Unable to find "${text}".`);
  return range;
}

function bookmarkReferenceStates(editor: Editor): Array<{
  targetId: string;
  targetName: string;
  display: string;
  orphaned: boolean;
}> {
  const states: Array<{
    targetId: string;
    targetName: string;
    display: string;
    orphaned: boolean;
  }> = [];
  editor.state.doc.descendants((node) => {
    if (
      node.type.name !== 'documentCrossReference' ||
      node.attrs.targetType !== 'bookmark'
    ) {
      return;
    }
    states.push({
      targetId: String(node.attrs.targetId),
      targetName: String(node.attrs.targetName),
      display: String(node.attrs.display),
      orphaned: Boolean(node.attrs.orphaned),
    });
  });
  return states;
}

async function assertNativeBookmarkReference(blob: Blob): Promise<void> {
  const archive = await JSZip.loadAsync(await blob.arrayBuffer());
  const xml = (await archive.file('word/document.xml')?.async('string')) ?? '';
  expect(xml).toContain('<w:bookmarkStart w:id="17" w:name="Architecture"/>');
  expect(xml).toMatch(
    /<w:fldSimple[^>]*w:instr="REF Architecture \\h"[^>]*>[\s\S]*?Architecture target[\s\S]*?<\/w:fldSimple>/,
  );
}
