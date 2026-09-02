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
  activeDocumentBookmark,
  editorDocumentBookmarks,
  normalizeDocumentBookmarkName,
  validateDocumentBookmarkName,
} from '../src/internal/features/work/work-document-bookmarks';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import {
  applyImportedDocxBookmarkMarkers,
  markDocxBookmarks,
} from '../src/internal/features/work/work-docx-bookmark-import';
import { parseXml } from '../src/internal/features/work/work-ooxml-package';

describe('document bookmarks', () => {
  test('validates Word-compatible bookmark names', () => {
    expect(validateDocumentBookmarkName('Architecture_2')).toBeNull();
    expect(validateDocumentBookmarkName('章节_2')).toBeNull();
    expect(validateDocumentBookmarkName('2Architecture')).not.toBeNull();
    expect(validateDocumentBookmarkName('Architecture 2')).not.toBeNull();
    expect(validateDocumentBookmarkName(`A${'b'.repeat(40)}`)).not.toBeNull();
    expect(normalizeDocumentBookmarkName('_Hidden_1')).toBe('_Hidden_1');
  });

  test('inserts, deletes, and restores a cross-paragraph bookmark atomically', () => {
    const editor = createEditor('<p>Alpha target</p><p>Omega target</p>');
    try {
      const alpha = textRange(editor, 'Alpha');
      const omega = textRange(editor, 'Omega');
      editor.commands.setTextSelection({ from: alpha.from, to: omega.to });

      expect(editor.can().insertDocumentBookmark('Architecture')).toBe(true);
      expect(editor.commands.insertDocumentBookmark('Architecture')).toBe(true);
      expect(editorDocumentBookmarks(editor)).toEqual([
        expect.objectContaining({
          name: 'Architecture',
          nativeId: expect.any(Number),
        }),
      ]);
      expect(boundaryKinds(editor)).toEqual(['start', 'end']);
      expect(activeDocumentBookmark(editor)?.name).toBe('Architecture');

      expect(
        editor.commands.deleteDocumentBookmark(
          editorDocumentBookmarks(editor)[0]?.id ?? '',
        ),
      ).toBe(true);
      expect(editorDocumentBookmarks(editor)).toEqual([]);
      expect(editor.state.doc.textContent).toContain(
        'Alpha targetOmega target',
      );

      expect(editor.commands.undo()).toBe(true);
      expect(editorDocumentBookmarks(editor)).toHaveLength(1);
      expect(editor.commands.undo()).toBe(true);
      expect(editorDocumentBookmarks(editor)).toEqual([]);
      expect(editor.commands.redo()).toBe(true);
      expect(editorDocumentBookmarks(editor)[0]?.name).toBe('Architecture');
    } finally {
      editor.destroy();
    }
  });

  test('inserts a bookmark into an empty document section', () => {
    const editor = createEditor('<p></p>');
    try {
      expect(editor.commands.insertDocumentBookmark('Fields_target')).toBe(
        true,
      );
      expect(editorDocumentBookmarks(editor)).toEqual([
        expect.objectContaining({ name: 'Fields_target' }),
      ]);
      expect(boundaryKinds(editor)).toEqual(['start', 'end']);
    } finally {
      editor.destroy();
    }
  });

  test('keeps the original identity when a complete bookmark is copied before it', () => {
    const editor = createEditor('<p><a href="#Architecture">Alpha</a></p>');
    try {
      editor.commands.setTextSelection(textRange(editor, 'Alpha'));
      expect(editor.commands.insertDocumentBookmark('Architecture')).toBe(true);
      const original = editorDocumentBookmarks(editor)[0];
      if (!original) throw new Error('Expected an original bookmark.');
      const copy = editor.state.doc.slice(
        original.from,
        original.to + 1,
        false,
      );
      editor.view.dispatch(
        closeHistory(editor.state.tr).insert(original.from, copy.content),
      );

      const bookmarks = editorDocumentBookmarks(editor);
      expect(bookmarks).toHaveLength(2);
      expect(bookmarks.map(({ name }) => name)).toEqual([
        'Architecture_2',
        'Architecture',
      ]);
      expect(bookmarks[1]?.id).toBe(original.id);
      expect(bookmarks[0]?.id).not.toBe(original.id);
      expect(bookmarks[0]?.nativeId).not.toBe(original.nativeId);
      expect(linkHrefs(editor)).toEqual(['#Architecture_2', '#Architecture']);
      expect(linkClasses(editor)).toEqual(['', '']);

      expect(editor.commands.undo()).toBe(true);
      expect(editorDocumentBookmarks(editor)).toEqual([
        expect.objectContaining({
          id: original.id,
          name: 'Architecture',
          nativeId: original.nativeId,
        }),
      ]);
      expect(editor.commands.redo()).toBe(true);
      expect(editorDocumentBookmarks(editor).map(({ name }) => name)).toEqual([
        'Architecture_2',
        'Architecture',
      ]);
    } finally {
      editor.destroy();
    }
  });

  test('marks internal links missing when their bookmark is removed and repairs them on undo', () => {
    const editor = createEditor(
      [
        '<p><span data-document-bookmark-boundary="true"',
        ' data-bookmark-kind="start" data-bookmark-id="bookmark-target"',
        ' data-bookmark-name="Target" data-office-bookmark-id="17"></span>',
        'Target<span data-document-bookmark-boundary="true"',
        ' data-bookmark-kind="end" data-bookmark-id="bookmark-target"',
        ' data-bookmark-name="Target" data-office-bookmark-id="17"></span></p>',
        '<p><a href="#Target">Jump</a></p>',
      ].join(''),
    );
    try {
      expect(linkClasses(editor)).toEqual(['']);
      expect(editor.commands.deleteDocumentBookmark('bookmark-target')).toBe(
        true,
      );
      expect(linkClasses(editor)).toEqual(['work-document-link-missing']);

      expect(editor.commands.undo()).toBe(true);
      expect(linkClasses(editor)).toEqual(['']);
    } finally {
      editor.destroy();
    }
  });

  test('repairs mismatched paired-boundary metadata on editor initialization', () => {
    const editor = createEditor(
      [
        '<p><span data-document-bookmark-boundary="true"',
        ' data-bookmark-kind="start" data-bookmark-id="bookmark-target"',
        ' data-bookmark-name="Target" data-office-bookmark-id="17"></span>',
        'Target<span data-document-bookmark-boundary="true"',
        ' data-bookmark-kind="end" data-bookmark-id="bookmark-target"',
        ' data-bookmark-name="Wrong" data-office-bookmark-id="18"></span></p>',
      ].join(''),
    );
    try {
      const html = editor.getHTML();
      expect(html.match(/data-bookmark-name="Target"/g)).toHaveLength(2);
      expect(html.match(/data-office-bookmark-id="17"/g)).toHaveLength(2);
      expect(html).not.toContain('data-bookmark-name="Wrong"');
    } finally {
      editor.destroy();
    }
  });

  test('normalizes invalid and duplicate imported names while keeping Word last-target behavior', () => {
    const wordDocument = parseXml(
      [
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">',
        '<w:body>',
        '<w:p><w:bookmarkStart w:id="1" w:name="Target"/><w:r><w:t>First</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>',
        '<w:p><w:bookmarkStart w:id="2" w:name="Target"/><w:r><w:t>Second</w:t></w:r><w:bookmarkEnd w:id="2"/></w:p>',
        '<w:p><w:bookmarkStart w:id="3" w:name="9 bad"/><w:r><w:t>Third</w:t></w:r><w:bookmarkEnd w:id="3"/></w:p>',
        '<w:p><w:fldSimple w:instr="REF Target \\h"><w:r><w:t>Second</w:t></w:r></w:fldSimple></w:p>',
        '</w:body></w:document>',
      ].join(''),
    );
    const markers = markDocxBookmarks(wordDocument);
    expect(markers.bookmarks.map(({ name }) => name)).toEqual([
      'Target_2',
      'Target',
      'Bookmark_9_bad',
    ]);
    expect(markers.references).toEqual([
      expect.objectContaining({
        targetId: markers.bookmarks[1]?.id,
        targetName: 'Target',
        instruction: 'REF Target \\h',
        display: 'Second',
      }),
    ]);
    const reference = markers.references[0];
    if (!reference) throw new Error('Expected a bookmark REF marker.');

    const html = new DOMParser().parseFromString(
      [
        `<p>${markers.bookmarks[0]?.start}First${markers.bookmarks[0]?.end}</p>`,
        `<p>${markers.bookmarks[1]?.start}Second${markers.bookmarks[1]?.end}</p>`,
        `<p>${markers.bookmarks[2]?.start}Third${markers.bookmarks[2]?.end}</p>`,
        `<p>${reference.start}Second${reference.end}</p>`,
        '<p><a href="#Target">Duplicate target</a> ',
        '<a href="#9 BAD">Normalized target</a></p>',
      ].join(''),
      'text/html',
    );
    applyImportedDocxBookmarkMarkers(html, markers);

    expect(
      Array.from(
        html.body.querySelectorAll<HTMLElement>(
          'span[data-bookmark-kind="start"]',
        ),
        (boundary) => boundary.dataset.bookmarkName,
      ),
    ).toEqual(['Target_2', 'Target', 'Bookmark_9_bad']);
    expect(
      Array.from(html.body.querySelectorAll<HTMLAnchorElement>('a'), (link) =>
        link.getAttribute('href'),
      ),
    ).toEqual(['#Target', '#Bookmark_9_bad']);
    expect(
      html.body.querySelector<HTMLElement>(
        'span[data-reference-target-type="bookmark"]',
      )?.dataset,
    ).toMatchObject({
      referenceTargetId: markers.bookmarks[1]?.id,
      referenceTargetName: 'Target',
      referenceInstruction: 'REF Target \\h',
      referenceDisplay: 'Second',
    });
  });

  test('round-trips cross-paragraph bookmarks and internal versus external hyperlinks in DOCX', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html = [
      '<section data-document-section="true" data-section-id="section-1">',
      '<p><span data-document-bookmark-boundary="true"',
      ' data-bookmark-kind="start" data-bookmark-id="bookmark-target"',
      ' data-bookmark-name="Architecture" data-office-bookmark-id="17"></span>',
      'Architecture begins</p>',
      '<p>Architecture ends<span data-document-bookmark-boundary="true"',
      ' data-bookmark-kind="end" data-bookmark-id="bookmark-target"',
      ' data-bookmark-name="Architecture" data-office-bookmark-id="17"></span></p>',
      '<p><a href="#Architecture">Jump inside</a> ',
      '<a href="https://a3s.dev/office">Open site</a></p>',
      '</section>',
    ].join('');

    const first = await createArtifactBlob(artifact);
    const firstArchive = await JSZip.loadAsync(await first.arrayBuffer());
    await assertNativeBookmarkAndLinks(firstArchive);

    const imported = await importOfficeFile(
      new File([first], 'bookmarks.docx', { type: first.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    expect(imported.content.html).toContain('data-bookmark-kind="start"');
    expect(imported.content.html).toContain('data-bookmark-kind="end"');
    expect(imported.content.html).toContain(
      'data-bookmark-name="Architecture"',
    );
    expect(imported.content.html).toContain('href="#Architecture"');
    expect(imported.content.html).toContain('href="https://a3s.dev/office"');
    expect(imported.compatibility.issues).toContainEqual(
      expect.objectContaining({
        code: 'docx.bookmarks-links',
        severity: 'info',
      }),
    );

    const reopened = await createArtifactBlob(imported);
    const reopenedArchive = await JSZip.loadAsync(await reopened.arrayBuffer());
    await assertNativeBookmarkAndLinks(reopenedArchive);
  });
});

function createEditor(html: string): Editor {
  return new Editor({
    extensions: createWorkDocumentExtensions(),
    content: `<section data-document-section="true">${html}</section>`,
  });
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

function boundaryKinds(editor: Editor): string[] {
  const kinds: string[] = [];
  editor.state.doc.descendants((node) => {
    if (node.type.name === 'documentBookmarkBoundary') {
      kinds.push(String(node.attrs.kind));
    }
  });
  return kinds;
}

function linkClasses(editor: Editor): string[] {
  const classes: string[] = [];
  editor.state.doc.descendants((node) => {
    if (!node.isText) return;
    for (const mark of node.marks) {
      if (mark.type.name === 'link')
        classes.push(String(mark.attrs.class ?? ''));
    }
  });
  return classes;
}

function linkHrefs(editor: Editor): string[] {
  const hrefs: string[] = [];
  editor.state.doc.descendants((node) => {
    if (!node.isText) return;
    for (const mark of node.marks) {
      if (mark.type.name === 'link') hrefs.push(String(mark.attrs.href ?? ''));
    }
  });
  return hrefs;
}

async function assertNativeBookmarkAndLinks(archive: JSZip): Promise<void> {
  const documentEntry = archive.file('word/document.xml');
  const relationshipsEntry = archive.file('word/_rels/document.xml.rels');
  expect(documentEntry).not.toBeNull();
  expect(relationshipsEntry).not.toBeNull();
  const [documentXml = '', relationshipsXml = ''] = await Promise.all([
    documentEntry?.async('string'),
    relationshipsEntry?.async('string'),
  ]);
  const paragraphs = Array.from(
    documentXml.matchAll(/<w:p(?:\s[^>]*)?>[\s\S]*?<\/w:p>/g),
    ([paragraph]) => paragraph,
  );
  const startParagraph = paragraphs.find((paragraph) =>
    paragraph.includes('Architecture begins'),
  );
  const endParagraph = paragraphs.find((paragraph) =>
    paragraph.includes('Architecture ends'),
  );
  expect(startParagraph).toContain(
    '<w:bookmarkStart w:id="17" w:name="Architecture"/>',
  );
  expect(endParagraph).toContain('<w:bookmarkEnd w:id="17"/>');
  expect(documentXml).toMatch(
    /<w:hyperlink(?=[^>]*w:anchor="Architecture")[^>]*>[\s\S]*?Jump inside[\s\S]*?<\/w:hyperlink>/,
  );
  const internal = documentXml.match(
    /<w:hyperlink(?=[^>]*w:anchor="Architecture")[^>]*>/,
  )?.[0];
  expect(internal).not.toContain('r:id=');
  const externalRelationshipId = documentXml.match(
    /<w:hyperlink(?=[^>]*r:id="([^"]+)")[^>]*>[\s\S]*?Open site[\s\S]*?<\/w:hyperlink>/,
  )?.[1];
  expect(externalRelationshipId).toBeTruthy();
  expect(relationshipsXml).toMatch(
    new RegExp(
      `<Relationship(?=[^>]*Id="${externalRelationshipId}")(?=[^>]*Target="https://a3s.dev/office")[^>]*/>`,
    ),
  );
  expect(relationshipsXml).not.toContain('Target="#Architecture"');
}
