import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import {
  createArtifact,
  createArtifactBlob,
  importOfficeFile,
} from '../src/core';
import { normalizeDocumentNotesHtml } from '../src/internal/features/work/work-document-notes';
import {
  attribute,
  descendants,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';

describe('DOCX note and comment identities', () => {
  test('round-trips imported native footnote and endnote IDs', async () => {
    const seed = await createArtifactBlob(noteArtifact());
    const source = await JSZip.loadAsync(await seed.arrayBuffer());
    await replaceNoteId(source, 'footnote', '1', '27');
    await replaceNoteId(source, 'endnote', '1', '41');
    const sourceBlob = new Blob(
      [await source.generateAsync({ type: 'arraybuffer' })],
      { type: seed.type },
    );

    const imported = await importOfficeFile(
      new File([sourceBlob], 'native-note-ids.docx', {
        type: sourceBlob.type,
      }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    expect(
      html.body.querySelector<HTMLElement>(
        'sup[data-document-note-reference][data-note-kind="footnote"]',
      )?.dataset.noteId,
    ).toBe('docx-footnote-27');
    expect(
      html.body.querySelector<HTMLElement>(
        'sup[data-document-note-reference][data-note-kind="endnote"]',
      )?.dataset.noteId,
    ).toBe('docx-endnote-41');

    const output = await JSZip.loadAsync(
      await (await createArtifactBlob(imported)).arrayBuffer(),
    );
    expect(await noteIds(output, 'footnote')).toContain('27');
    expect(await noteIds(output, 'endnote')).toContain('41');
    expect(await documentReferenceIds(output, 'footnote')).toEqual(['27']);
    expect(await documentReferenceIds(output, 'endnote')).toEqual(['41']);
  });

  test('retains imported native comment and reply IDs after thread reordering', async () => {
    const sourceContent = commentContent([
      comment('docx-comment-42', 'First', [
        reply('docx-comment-reply-91', 'First reply'),
      ]),
      comment('docx-comment-7', 'Second'),
    ]);
    const seed = await createDocx(sourceContent);

    const reordered = commentContent([
      comment('docx-comment-7', 'Second'),
      comment('docx-comment-42', 'First', [
        reply('docx-comment-reply-91', 'First reply'),
      ]),
    ]);
    const output = await JSZip.loadAsync(
      await (
        await createDocx(reordered, await seed.arrayBuffer())
      ).arrayBuffer(),
    );
    const comments = await xmlEntry(output, 'word/comments.xml');
    const ids = descendants(comments, 'comment').map((item) =>
      attribute(item, 'id'),
    );
    expect(ids).toEqual(['7', '42', '91']);

    const extended = await xmlEntry(output, 'word/commentsExtended.xml');
    const commentById = new Map(
      descendants(comments, 'comment').map((item) => [
        attribute(item, 'id'),
        lastParagraphId(item),
      ]),
    );
    const replyEntry = descendants(extended, 'commentEx').find(
      (item) => attribute(item, 'paraId') === commentById.get('91'),
    );
    expect(
      attribute(replyEntry ?? extended.documentElement, 'paraIdParent'),
    ).toBe(commentById.get('42'));
  });

  test('emits and reimports resolved state for a comment without replies', async () => {
    const content = commentContent([
      { ...comment('docx-comment-42', 'Resolved root'), resolved: true },
    ]);
    const blob = await createDocx(content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const comments = await xmlEntry(archive, 'word/comments.xml');
    const extended = await xmlEntry(archive, 'word/commentsExtended.xml');
    const paraId = lastParagraphId(elementById(comments, 'comment', '42'));
    const metadata = descendants(extended, 'commentEx').find(
      (item) => attribute(item, 'paraId') === paraId,
    );
    expect(attribute(metadata ?? extended.documentElement, 'done')).toBe('1');

    const imported = await importOfficeFile(
      new File([blob], 'resolved-root.docx', { type: blob.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    expect(imported.content.comments).toEqual([
      expect.objectContaining({ id: 'docx-comment-42', resolved: true }),
    ]);
  });

  test('gives copied native notes and colliding comment records fresh OOXML IDs', async () => {
    const notesHtml = normalizeDocumentNotesHtml(
      [
        '<section data-document-section="true">',
        '<p>A<sup data-document-note-reference="true" data-note-kind="footnote" data-note-id="docx-footnote-27">1</sup></p>',
        '<p>B<sup data-document-note-reference="true" data-note-kind="footnote" data-note-id="docx-footnote-27">1</sup></p>',
        '<aside data-document-note="true" data-note-kind="footnote" data-note-id="docx-footnote-27"><p>Copied note</p></aside>',
        '</section>',
      ].join(''),
    );
    const notes = await createDocx({
      type: 'document',
      pageSize: 'a4',
      html: notesHtml,
    });
    const notesArchive = await JSZip.loadAsync(await notes.arrayBuffer());
    expect(new Set(await noteIds(notesArchive, 'footnote'))).toEqual(
      new Set(['1', '27']),
    );

    const collisions = commentContent([
      comment('docx-comment-42', 'Root', [
        reply('docx-comment-reply-42', 'Reply'),
      ]),
    ]);
    const commentsArchive = await JSZip.loadAsync(
      await (await createDocx(collisions)).arrayBuffer(),
    );
    const comments = await xmlEntry(commentsArchive, 'word/comments.xml');
    const ids = descendants(comments, 'comment').map((item) =>
      attribute(item, 'id'),
    );
    expect(ids).toEqual(['0', '1']);
    expect(ids).not.toContain('42');
  });

  test('round-trips signed native comment IDs without using them for new records', async () => {
    const content = commentContent([
      comment('docx-comment--5', 'Signed identity'),
      comment('new-comment', 'New identity'),
    ]);
    const blob = await createDocx(content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const ids = descendants(
      await xmlEntry(archive, 'word/comments.xml'),
      'comment',
    ).map((item) => attribute(item, 'id'));
    expect(ids).toEqual(['-5', '0']);

    const imported = await importOfficeFile(
      new File([blob], 'signed-comment-id.docx', { type: blob.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    expect(imported.content.comments?.map(({ id }) => id)).toEqual([
      'docx-comment--5',
      'docx-comment-0',
    ]);
  });
});

function noteArtifact() {
  const artifact = createArtifact('blank-document');
  if (artifact.content.type !== 'document') {
    throw new Error('Expected a document artifact.');
  }
  artifact.content.html = [
    '<section data-document-section="true">',
    '<p>Body<sup data-document-note-reference="true" data-note-kind="footnote" data-note-id="foot" data-note-number="1">1</sup>',
    '<sup data-document-note-reference="true" data-note-kind="endnote" data-note-id="end" data-note-number="1">1</sup></p>',
    '<aside data-document-note="true" data-note-kind="footnote" data-note-id="foot" data-note-number="1"><p>Foot text</p></aside>',
    '<aside data-document-note="true" data-note-kind="endnote" data-note-id="end" data-note-number="1"><p>End text</p></aside>',
    '</section>',
  ].join('');
  return artifact;
}

async function createDocx(
  content: ReturnType<typeof commentContent>,
  source?: ArrayBuffer,
): Promise<Blob> {
  const { createDocxBlob } = await import(
    '../src/internal/features/work/work-docx-export'
  );
  return createDocxBlob(content, source);
}

function commentContent(comments: ReturnType<typeof comment>[]) {
  return {
    type: 'document' as const,
    pageSize: 'a4' as const,
    html: [
      '<section data-document-section="true">',
      ...comments.map(
        (item) =>
          `<p><span data-document-comment="true" data-comment-id="${item.id}">${item.text}</span></p>`,
      ),
      '</section>',
    ].join(''),
    comments,
  };
}

function comment(
  id: string,
  text: string,
  replies?: ReturnType<typeof reply>[],
) {
  return {
    id,
    author: 'Reviewer',
    date: '2026-08-12T00:00:00.000Z',
    text,
    resolved: false,
    replies,
  };
}

function reply(id: string, text: string) {
  return {
    id,
    author: 'Responder',
    date: '2026-08-12T00:01:00.000Z',
    text,
  };
}

async function replaceNoteId(
  archive: JSZip,
  kind: 'footnote' | 'endnote',
  from: string,
  to: string,
): Promise<void> {
  for (const path of ['word/document.xml', `word/${kind}s.xml`]) {
    const source = await archive.file(path)?.async('string');
    if (!source) throw new Error(`Missing ${path}.`);
    archive.file(
      path,
      source.replace(
        new RegExp(
          `(<w:${kind}(?:Reference)?\\b[^>]*\\bw:id=")${from}(")`,
          'g',
        ),
        `$1${to}$2`,
      ),
    );
  }
}

async function noteIds(
  archive: JSZip,
  kind: 'footnote' | 'endnote',
): Promise<string[]> {
  const document = await xmlEntry(archive, `word/${kind}s.xml`);
  return descendants(document, kind)
    .filter((item) => !attribute(item, 'type'))
    .map((item) => attribute(item, 'id') ?? '');
}

async function documentReferenceIds(
  archive: JSZip,
  kind: 'footnote' | 'endnote',
): Promise<string[]> {
  const document = await xmlEntry(archive, 'word/document.xml');
  return descendants(document, `${kind}Reference`).map(
    (item) => attribute(item, 'id') ?? '',
  );
}

function lastParagraphId(comment: Element): string {
  const paragraph = descendants(comment, 'p').at(-1);
  return paragraph ? (attribute(paragraph, 'paraId') ?? '') : '';
}

function elementById(
  document: Document,
  localName: string,
  id: string,
): Element {
  const element = descendants(document, localName).find(
    (item) => attribute(item, 'id') === id,
  );
  if (!element) throw new Error(`Missing ${localName} ${id}.`);
  return element;
}

async function xmlEntry(archive: JSZip, path: string): Promise<Document> {
  const source = await archive.file(path)?.async('uint8array');
  if (!source) throw new Error(`Missing ${path}.`);
  return parseXml(new TextDecoder().decode(source), path);
}
