import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import {
  attribute,
  descendants,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';
import type {
  WorkDocumentComment,
  WorkDocumentContent,
} from '../src/internal/features/work/work-types';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from '../src/internal/features/work/work-docx-settings-xml';

const VENDOR_NAMESPACE = 'urn:a3s:test:note-comment-extension';
const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const RELATIONSHIPS_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const WORD_2015_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2012/wordml';
const COMMENTS_IDS_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2016/wordml/cid';
const MARKUP_COMPATIBILITY_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';

describe('DOCX note and comment extension preservation', () => {
  test('retains passive root metadata only on uniquely matched notes and comments', async () => {
    const content = documentContent(true);
    const seed = await createDocxBlob(content);
    const source = await JSZip.loadAsync(await seed.arrayBuffer());
    await decoratePart(source, 'word/footnotes.xml', 'footnote', '27', 'foot');
    const footnotes = await xmlEntry(source, 'word/footnotes.xml');
    elementById(footnotes, 'footnote', '27').setAttributeNS(
      WORD_NAMESPACE,
      'w:type',
      'normal',
    );
    source.file(
      'word/footnotes.xml',
      new XMLSerializer().serializeToString(footnotes),
    );
    await decoratePart(source, 'word/endnotes.xml', 'endnote', '41', 'end');
    await decoratePart(source, 'word/comments.xml', 'comment', '42', 'comment');
    await decorateCommentExtended(source, '42', 'comment-ex');

    const output = await JSZip.loadAsync(
      await (
        await createDocxBlob(
          content,
          await source.generateAsync({ type: 'arraybuffer' }),
        )
      ).arrayBuffer(),
    );
    await expectDecoratedPart(
      output,
      'word/footnotes.xml',
      'footnote',
      '27',
      'foot',
    );
    await expectDecoratedPart(
      output,
      'word/endnotes.xml',
      'endnote',
      '41',
      'end',
    );
    await expectDecoratedPart(
      output,
      'word/comments.xml',
      'comment',
      '42',
      'comment',
    );
    const comments = await xmlEntry(output, 'word/comments.xml');
    const extended = await xmlEntry(output, 'word/commentsExtended.xml');
    const commentParaId = lastParagraphId(
      elementById(comments, 'comment', '42'),
    );
    const commentEx = descendants(extended, 'commentEx').find(
      (item) => attribute(item, 'paraId') === commentParaId,
    );
    expect(vendorAttribute(commentEx, 'token')).toBe('comment-ex');
    expect(descendants(commentEx ?? extended, 'passiveMeta')).toHaveLength(1);
  });

  test('does not revive deleted note and comment semantic parts from the source package', async () => {
    const content = documentContent(true);
    const seed = await createDocxBlob(content);
    const output = await JSZip.loadAsync(
      await (
        await createDocxBlob(
          {
            type: 'document',
            pageSize: 'a4',
            html: '<section data-document-section="true"><p>Clean</p></section>',
          },
          await seed.arrayBuffer(),
        )
      ).arrayBuffer(),
    );

    const footnotes = await xmlEntry(output, 'word/footnotes.xml');
    const endnotes = await xmlEntry(output, 'word/endnotes.xml');
    const comments = await xmlEntry(output, 'word/comments.xml');
    expect(
      descendants(footnotes, 'footnote').filter(
        (item) => !attribute(item, 'type'),
      ),
    ).toHaveLength(0);
    expect(
      descendants(endnotes, 'endnote').filter(
        (item) => !attribute(item, 'type'),
      ),
    ).toHaveLength(0);
    expect(descendants(comments, 'comment')).toHaveLength(0);
    expect(output.file('word/commentsExtended.xml')).toBeNull();
  });

  test('rebinds durable comment IDs to regenerated paragraph IDs and drops deleted records', async () => {
    const sourceContent = documentContent(true);
    const seed = await createDocxBlob(sourceContent);
    const source = await JSZip.loadAsync(await seed.arrayBuffer());
    const sourceComments = await xmlEntry(source, 'word/comments.xml');
    const sourceParaIds = new Map(
      descendants(sourceComments, 'comment').map((item) => [
        attribute(item, 'id') ?? '',
        lastParagraphId(item),
      ]),
    );
    const reboundSourceIds = new Map([
      ['42', '6A000042'],
      ['91', '6A000091'],
      ['7', '6A000007'],
    ]);
    await rewriteCommentParagraphIds(source, sourceParaIds, reboundSourceIds);
    addCommentsIdsPart(source, [
      [reboundSourceIds.get('42') ?? '', '11223344'],
      [reboundSourceIds.get('91') ?? '', '22334455'],
      [reboundSourceIds.get('7') ?? '', '33445566'],
    ]);

    const edited = documentContent(false);
    const output = await JSZip.loadAsync(
      await (
        await createDocxBlob(
          edited,
          await source.generateAsync({ type: 'arraybuffer' }),
        )
      ).arrayBuffer(),
    );
    const comments = await xmlEntry(output, 'word/comments.xml');
    const expectedParaIds = new Map(
      descendants(comments, 'comment').map((item) => [
        attribute(item, 'id') ?? '',
        lastParagraphId(item),
      ]),
    );
    const commentsIds = await xmlEntry(output, 'word/commentsIds.xml');
    const entries = descendants(commentsIds, 'commentId').map((item) => ({
      durableId: attribute(item, 'durableId'),
      paraId: attribute(item, 'paraId'),
    }));
    expect(entries).toEqual([
      { durableId: '11223344', paraId: expectedParaIds.get('42') },
      { durableId: '22334455', paraId: expectedParaIds.get('91') },
    ]);
  });

  test('does not preserve an orphan commentsIds part without its package relationship', async () => {
    const content = documentContent(false);
    const seed = await createDocxBlob(content);
    const source = await JSZip.loadAsync(await seed.arrayBuffer());
    const comments = await xmlEntry(source, 'word/comments.xml');
    addCommentsIdsPart(
      source,
      [[lastParagraphId(elementById(comments, 'comment', '42')), '11223344']],
      false,
    );

    const output = await JSZip.loadAsync(
      await (
        await createDocxBlob(
          content,
          await source.generateAsync({ type: 'arraybuffer' }),
        )
      ).arrayBuffer(),
    );
    expect(output.file('word/commentsIds.xml')).toBeNull();
  });
});

function documentContent(includeSecond: boolean): WorkDocumentContent {
  const comments: WorkDocumentComment[] = [
    {
      id: 'docx-comment-42',
      author: 'Reviewer',
      date: '2026-08-12T00:00:00.000Z',
      text: 'First comment',
      resolved: true,
      replies: [
        {
          id: 'docx-comment-reply-91',
          author: 'Responder',
          date: '2026-08-12T00:01:00.000Z',
          text: 'First reply',
        },
      ],
    },
  ];
  if (includeSecond) {
    comments.unshift({
      id: 'docx-comment-7',
      author: 'Other',
      date: '2026-08-12T00:02:00.000Z',
      text: 'Deleted comment',
      resolved: false,
    });
  }
  return {
    type: 'document',
    pageSize: 'a4',
    html: [
      '<section data-document-section="true">',
      includeSecond
        ? '<p><span data-document-comment="true" data-comment-id="docx-comment-7">Delete me</span></p>'
        : '',
      '<p>Main<sup data-document-note-reference="true" data-note-kind="footnote" data-note-id="docx-footnote-27" data-note-number="1">1</sup>',
      '<sup data-document-note-reference="true" data-note-kind="endnote" data-note-id="docx-endnote-41" data-note-number="1">1</sup></p>',
      '<p><span data-document-comment="true" data-comment-id="docx-comment-42">Keep me</span></p>',
      '<aside data-document-note="true" data-note-kind="footnote" data-note-id="docx-footnote-27" data-note-number="1"><p>Foot text</p></aside>',
      '<aside data-document-note="true" data-note-kind="endnote" data-note-id="docx-endnote-41" data-note-number="1"><p>End text</p></aside>',
      '</section>',
    ].join(''),
    comments,
  };
}

async function decoratePart(
  archive: JSZip,
  path: string,
  localName: string,
  id: string,
  token: string,
): Promise<void> {
  const document = await xmlEntry(archive, path);
  declareVendor(document.documentElement);
  const element = elementById(document, localName, id);
  decorateElement(document, element, token);
  archive.file(path, new XMLSerializer().serializeToString(document));
}

async function decorateCommentExtended(
  archive: JSZip,
  commentId: string,
  token: string,
): Promise<void> {
  const comments = await xmlEntry(archive, 'word/comments.xml');
  const paraId = lastParagraphId(elementById(comments, 'comment', commentId));
  const document = await xmlEntry(archive, 'word/commentsExtended.xml');
  declareVendor(document.documentElement);
  const element = descendants(document, 'commentEx').find(
    (item) => attribute(item, 'paraId') === paraId,
  );
  if (!element) throw new Error(`Missing commentEx for ${commentId}.`);
  decorateElement(document, element, token);
  archive.file(
    'word/commentsExtended.xml',
    new XMLSerializer().serializeToString(document),
  );
}

function decorateElement(document: Document, element: Element, token: string) {
  element.setAttributeNS(VENDOR_NAMESPACE, 'vendor:token', token);
  const passive = document.createElementNS(
    VENDOR_NAMESPACE,
    'vendor:passiveMeta',
  );
  passive.textContent = `${token}-extension`;
  element.append(passive);
  const relationshipBound = document.createElementNS(
    VENDOR_NAMESPACE,
    'vendor:relationshipBound',
  );
  relationshipBound.setAttributeNS(RELATIONSHIPS_NAMESPACE, 'r:id', 'rId999');
  element.append(relationshipBound);
  const semantic = document.createElementNS(
    WORD_2015_NAMESPACE,
    'w15:semanticMeta',
  );
  element.append(semantic);
}

async function expectDecoratedPart(
  archive: JSZip,
  path: string,
  localName: string,
  id: string,
  token: string,
): Promise<void> {
  const document = await xmlEntry(archive, path);
  const element = elementById(document, localName, id);
  expect(vendorAttribute(element, 'token')).toBe(token);
  expect(descendants(element, 'passiveMeta')).toHaveLength(1);
  expect(descendants(element, 'relationshipBound')).toHaveLength(0);
  expect(descendants(element, 'semanticMeta')).toHaveLength(0);
}

function addCommentsIdsPart(
  archive: JSZip,
  entries: Array<[string, string]>,
  relationship = true,
): void {
  archive.file(
    'word/commentsIds.xml',
    [
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w16cid:commentsIds xmlns:w16cid="${COMMENTS_IDS_NAMESPACE}">`,
      ...entries.map(
        ([paraId, durableId]) =>
          `<w16cid:commentId w16cid:paraId="${paraId}" w16cid:durableId="${durableId}"/>`,
      ),
      '</w16cid:commentsIds>',
    ].join(''),
  );
  appendXml(
    archive,
    '[Content_Types].xml',
    `<Override PartName="/word/commentsIds.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"/>`,
    '</Types>',
  );
  if (relationship) {
    appendXml(
      archive,
      'word/_rels/document.xml.rels',
      '<Relationship Id="rIdA3SCommentsIds" Type="http://schemas.microsoft.com/office/2016/09/relationships/commentsIds" Target="commentsIds.xml"/>',
      '</Relationships>',
    );
  }
}

async function rewriteCommentParagraphIds(
  archive: JSZip,
  current: ReadonlyMap<string, string>,
  replacement: ReadonlyMap<string, string>,
): Promise<void> {
  for (const path of ['word/comments.xml', 'word/commentsExtended.xml']) {
    let source = await archive.file(path)?.async('string');
    if (!source) throw new Error(`Missing ${path}.`);
    for (const [commentId, paragraphId] of current) {
      const next = replacement.get(commentId);
      if (next) source = source.replaceAll(paragraphId, next);
    }
    archive.file(path, source);
  }
}

function appendXml(
  archive: JSZip,
  path: string,
  value: string,
  closingTag: string,
): void {
  const entry = archive.file(path);
  if (!entry) throw new Error(`Missing ${path}.`);
  archive.file(
    path,
    entry
      .async('string')
      .then((source) => source.replace(closingTag, `${value}${closingTag}`)),
  );
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

function lastParagraphId(comment: Element): string {
  return attribute(descendants(comment, 'p').at(-1) ?? comment, 'paraId') ?? '';
}

function declareVendor(root: Element): void {
  root.setAttributeNS(
    'http://www.w3.org/2000/xmlns/',
    'xmlns:vendor',
    VENDOR_NAMESPACE,
  );
  root.setAttributeNS(
    'http://www.w3.org/2000/xmlns/',
    'xmlns:mc',
    MARKUP_COMPATIBILITY_NAMESPACE,
  );
  root.setAttributeNS(MARKUP_COMPATIBILITY_NAMESPACE, 'mc:Ignorable', 'vendor');
}

function vendorAttribute(element: Element | undefined, name: string) {
  if (!element) return null;
  return (
    Array.from(element.attributes).find(
      (item) =>
        xmlAttributeLocalName(item) === name &&
        xmlAttributeNamespace(element, item) === VENDOR_NAMESPACE,
    )?.value ?? null
  );
}

async function xmlEntry(archive: JSZip, path: string): Promise<Document> {
  const source = await archive.file(path)?.async('uint8array');
  if (!source) throw new Error(`Missing ${path}.`);
  return parseXml(new TextDecoder().decode(source), path);
}
