import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import { analyzeDocxCompatibility } from '../src/internal/features/work/work-office-diagnostics';
import {
  attribute,
  descendants,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from '../src/internal/features/work/work-docx-settings-xml';
import type { WorkDocumentContent } from '../src/internal/features/work/work-types';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STRICT_WORD_NAMESPACE =
  'http://purl.oclc.org/ooxml/wordprocessingml/main';
const VENDOR_NAMESPACE = 'urn:a3s:test:note-comment-boundary';
const MARKUP_COMPATIBILITY_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';
const COMMENTS_IDS_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2016/wordml/cid';

describe('DOCX note and comment extension boundaries', () => {
  test('preserves passive metadata from strict UTF-16 note and comment parts', async () => {
    const content = documentContent();
    const seed = await createDocxBlob(content);
    const source = await JSZip.loadAsync(await seed.arrayBuffer());
    for (const [path, localName, id, token] of [
      ['word/footnotes.xml', 'footnote', '27', 'strict-note'],
      ['word/comments.xml', 'comment', '42', 'strict-comment'],
    ] as const) {
      const document = await xmlEntry(source, path);
      decorate(document, elementById(document, localName, id), token);
      source.file(
        path,
        utf16Xml(
          new XMLSerializer()
            .serializeToString(document)
            .replaceAll(WORD_NAMESPACE, STRICT_WORD_NAMESPACE),
        ),
      );
    }

    const output = await JSZip.loadAsync(
      await (
        await createDocxBlob(
          content,
          await source.generateAsync({ type: 'arraybuffer' }),
        )
      ).arrayBuffer(),
    );
    expect(
      vendorAttribute(
        elementById(
          await xmlEntry(output, 'word/footnotes.xml'),
          'footnote',
          '27',
        ),
        'token',
      ),
    ).toBe('strict-note');
    expect(
      vendorAttribute(
        elementById(
          await xmlEntry(output, 'word/comments.xml'),
          'comment',
          '42',
        ),
        'token',
      ),
    ).toBe('strict-comment');
  });

  test('drops metadata from duplicate and namespace-spoofed identities', async () => {
    const content = documentContent();
    const seed = await createDocxBlob(content);
    const source = await JSZip.loadAsync(await seed.arrayBuffer());

    const footnotes = await xmlEntry(source, 'word/footnotes.xml');
    const footnote = elementById(footnotes, 'footnote', '27');
    decorate(footnotes, footnote, 'ambiguous-note');
    footnote.parentNode?.append(footnote.cloneNode(true));
    source.file(
      'word/footnotes.xml',
      new XMLSerializer().serializeToString(footnotes),
    );

    const endnotes = await xmlEntry(source, 'word/endnotes.xml');
    const endnote = elementById(endnotes, 'endnote', '41');
    decorate(endnotes, endnote, 'spoofed-note');
    removeWordId(endnote);
    endnote.setAttributeNS(VENDOR_NAMESPACE, 'vendor:id', '41');
    source.file(
      'word/endnotes.xml',
      new XMLSerializer().serializeToString(endnotes),
    );

    const comments = await xmlEntry(source, 'word/comments.xml');
    const comment = elementById(comments, 'comment', '42');
    decorate(comments, comment, 'ambiguous-comment');
    comment.parentNode?.append(comment.cloneNode(true));
    source.file(
      'word/comments.xml',
      new XMLSerializer().serializeToString(comments),
    );

    const output = await JSZip.loadAsync(
      await (
        await createDocxBlob(
          content,
          await source.generateAsync({ type: 'arraybuffer' }),
        )
      ).arrayBuffer(),
    );
    expect(
      vendorAttribute(
        elementById(
          await xmlEntry(output, 'word/footnotes.xml'),
          'footnote',
          '27',
        ),
        'token',
      ),
    ).toBeNull();
    expect(
      vendorAttribute(
        elementById(
          await xmlEntry(output, 'word/endnotes.xml'),
          'endnote',
          '41',
        ),
        'token',
      ),
    ).toBeNull();
    expect(
      vendorAttribute(
        elementById(
          await xmlEntry(output, 'word/comments.xml'),
          'comment',
          '42',
        ),
        'token',
      ),
    ).toBeNull();
  });

  test('drops malformed durable IDs and unsupported modern comment sidecars', async () => {
    const content = documentContent();
    const seed = await createDocxBlob(content);
    const source = await JSZip.loadAsync(await seed.arrayBuffer());
    source.file(
      'word/commentsIds.xml',
      `<w16cid:commentsIds xmlns:w16cid="${COMMENTS_IDS_NAMESPACE}" xmlns:fake="urn:a3s:fake"><w16cid:commentId fake:paraId="0000002B" fake:durableId="11223344"/></w16cid:commentsIds>`,
    );
    source.file(
      'word/commentsExtensible.xml',
      '<w16cex:commentsExtensible xmlns:w16cex="http://schemas.microsoft.com/office/word/2018/wordml/cex"/>',
    );
    source.file(
      'word/people.xml',
      '<w15:people xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"/>',
    );
    source.file(
      'word/vendorCommentsIds.xml',
      `<w16cid:commentsIds xmlns:w16cid="${COMMENTS_IDS_NAMESPACE}"/>`,
    );
    await appendRegistrations(source);
    const sourceBytes = await source.generateAsync({ type: 'arraybuffer' });
    const compatibility = await analyzeDocxCompatibility(
      new File([sourceBytes], 'modern-comments.docx'),
      [],
    );
    expect(compatibility.issues).toContainEqual(
      expect.objectContaining({ code: 'docx.comments.modern-sidecars' }),
    );

    const output = await JSZip.loadAsync(
      await (await createDocxBlob(content, sourceBytes)).arrayBuffer(),
    );
    expect(output.file('word/commentsIds.xml')).toBeNull();
    expect(output.file('word/commentsExtensible.xml')).toBeNull();
    expect(output.file('word/people.xml')).toBeNull();
    expect(output.file('word/vendorCommentsIds.xml')).toBeNull();
    const relationships = await xmlEntry(
      output,
      'word/_rels/document.xml.rels',
    );
    expect(
      descendants(relationships, 'Relationship').some((item) =>
        /commentsIds|commentsExtensible|people/.test(
          attribute(item, 'Type') ?? '',
        ),
      ),
    ).toBe(false);
  });
});

function documentContent(): WorkDocumentContent {
  return {
    type: 'document',
    pageSize: 'a4',
    html: [
      '<section data-document-section="true">',
      '<p>Text<sup data-document-note-reference="true" data-note-kind="footnote" data-note-id="docx-footnote-27" data-note-number="1">1</sup>',
      '<sup data-document-note-reference="true" data-note-kind="endnote" data-note-id="docx-endnote-41" data-note-number="1">1</sup></p>',
      '<p><span data-document-comment="true" data-comment-id="docx-comment-42">Commented</span></p>',
      '<aside data-document-note="true" data-note-kind="footnote" data-note-id="docx-footnote-27"><p>Foot</p></aside>',
      '<aside data-document-note="true" data-note-kind="endnote" data-note-id="docx-endnote-41"><p>End</p></aside>',
      '</section>',
    ].join(''),
    comments: [
      {
        id: 'docx-comment-42',
        author: 'Reviewer',
        date: '2026-08-12T00:00:00.000Z',
        text: 'Comment',
        resolved: true,
      },
    ],
  };
}

function decorate(document: Document, element: Element, token: string): void {
  const root = document.documentElement;
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
  element.setAttributeNS(VENDOR_NAMESPACE, 'vendor:token', token);
  element.append(document.createElementNS(VENDOR_NAMESPACE, 'vendor:passive'));
}

function removeWordId(element: Element): void {
  for (const item of Array.from(element.attributes)) {
    if (
      xmlAttributeLocalName(item) === 'id' &&
      [WORD_NAMESPACE, STRICT_WORD_NAMESPACE].includes(
        xmlAttributeNamespace(element, item) ?? '',
      )
    ) {
      element.removeAttributeNode(item);
    }
  }
}

async function appendRegistrations(archive: JSZip): Promise<void> {
  await appendBeforeClose(
    archive,
    '[Content_Types].xml',
    [
      '<Override PartName="/word/commentsIds.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"/>',
      '<Override PartName="/word/commentsExtensible.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml"/>',
      '<Override PartName="/word/people.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml"/>',
      '<Override PartName="/word/vendorCommentsIds.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml"/>',
    ].join(''),
    '</Types>',
  );
  await appendBeforeClose(
    archive,
    'word/_rels/document.xml.rels',
    [
      '<Relationship Id="rIdA3S1" Type="http://schemas.microsoft.com/office/2016/09/relationships/commentsIds" Target="commentsIds.xml"/>',
      '<Relationship Id="rIdA3S2" Type="http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible" Target="commentsExtensible.xml"/>',
      '<Relationship Id="rIdA3S3" Type="http://schemas.microsoft.com/office/2011/relationships/people" Target="people.xml"/>',
      '<Relationship Id="rIdA3S4" Type="http://schemas.microsoft.com/office/2016/09/relationships/commentsIds" Target="vendorCommentsIds.xml"/>',
    ].join(''),
    '</Relationships>',
  );
}

async function appendBeforeClose(
  archive: JSZip,
  path: string,
  value: string,
  closingTag: string,
): Promise<void> {
  const source = await archive.file(path)?.async('string');
  if (!source) throw new Error(`Missing ${path}.`);
  archive.file(path, source.replace(closingTag, `${value}${closingTag}`));
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

function vendorAttribute(element: Element, localName: string): string | null {
  return (
    Array.from(element.attributes).find(
      (item) =>
        xmlAttributeLocalName(item) === localName &&
        xmlAttributeNamespace(element, item) === VENDOR_NAMESPACE,
    )?.value ?? null
  );
}

function utf16Xml(source: string): Uint8Array {
  const normalized = source.replace(/^<\?xml[^>]*\?>/, '').replace(/^\s+/, '');
  const value = `<?xml version="1.0" encoding="UTF-16"?>${normalized}`;
  const bytes = new Uint8Array(2 + value.length * 2);
  bytes[0] = 0xff;
  bytes[1] = 0xfe;
  for (let index = 0; index < value.length; index += 1) {
    const code = value.charCodeAt(index);
    bytes[2 + index * 2] = code & 0xff;
    bytes[3 + index * 2] = code >> 8;
  }
  return bytes;
}

async function xmlEntry(archive: JSZip, path: string): Promise<Document> {
  const source = await archive.file(path)?.async('uint8array');
  if (!source) throw new Error(`Missing ${path}.`);
  const text =
    source[0] === 0xff && source[1] === 0xfe
      ? new TextDecoder('utf-16le').decode(source.subarray(2))
      : new TextDecoder().decode(source);
  return parseXml(text, path);
}
