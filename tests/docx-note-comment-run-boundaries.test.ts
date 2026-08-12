import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import { analyzeDocxCompatibility } from '../src/internal/features/work/work-office-diagnostics';
import {
  attribute,
  descendants,
  directChild,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from '../src/internal/features/work/work-docx-settings-xml';
import type { WorkDocumentContent } from '../src/internal/features/work/work-types';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const VENDOR_NAMESPACE = 'urn:a3s:test:note-comment-run-boundary';
const RELATIONSHIPS_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const MARKUP_COMPATIBILITY_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';
const MATH_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/math';

describe('DOCX note and comment run boundaries', () => {
  test('follows a unique note paragraph across reordering', async () => {
    const original = noteContent('<p>First</p><p>Second</p>');
    const seed = await createDocxBlob(original);
    const source = await JSZip.loadAsync(await seed.arrayBuffer());
    const footnotes = await xmlEntry(source, 'word/footnotes.xml');
    declareVendor(footnotes.documentElement);
    const run = runContaining(
      elementById(footnotes, 'footnote', '27'),
      'Second',
    );
    addSmallCaps(run, footnotes);
    decorate(run, footnotes, 'moved-run');
    source.file(
      'word/footnotes.xml',
      new XMLSerializer().serializeToString(footnotes),
    );

    const output = await exportWithSource(
      noteContent('<p>Second</p><p>First</p>'),
      source,
    );
    const outputNote = elementById(
      await xmlEntry(output, 'word/footnotes.xml'),
      'footnote',
      '27',
    );
    expect(vendorAttribute(runContaining(outputNote, 'Second'), 'token')).toBe(
      'moved-run',
    );
    expect(
      vendorAttribute(runContaining(outputNote, 'First'), 'token'),
    ).toBeNull();
    expect(
      descendants(
        requiredChild(runContaining(outputNote, 'Second'), 'rPr'),
        'smallCaps',
      ),
    ).toHaveLength(1);
  });

  test('does not reconnect a note run behind an unresolved hyperlink', async () => {
    const content = noteContent('<p>Linked</p>');
    const seed = await createDocxBlob(content);
    const source = await JSZip.loadAsync(await seed.arrayBuffer());
    const footnotes = await xmlEntry(source, 'word/footnotes.xml');
    declareVendor(footnotes.documentElement);
    const run = runContaining(
      elementById(footnotes, 'footnote', '27'),
      'Linked',
    );
    addSmallCaps(run, footnotes);
    decorate(run, footnotes, 'unsafe-run');
    const hyperlink = wordElement(footnotes, 'hyperlink');
    hyperlink.setAttributeNS(RELATIONSHIPS_NAMESPACE, 'r:id', 'rIdUnsafe');
    run.parentNode?.insertBefore(hyperlink, run);
    hyperlink.append(run);
    source.file(
      'word/footnotes.xml',
      new XMLSerializer().serializeToString(footnotes),
    );

    const output = await exportWithSource(content, source);
    const outputNote = elementById(
      await xmlEntry(output, 'word/footnotes.xml'),
      'footnote',
      '27',
    );
    expect(descendants(outputNote, 'passiveRun')).toHaveLength(0);
    expect(descendants(outputNote, 'smallCaps')).toHaveLength(0);
    expect(descendants(outputNote, 'hyperlink')).toHaveLength(0);
  });

  test('keeps reply run formatting attached to its native ID after reordering', async () => {
    const seed = await createDocxBlob(commentContent(false));
    const source = await JSZip.loadAsync(await seed.arrayBuffer());
    const comments = await xmlEntry(source, 'word/comments.xml');
    declareVendor(comments.documentElement);
    const replyRun = runContaining(
      elementById(comments, 'comment', '91'),
      'Reply text',
    );
    const properties = wordElement(comments, 'rPr');
    properties.append(wordElement(comments, 'b'));
    replyRun.insertBefore(properties, replyRun.firstChild);
    decorate(replyRun, comments, 'reply-run');
    source.file(
      'word/comments.xml',
      new XMLSerializer().serializeToString(comments),
    );

    const output = await exportWithSource(commentContent(true), source);
    const outputComments = await xmlEntry(output, 'word/comments.xml');
    const outputReply = elementById(outputComments, 'comment', '91');
    const outputRun = runContaining(outputReply, 'Reply text');
    expect(vendorAttribute(outputRun, 'token')).toBe('reply-run');
    expect(descendants(requiredChild(outputRun, 'rPr'), 'b')).toHaveLength(1);
    expect(
      descendants(elementById(outputComments, 'comment', '7'), 'passiveRun'),
    ).toHaveLength(0);
  });

  test('rejects mixed semantic content instead of attaching its formatting to plain text', async () => {
    const content = commentContent(false);
    const seed = await createDocxBlob(content);
    const source = await JSZip.loadAsync(await seed.arrayBuffer());
    const comments = await xmlEntry(source, 'word/comments.xml');
    declareVendor(comments.documentElement);
    const run = runContaining(
      elementById(comments, 'comment', '42'),
      'Root text',
    );
    const properties = wordElement(comments, 'rPr');
    properties.append(wordElement(comments, 'b'));
    run.insertBefore(properties, run.firstChild);
    decorate(run, comments, 'mixed-run');
    run.append(comments.createElementNS(MATH_NAMESPACE, 'm:oMath'));
    source.file(
      'word/comments.xml',
      new XMLSerializer().serializeToString(comments),
    );
    const sourceBytes = await source.generateAsync({ type: 'arraybuffer' });
    const compatibility = await analyzeDocxCompatibility(
      new File([sourceBytes], 'mixed-comment-run.docx'),
      [],
    );
    expect(compatibility.issues).toContainEqual(
      expect.objectContaining({
        code: 'docx.comments.formatting',
        message: expect.stringContaining('unchanged direct text runs'),
      }),
    );

    const output = await exportWithSource(content, source);
    const outputComment = elementById(
      await xmlEntry(output, 'word/comments.xml'),
      'comment',
      '42',
    );
    expect(descendants(outputComment, 'passiveRun')).toHaveLength(0);
    expect(descendants(outputComment, 'b')).toHaveLength(0);
    expect(descendants(outputComment, 'oMath')).toHaveLength(0);
  });
});

function noteContent(noteHtml: string): WorkDocumentContent {
  return {
    type: 'document',
    pageSize: 'a4',
    html: [
      '<section data-document-section="true">',
      '<p>Body<sup data-document-note-reference="true" data-note-kind="footnote" data-note-id="docx-footnote-27" data-note-number="1">1</sup></p>',
      `<aside data-document-note="true" data-note-kind="footnote" data-note-id="docx-footnote-27" data-note-number="1">${noteHtml}</aside>`,
      '</section>',
    ].join(''),
  };
}

function commentContent(reordered: boolean): WorkDocumentContent {
  const first = {
    id: 'docx-comment-42',
    author: 'Reviewer',
    date: '2026-08-12T00:00:00.000Z',
    text: 'Root text',
    resolved: false,
    replies: [
      {
        id: 'docx-comment-reply-91',
        author: 'Responder',
        date: '2026-08-12T00:01:00.000Z',
        text: 'Reply text',
      },
    ],
  };
  const second = {
    id: 'docx-comment-7',
    author: 'Other',
    date: '2026-08-12T00:02:00.000Z',
    text: 'Second text',
    resolved: false,
  };
  const comments = reordered ? [second, first] : [first, second];
  return {
    type: 'document',
    pageSize: 'a4',
    html: `<section data-document-section="true">${comments
      .map(
        (comment) =>
          `<p><span data-document-comment="true" data-comment-id="${comment.id}">${comment.text}</span></p>`,
      )
      .join('')}</section>`,
    comments,
  };
}

function addSmallCaps(run: Element, document: Document): void {
  let properties = directChild(run, 'rPr');
  if (!properties) {
    properties = wordElement(document, 'rPr');
    run.insertBefore(properties, run.firstChild);
  }
  properties.append(wordElement(document, 'smallCaps'));
}

function decorate(element: Element, document: Document, token: string): void {
  element.setAttributeNS(VENDOR_NAMESPACE, 'vendor:token', token);
  element.append(
    document.createElementNS(VENDOR_NAMESPACE, 'vendor:passiveRun'),
  );
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

function wordElement(document: Document, localName: string): Element {
  return document.createElementNS(WORD_NAMESPACE, `w:${localName}`);
}

function runContaining(scope: Element, text: string): Element {
  const run = descendants(scope, 'r').find(
    (item) =>
      descendants(item, 't')
        .map((child) => child.textContent ?? '')
        .join('') === text,
  );
  if (!run) throw new Error(`Missing run containing ${text}.`);
  return run;
}

function requiredChild(element: Element, localName: string): Element {
  const child = directChild(element, localName);
  if (!child) throw new Error(`Missing ${localName}.`);
  return child;
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

async function exportWithSource(
  content: WorkDocumentContent,
  source: JSZip,
): Promise<JSZip> {
  const blob = await createDocxBlob(
    content,
    await source.generateAsync({ type: 'arraybuffer' }),
  );
  return JSZip.loadAsync(await blob.arrayBuffer());
}

async function xmlEntry(archive: JSZip, path: string): Promise<Document> {
  const source = await archive.file(path)?.async('uint8array');
  if (!source) throw new Error(`Missing ${path}.`);
  return parseXml(new TextDecoder().decode(source), path);
}
