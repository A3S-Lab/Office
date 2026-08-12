import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import {
  descendants,
  directChildren,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';
import { decodeXmlBytes } from '../src/internal/features/work/work-ooxml-xml';
import type {
  WorkDocumentComment,
  WorkDocumentContent,
} from '../src/internal/features/work/work-types';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from '../src/internal/features/work/work-docx-settings-xml';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const PACKAGE_RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/relationships';
const HYPERLINK_RELATIONSHIP = `${RELATIONSHIP_NAMESPACE}/hyperlink`;
const VENDOR_NAMESPACE = 'urn:a3s:test:note-comment-hyperlink';
const MARKUP_COMPATIBILITY_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';

describe('DOCX note and comment hyperlink preservation', () => {
  test('keeps generated note destinations authoritative while retaining passive wrapper metadata', async () => {
    const href = 'https://a3s.dev/office?from=footnote';
    const content = noteContent(
      `<p>Visit <a href="${href}">Office docs</a> today.</p>`,
    );
    const source = await seedArchive(content);
    const footnotes = await xmlEntry(source, 'word/footnotes.xml');
    const hyperlink = hyperlinkContaining(footnotes, 'Office docs');
    declareVendor(footnotes.documentElement);
    decorateParagraph(footnotes, hyperlink.parentElement, 'note-paragraph');
    decorateHyperlink(footnotes, hyperlink, 'note-link');
    addSmallCaps(footnotes, descendants(hyperlink, 'r')[0]);
    setWordAttribute(hyperlink, 'tooltip', 'Open Office docs');
    const relationshipId = relationshipAttribute(hyperlink, 'id');
    expect(relationshipId).toBeTruthy();
    const relationships = await xmlEntry(
      source,
      'word/_rels/footnotes.xml.rels',
    );
    const relationship = relationshipById(relationships, relationshipId ?? '');
    relationship.setAttribute('Id', 'rIdSourceOnly');
    setRelationshipAttribute(hyperlink, 'id', 'rIdSourceOnly');
    source.file(
      'word/footnotes.xml',
      new XMLSerializer().serializeToString(footnotes),
    );
    source.file(
      'word/_rels/footnotes.xml.rels',
      new XMLSerializer().serializeToString(relationships),
    );

    const output = await exportWithSource(content, source);
    const outputNotes = await xmlEntry(output, 'word/footnotes.xml');
    const outputHyperlink = hyperlinkContaining(outputNotes, 'Office docs');
    expect(vendorAttribute(outputHyperlink, 'token')).toBe('note-link');
    expect(descendants(outputHyperlink, 'passiveLink')).toHaveLength(1);
    expect(wordAttribute(outputHyperlink, 'tooltip')).toBe('Open Office docs');
    expect(vendorAttribute(outputHyperlink.parentElement, 'token')).toBe(
      'note-paragraph',
    );
    expect(
      descendants(
        outputHyperlink.parentElement ?? outputNotes,
        'passiveParagraph',
      ),
    ).toHaveLength(1);
    expect(descendants(outputHyperlink, 'smallCaps')).toHaveLength(1);
    expect(
      await externalHyperlinkTarget(
        output,
        'word/footnotes.xml',
        outputHyperlink,
      ),
    ).toBe(href);
  });

  test('applies stable wrapper preservation to endnote relationships', async () => {
    const href = 'https://a3s.dev/endnotes';
    const content = noteContent(
      `<p><a href="${href}">Endnote docs</a></p>`,
      'endnote',
    );
    const source = await seedArchive(content);
    const endnotes = await xmlEntry(source, 'word/endnotes.xml');
    const hyperlink = hyperlinkContaining(endnotes, 'Endnote docs');
    declareVendor(endnotes.documentElement);
    decorateHyperlink(endnotes, hyperlink, 'endnote-link');
    source.file(
      'word/endnotes.xml',
      new XMLSerializer().serializeToString(endnotes),
    );

    const output = await exportWithSource(content, source);
    const outputHyperlink = hyperlinkContaining(
      await xmlEntry(output, 'word/endnotes.xml'),
      'Endnote docs',
    );
    expect(vendorAttribute(outputHyperlink, 'token')).toBe('endnote-link');
    expect(
      await externalHyperlinkTarget(
        output,
        'word/endnotes.xml',
        outputHyperlink,
      ),
    ).toBe(href);
  });

  test('does not partially restore wrapper metadata when another note link changes', async () => {
    const firstHref = 'https://a3s.dev/stable';
    const oldSecondHref = 'https://a3s.dev/old';
    const newSecondHref = 'https://a3s.dev/new';
    const source = await seedArchive(
      noteContent(
        `<p><a href="${firstHref}">Stable</a> and <a href="${oldSecondHref}">Changed</a></p>`,
      ),
    );
    const sourceNotes = await xmlEntry(source, 'word/footnotes.xml');
    declareVendor(sourceNotes.documentElement);
    decorateHyperlink(
      sourceNotes,
      hyperlinkContaining(sourceNotes, 'Stable'),
      'must-not-leak',
    );
    source.file(
      'word/footnotes.xml',
      new XMLSerializer().serializeToString(sourceNotes),
    );

    const output = await exportWithSource(
      noteContent(
        `<p><a href="${firstHref}">Stable</a> and <a href="${newSecondHref}">Changed</a></p>`,
      ),
      source,
    );
    const outputNotes = await xmlEntry(output, 'word/footnotes.xml');
    const stable = hyperlinkContaining(outputNotes, 'Stable');
    const changed = hyperlinkContaining(outputNotes, 'Changed');
    expect(vendorAttribute(stable, 'token')).toBeNull();
    expect(descendants(stable, 'passiveLink')).toHaveLength(0);
    expect(
      await externalHyperlinkTarget(output, 'word/footnotes.xml', stable),
    ).toBe(firstHref);
    expect(
      await externalHyperlinkTarget(output, 'word/footnotes.xml', changed),
    ).toBe(newSecondHref);
  });

  test('restores an unchanged external comment hyperlink and its run boundaries', async () => {
    const content = commentContent('Read docs now');
    const source = await seedArchive(content);
    await addCommentHyperlink(source, {
      commentId: '42',
      linkedText: 'docs',
      target: 'https://a3s.dev/docs',
      token: 'comment-link',
      tooltip: 'Read the docs',
    });

    const output = await exportWithSource(content, source);
    const comments = await xmlEntry(output, 'word/comments.xml');
    const comment = elementById(comments, 'comment', '42');
    const hyperlink = hyperlinkContaining(comment, 'docs');
    expect(vendorAttribute(hyperlink, 'token')).toBe('comment-link');
    expect(descendants(hyperlink, 'passiveLink')).toHaveLength(1);
    expect(wordAttribute(hyperlink, 'tooltip')).toBe('Read the docs');
    expect(descendants(hyperlink, 'smallCaps')).toHaveLength(1);
    expect(textOf(comment)).toContain('Read docs now');
    expect(
      await externalHyperlinkTarget(output, 'word/comments.xml', hyperlink),
    ).toBe('https://a3s.dev/docs');
  });

  test('restores an internal hyperlink on a reply after comment reordering', async () => {
    const content = threadedCommentContent(false);
    const source = await seedArchive(content);
    await addCommentHyperlink(source, {
      anchor: 'reply-destination',
      commentId: '91',
      linkedText: 'Reply docs',
      token: 'reply-link',
    });

    const output = await exportWithSource(threadedCommentContent(true), source);
    const comments = await xmlEntry(output, 'word/comments.xml');
    const reply = elementById(comments, 'comment', '91');
    const hyperlink = hyperlinkContaining(reply, 'Reply docs');
    expect(wordAttribute(hyperlink, 'anchor')).toBe('reply-destination');
    expect(relationshipAttribute(hyperlink, 'id')).toBeNull();
    expect(vendorAttribute(hyperlink, 'token')).toBe('reply-link');
    expect(
      descendants(elementById(comments, 'comment', '7'), 'hyperlink'),
    ).toHaveLength(0);
  });

  test('drops a source-backed comment hyperlink after its text changes', async () => {
    const source = await seedArchive(commentContent('Read docs now'));
    await addCommentHyperlink(source, {
      commentId: '42',
      linkedText: 'docs',
      target: 'https://a3s.dev/stale',
      token: 'stale-link',
    });

    const output = await exportWithSource(
      commentContent('Read updated docs now'),
      source,
    );
    const comments = await xmlEntry(output, 'word/comments.xml');
    expect(
      descendants(elementById(comments, 'comment', '42'), 'hyperlink'),
    ).toHaveLength(0);
    expect(
      await relationshipTargets(output, 'word/comments.xml'),
    ).not.toContain('https://a3s.dev/stale');
  });
});

interface CommentHyperlinkOptions {
  anchor?: string;
  commentId: string;
  linkedText: string;
  target?: string;
  token: string;
  tooltip?: string;
}

async function addCommentHyperlink(
  archive: JSZip,
  options: CommentHyperlinkOptions,
): Promise<void> {
  const comments = await xmlEntry(archive, 'word/comments.xml');
  const comment = elementById(comments, 'comment', options.commentId);
  const paragraph = descendants(comment, 'p').at(-1);
  if (!paragraph)
    throw new Error(`Missing comment paragraph ${options.commentId}.`);
  const run = descendants(paragraph, 'r').find((item) =>
    textOf(item).includes(options.linkedText),
  );
  if (!run || run.parentElement !== paragraph)
    throw new Error(`Missing simple comment run ${options.commentId}.`);
  const fullText = textOf(run);
  const start = fullText.indexOf(options.linkedText);
  if (start < 0) throw new Error(`Missing linked text ${options.linkedText}.`);
  const before = fullText.slice(0, start);
  const after = fullText.slice(start + options.linkedText.length);
  const hyperlink = wordElement(comments, 'hyperlink');
  declareVendor(comments.documentElement);
  decorateHyperlink(comments, hyperlink, options.token);
  if (options.tooltip) setWordAttribute(hyperlink, 'tooltip', options.tooltip);
  if (options.anchor) setWordAttribute(hyperlink, 'anchor', options.anchor);
  if (options.target)
    setRelationshipAttribute(hyperlink, 'id', 'rIdSourceLink');
  const linkedRun = cloneTextRun(run, options.linkedText);
  const properties = wordElement(comments, 'rPr');
  properties.append(wordElement(comments, 'smallCaps'));
  linkedRun.insertBefore(properties, linkedRun.firstChild);
  hyperlink.append(linkedRun);
  if (before) paragraph.insertBefore(cloneTextRun(run, before), run);
  paragraph.insertBefore(hyperlink, run);
  if (after) paragraph.insertBefore(cloneTextRun(run, after), run);
  run.remove();
  archive.file(
    'word/comments.xml',
    new XMLSerializer().serializeToString(comments),
  );
  if (options.target) {
    archive.file(
      'word/_rels/comments.xml.rels',
      relationshipsXml([
        {
          id: 'rIdSourceLink',
          target: options.target,
          type: HYPERLINK_RELATIONSHIP,
        },
      ]),
    );
  }
}

function noteContent(
  noteHtml: string,
  kind: 'endnote' | 'footnote' = 'footnote',
): WorkDocumentContent {
  const nativeId = kind === 'footnote' ? '27' : '41';
  return {
    type: 'document',
    pageSize: 'a4',
    html: [
      '<section data-document-section="true">',
      `<p>Body<sup data-document-note-reference="true" data-note-kind="${kind}" data-note-id="docx-${kind}-${nativeId}" data-note-number="1">1</sup></p>`,
      `<aside data-document-note="true" data-note-kind="${kind}" data-note-id="docx-${kind}-${nativeId}" data-note-number="1">${noteHtml}</aside>`,
      '</section>',
    ].join(''),
  };
}

function commentContent(text: string): WorkDocumentContent {
  return {
    type: 'document',
    pageSize: 'a4',
    html: '<section data-document-section="true"><p><span data-document-comment="true" data-comment-id="docx-comment-42">Anchor</span></p></section>',
    comments: [comment('42', text)],
  };
}

function threadedCommentContent(reordered: boolean): WorkDocumentContent {
  const root = comment('42', 'Root text');
  root.replies = [
    {
      id: 'docx-comment-reply-91',
      author: 'Responder',
      date: '2026-08-12T00:01:00.000Z',
      text: 'Reply docs',
    },
  ];
  const other = comment('7', 'Other text');
  return {
    type: 'document',
    pageSize: 'a4',
    html: [
      '<section data-document-section="true">',
      '<p><span data-document-comment="true" data-comment-id="docx-comment-42">Root</span></p>',
      '<p><span data-document-comment="true" data-comment-id="docx-comment-7">Other</span></p>',
      '</section>',
    ].join(''),
    comments: reordered ? [other, root] : [root, other],
  };
}

function comment(id: string, text: string): WorkDocumentComment {
  return {
    id: `docx-comment-${id}`,
    author: 'Reviewer',
    date: '2026-08-12T00:00:00.000Z',
    text,
    resolved: false,
  };
}

async function seedArchive(content: WorkDocumentContent): Promise<JSZip> {
  const blob = await createDocxBlob(content);
  return JSZip.loadAsync(await blob.arrayBuffer());
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

async function externalHyperlinkTarget(
  archive: JSZip,
  ownerPart: string,
  hyperlink: Element,
): Promise<string | null> {
  const id = relationshipAttribute(hyperlink, 'id');
  if (!id) return null;
  const path = relationshipPartPath(ownerPart);
  const document = await xmlEntry(archive, path);
  return relationshipById(document, id).getAttribute('Target');
}

async function relationshipTargets(
  archive: JSZip,
  ownerPart: string,
): Promise<string[]> {
  const entry = archive.file(relationshipPartPath(ownerPart));
  if (!entry) return [];
  const document = parseXml(
    decodeXmlBytes(await entry.async('uint8array'), ownerPart),
    ownerPart,
  );
  return descendants(document, 'Relationship').flatMap((item) => {
    const target = item.getAttribute('Target');
    return target ? [target] : [];
  });
}

function relationshipPartPath(ownerPart: string): string {
  const separator = ownerPart.lastIndexOf('/');
  const directory = separator < 0 ? '' : ownerPart.slice(0, separator + 1);
  const file = ownerPart.slice(separator + 1);
  return `${directory}_rels/${file}.rels`;
}

function relationshipsXml(
  relationships: Array<{ id: string; target: string; type: string }>,
): string {
  return [
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="${PACKAGE_RELATIONSHIP_NAMESPACE}">`,
    ...relationships.map(
      ({ id, target, type }) =>
        `<Relationship Id="${id}" Type="${type}" Target="${target.replaceAll('&', '&amp;')}" TargetMode="External"/>`,
    ),
    '</Relationships>',
  ].join('');
}

function relationshipById(document: Document, id: string): Element {
  const matches = directChildren(
    document.documentElement,
    'Relationship',
  ).filter((item) => item.getAttribute('Id') === id);
  if (matches.length !== 1) throw new Error(`Missing relationship ${id}.`);
  return matches[0];
}

function cloneTextRun(run: Element, text: string): Element {
  const clone = run.cloneNode(true) as Element;
  const textElement = descendants(clone, 't')[0];
  if (!textElement) throw new Error('Missing comment run text.');
  textElement.textContent = text;
  return clone;
}

function hyperlinkContaining(scope: ParentNode, text: string): Element {
  const hyperlink = descendants(scope, 'hyperlink').find(
    (item) => textOf(item) === text,
  );
  if (!hyperlink) throw new Error(`Missing hyperlink ${text}.`);
  return hyperlink;
}

function elementById(
  document: Document,
  localName: string,
  id: string,
): Element {
  const element = descendants(document, localName).find(
    (item) => wordAttribute(item, 'id') === id,
  );
  if (!element) throw new Error(`Missing ${localName} ${id}.`);
  return element;
}

function textOf(scope: ParentNode): string {
  return descendants(scope, 't')
    .map((item) => item.textContent ?? '')
    .join('');
}

function wordElement(document: Document, localName: string): Element {
  return document.createElementNS(WORD_NAMESPACE, `w:${localName}`);
}

function setWordAttribute(
  element: Element,
  localName: string,
  value: string,
): void {
  element.setAttributeNS(WORD_NAMESPACE, `w:${localName}`, value);
}

function setRelationshipAttribute(
  element: Element,
  localName: string,
  value: string,
): void {
  element.setAttributeNS(RELATIONSHIP_NAMESPACE, `r:${localName}`, value);
}

function wordAttribute(element: Element | undefined, localName: string) {
  return namespacedAttribute(element, localName, new Set([WORD_NAMESPACE]));
}

function relationshipAttribute(
  element: Element | undefined,
  localName: string,
) {
  return namespacedAttribute(
    element,
    localName,
    new Set([RELATIONSHIP_NAMESPACE]),
  );
}

function namespacedAttribute(
  element: Element | null | undefined,
  localName: string,
  namespaces: ReadonlySet<string>,
): string | null {
  if (!element) return null;
  return (
    Array.from(element.attributes).find(
      (item) =>
        xmlAttributeLocalName(item) === localName &&
        namespaces.has(xmlAttributeNamespace(element, item) ?? ''),
    )?.value ?? null
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

function decorateHyperlink(
  document: Document,
  hyperlink: Element,
  token: string,
): void {
  hyperlink.setAttributeNS(VENDOR_NAMESPACE, 'vendor:token', token);
  const passive = document.createElementNS(
    VENDOR_NAMESPACE,
    'vendor:passiveLink',
  );
  passive.textContent = token;
  hyperlink.append(passive);
}

function decorateParagraph(
  document: Document,
  paragraph: Element | null,
  token: string,
): void {
  if (!paragraph) throw new Error('Missing hyperlink paragraph.');
  paragraph.setAttributeNS(VENDOR_NAMESPACE, 'vendor:token', token);
  const passive = document.createElementNS(
    VENDOR_NAMESPACE,
    'vendor:passiveParagraph',
  );
  passive.textContent = token;
  paragraph.append(passive);
}

function addSmallCaps(document: Document, run: Element | undefined): void {
  if (!run) throw new Error('Missing hyperlink run.');
  let properties = directChildren(run, 'rPr')[0];
  if (!properties) {
    properties = wordElement(document, 'rPr');
    run.insertBefore(properties, run.firstChild);
  }
  properties.append(wordElement(document, 'smallCaps'));
}

function vendorAttribute(
  element: Element | null,
  localName: string,
): string | null {
  return namespacedAttribute(element, localName, new Set([VENDOR_NAMESPACE]));
}

async function xmlEntry(archive: JSZip, path: string): Promise<Document> {
  const entry = archive.file(path);
  if (!entry) throw new Error(`Missing ${path}.`);
  return parseXml(decodeXmlBytes(await entry.async('uint8array'), path), path);
}
