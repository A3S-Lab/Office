import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import {
  attribute,
  descendants,
  directChild,
  directChildren,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from '../src/internal/features/work/work-docx-settings-xml';
import type {
  WorkDocumentComment,
  WorkDocumentContent,
} from '../src/internal/features/work/work-types';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STRICT_WORD_NAMESPACE =
  'http://purl.oclc.org/ooxml/wordprocessingml/main';
const VENDOR_NAMESPACE = 'urn:a3s:test:note-comment-run';
const RELATIONSHIPS_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const MARKUP_COMPATIBILITY_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';

describe('DOCX note and comment run preservation', () => {
  test('retains passive and unmodeled run properties on text-stable notes', async () => {
    const content = noteContent(
      '<p>Plain <strong>Rich</strong> tail</p>',
      'docx-footnote-27',
    );
    const seed = await createDocxBlob(content);
    const source = await JSZip.loadAsync(await seed.arrayBuffer());
    const footnotes = await xmlEntry(source, 'word/footnotes.xml');
    declareVendor(footnotes.documentElement);
    const run = runContaining(elementById(footnotes, 'footnote', '27'), 'Rich');
    const properties = directChild(run, 'rPr');
    if (!properties) throw new Error('Expected rich note run properties.');
    properties.append(wordElement(footnotes, 'smallCaps'));
    properties.append(wordValueElement(footnotes, 'color', 'FF0000'));
    decorate(properties, footnotes, 'properties');
    decorate(run, footnotes, 'run');
    const paragraph = run.parentElement;
    if (!paragraph || paragraph.localName !== 'p') {
      throw new Error('Expected the rich run paragraph.');
    }
    decorate(paragraph, footnotes, 'paragraph');
    source.file(
      'word/footnotes.xml',
      new XMLSerializer().serializeToString(footnotes),
    );

    const output = await exportWithSource(content, source);
    const outputRun = runContaining(
      elementById(
        await xmlEntry(output, 'word/footnotes.xml'),
        'footnote',
        '27',
      ),
      'Rich',
    );
    const outputProperties = directChild(outputRun, 'rPr');
    expect(directChild(outputProperties ?? outputRun, 'b')).toBeDefined();
    expect(
      directChild(outputProperties ?? outputRun, 'smallCaps'),
    ).toBeDefined();
    expect(
      directChildren(outputProperties ?? outputRun, 'color').some(
        (item) => attribute(item, 'val') === 'FF0000',
      ),
    ).toBe(false);
    expect(vendorAttribute(outputRun, 'token')).toBe('run');
    expect(vendorAttribute(outputProperties, 'token')).toBe('properties');
    expect(vendorAttribute(outputRun.parentElement ?? undefined, 'token')).toBe(
      'paragraph',
    );
    expect(descendants(outputRun, 'passiveRun')).toHaveLength(2);
    expect(descendants(outputRun, 'relationshipBound')).toHaveLength(0);
  });

  test('restores relationship-free rich runs for an unchanged comment', async () => {
    const content = commentContent('Alpha beta');
    const seed = await createDocxBlob(content);
    const source = await JSZip.loadAsync(await seed.arrayBuffer());
    const comments = await xmlEntry(source, 'word/comments.xml');
    declareVendor(comments.documentElement);
    const comment = elementById(comments, 'comment', '42');
    const paragraph = descendants(comment, 'p')[0];
    if (!paragraph) throw new Error('Expected a comment paragraph.');
    for (const run of directChildren(paragraph, 'r')) run.remove();
    paragraph.append(
      richRun(comments, 'Alpha ', ['b']),
      richRun(comments, 'beta', ['smallCaps', 'u'], true),
    );
    source.file(
      'word/comments.xml',
      new XMLSerializer().serializeToString(comments),
    );

    const output = await exportWithSource(content, source);
    const outputComment = elementById(
      await xmlEntry(output, 'word/comments.xml'),
      'comment',
      '42',
    );
    const runs = directChildren(descendants(outputComment, 'p')[0], 'r');
    expect(runs.map(runText)).toEqual(['Alpha ', 'beta']);
    expect(directChild(directChild(runs[0], 'rPr'), 'b')).toBeDefined();
    expect(directChild(directChild(runs[1], 'rPr'), 'smallCaps')).toBeDefined();
    expect(directChild(directChild(runs[1], 'rPr'), 'u')).toBeDefined();
    expect(vendorAttribute(runs[1], 'token')).toBe('comment-run');
    expect(descendants(runs[1], 'passiveRun')).toHaveLength(1);
    expect(descendants(runs[1], 'relationshipBound')).toHaveLength(0);

    const edited = await exportWithSource(
      commentContent('Edited comment'),
      source,
    );
    const editedComment = elementById(
      await xmlEntry(edited, 'word/comments.xml'),
      'comment',
      '42',
    );
    expect(descendants(editedComment, 'passiveRun')).toHaveLength(0);
    expect(descendants(editedComment, 'smallCaps')).toHaveLength(0);
    expect(descendants(editedComment, 'b')).toHaveLength(0);
  });

  test('translates safe properties from strict UTF-16 note runs', async () => {
    const content = noteContent('<p>Strict run</p>', 'docx-footnote-27');
    const seed = await createDocxBlob(content);
    const source = await JSZip.loadAsync(await seed.arrayBuffer());
    const footnotes = await xmlEntry(source, 'word/footnotes.xml');
    declareVendor(footnotes.documentElement);
    const run = runContaining(
      elementById(footnotes, 'footnote', '27'),
      'Strict run',
    );
    const properties = wordElement(footnotes, 'rPr');
    properties.append(wordElement(footnotes, 'smallCaps'));
    run.insertBefore(properties, run.firstChild);
    decorate(run, footnotes, 'strict-run');
    source.file(
      'word/footnotes.xml',
      utf16Xml(
        new XMLSerializer()
          .serializeToString(footnotes)
          .replaceAll(WORD_NAMESPACE, STRICT_WORD_NAMESPACE),
      ),
    );

    const output = await exportWithSource(content, source);
    const outputRun = runContaining(
      elementById(
        await xmlEntry(output, 'word/footnotes.xml'),
        'footnote',
        '27',
      ),
      'Strict run',
    );
    expect(wordChild(directChild(outputRun, 'rPr'), 'smallCaps')).toBeDefined();
    expect(vendorAttribute(outputRun, 'token')).toBe('strict-run');
    expect(outputRun.namespaceURI).toBe(WORD_NAMESPACE);
  });

  test('drops duplicate semantic run properties instead of choosing one', async () => {
    const content = noteContent(
      '<p>Duplicate property</p>',
      'docx-footnote-27',
    );
    const seed = await createDocxBlob(content);
    const source = await JSZip.loadAsync(await seed.arrayBuffer());
    const footnotes = await xmlEntry(source, 'word/footnotes.xml');
    const run = runContaining(
      elementById(footnotes, 'footnote', '27'),
      'Duplicate property',
    );
    const properties = wordElement(footnotes, 'rPr');
    const relationshipBoundProperty = wordElement(footnotes, 'bdr');
    relationshipBoundProperty.setAttributeNS(
      RELATIONSHIPS_NAMESPACE,
      'r:id',
      'rIdUnsafeProperty',
    );
    properties.append(
      wordElement(footnotes, 'smallCaps'),
      wordElement(footnotes, 'smallCaps'),
      relationshipBoundProperty,
    );
    run.insertBefore(properties, run.firstChild);
    source.file(
      'word/footnotes.xml',
      new XMLSerializer().serializeToString(footnotes),
    );

    const output = await exportWithSource(content, source);
    const outputRun = runContaining(
      elementById(
        await xmlEntry(output, 'word/footnotes.xml'),
        'footnote',
        '27',
      ),
      'Duplicate property',
    );
    expect(
      wordChildren(directChild(outputRun, 'rPr'), 'smallCaps'),
    ).toHaveLength(0);
    expect(wordChildren(directChild(outputRun, 'rPr'), 'bdr')).toHaveLength(0);
  });

  test('drops stale run metadata after text changes or ambiguous paragraph matching', async () => {
    const original = noteContent('<p>Original</p>', 'docx-footnote-27');
    const seed = await createDocxBlob(original);
    const source = await JSZip.loadAsync(await seed.arrayBuffer());
    const footnotes = await xmlEntry(source, 'word/footnotes.xml');
    declareVendor(footnotes.documentElement);
    decorate(
      runContaining(elementById(footnotes, 'footnote', '27'), 'Original'),
      footnotes,
      'stale',
    );
    source.file(
      'word/footnotes.xml',
      new XMLSerializer().serializeToString(footnotes),
    );

    const edited = await exportWithSource(
      noteContent('<p>Edited</p>', 'docx-footnote-27'),
      source,
    );
    expect(
      descendants(
        elementById(
          await xmlEntry(edited, 'word/footnotes.xml'),
          'footnote',
          '27',
        ),
        'passiveRun',
      ),
    ).toHaveLength(0);

    const duplicateContent = noteContent(
      '<p>Repeat</p><p>Repeat</p>',
      'docx-footnote-27',
    );
    const duplicateSeed = await createDocxBlob(duplicateContent);
    const duplicateSource = await JSZip.loadAsync(
      await duplicateSeed.arrayBuffer(),
    );
    const duplicateNotes = await xmlEntry(
      duplicateSource,
      'word/footnotes.xml',
    );
    declareVendor(duplicateNotes.documentElement);
    decorate(
      descendants(elementById(duplicateNotes, 'footnote', '27'), 'p').map(
        (paragraph) => runContaining(paragraph, 'Repeat'),
      )[0],
      duplicateNotes,
      'ambiguous',
    );
    duplicateSource.file(
      'word/footnotes.xml',
      new XMLSerializer().serializeToString(duplicateNotes),
    );
    const duplicateOutput = await exportWithSource(
      duplicateContent,
      duplicateSource,
    );
    expect(
      descendants(
        elementById(
          await xmlEntry(duplicateOutput, 'word/footnotes.xml'),
          'footnote',
          '27',
        ),
        'passiveRun',
      ),
    ).toHaveLength(0);
  });
});

function noteContent(html: string, id: string): WorkDocumentContent {
  return {
    type: 'document',
    pageSize: 'a4',
    html: [
      '<section data-document-section="true">',
      `<p>Body<sup data-document-note-reference="true" data-note-kind="footnote" data-note-id="${id}" data-note-number="1">1</sup></p>`,
      `<aside data-document-note="true" data-note-kind="footnote" data-note-id="${id}" data-note-number="1">${html}</aside>`,
      '</section>',
    ].join(''),
  };
}

function commentContent(text: string): WorkDocumentContent {
  const comments: WorkDocumentComment[] = [
    {
      id: 'docx-comment-42',
      author: 'Reviewer',
      date: '2026-08-12T00:00:00.000Z',
      text,
      resolved: false,
    },
  ];
  return {
    type: 'document',
    pageSize: 'a4',
    html: `<section data-document-section="true"><p><span data-document-comment="true" data-comment-id="docx-comment-42">Anchor</span></p></section>`,
    comments,
  };
}

async function exportWithSource(
  content: WorkDocumentContent,
  source: JSZip,
): Promise<JSZip> {
  const output = await createDocxBlob(
    content,
    await source.generateAsync({ type: 'arraybuffer' }),
  );
  return JSZip.loadAsync(await output.arrayBuffer());
}

function richRun(
  document: Document,
  text: string,
  properties: string[],
  decorated = false,
): Element {
  const run = wordElement(document, 'r');
  const runProperties = wordElement(document, 'rPr');
  for (const property of properties)
    runProperties.append(wordElement(document, property));
  const textElement = wordElement(document, 't');
  textElement.setAttributeNS(
    'http://www.w3.org/XML/1998/namespace',
    'xml:space',
    'preserve',
  );
  textElement.textContent = text;
  run.append(runProperties, textElement);
  if (decorated) decorate(run, document, 'comment-run');
  return run;
}

function decorate(element: Element, document: Document, token: string): void {
  element.setAttributeNS(VENDOR_NAMESPACE, 'vendor:token', token);
  const passive = document.createElementNS(
    VENDOR_NAMESPACE,
    'vendor:passiveRun',
  );
  passive.textContent = token;
  element.append(passive);
  const relationshipBound = document.createElementNS(
    VENDOR_NAMESPACE,
    'vendor:relationshipBound',
  );
  relationshipBound.setAttributeNS(
    RELATIONSHIPS_NAMESPACE,
    'r:id',
    'rIdUnsafe',
  );
  element.append(relationshipBound);
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

function wordValueElement(
  document: Document,
  localName: string,
  value: string,
): Element {
  const element = wordElement(document, localName);
  element.setAttributeNS(WORD_NAMESPACE, 'w:val', value);
  return element;
}

function runContaining(scope: Element, text: string): Element {
  const run = descendants(scope, 'r').find((item) => runText(item) === text);
  if (!run) throw new Error(`Missing run containing ${text}.`);
  return run;
}

function runText(run: Element): string {
  return directChildren(run, 't')
    .map((item) => item.textContent ?? '')
    .join('');
}

function wordChild(
  element: Element | undefined,
  localName: string,
): Element | undefined {
  return wordChildren(element, localName)[0];
}

function wordChildren(
  element: Element | undefined,
  localName: string,
): Element[] {
  return element
    ? Array.from(element.children).filter(
        (child) =>
          child.localName === localName &&
          child.namespaceURI === WORD_NAMESPACE,
      )
    : [];
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

function vendorAttribute(
  element: Element | undefined,
  localName: string,
): string | null {
  if (!element) return null;
  return (
    Array.from(element.attributes).find(
      (item) =>
        xmlAttributeLocalName(item) === localName &&
        xmlAttributeNamespace(element, item) === VENDOR_NAMESPACE,
    )?.value ?? null
  );
}

async function xmlEntry(archive: JSZip, path: string): Promise<Document> {
  const source = await archive.file(path)?.async('uint8array');
  if (!source) throw new Error(`Missing ${path}.`);
  return parseXml(new TextDecoder().decode(source), path);
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
