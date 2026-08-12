import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import { preserveDocxNoteCommentHyperlinks } from '../src/internal/features/work/work-docx-note-comment-hyperlink-preservation';
import {
  descendants,
  directChildren,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';
import { decodeXmlBytes } from '../src/internal/features/work/work-ooxml-xml';
import type { WorkDocumentContent } from '../src/internal/features/work/work-types';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from '../src/internal/features/work/work-docx-settings-xml';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STRICT_WORD_NAMESPACE =
  'http://purl.oclc.org/ooxml/wordprocessingml/main';
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const STRICT_RELATIONSHIP_NAMESPACE =
  'http://purl.oclc.org/ooxml/officeDocument/relationships';
const PACKAGE_RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/relationships';
const HYPERLINK_RELATIONSHIP = `${RELATIONSHIP_NAMESPACE}/hyperlink`;
const STRICT_HYPERLINK_RELATIONSHIP = `${STRICT_RELATIONSHIP_NAMESPACE}/hyperlink`;
const COMMENTS_PATH = 'word/comments.xml';
const COMMENTS_RELATIONSHIPS_PATH = 'word/_rels/comments.xml.rels';

describe('DOCX note and comment hyperlink boundaries', () => {
  test.each([
    { name: 'missing relationship', relationship: null },
    {
      name: 'wrong relationship type',
      relationship: relationshipXml(
        'rIdBoundary',
        'https://a3s.dev/wrong-type',
        `${RELATIONSHIP_NAMESPACE}/image`,
      ),
    },
    {
      name: 'internal target mode',
      relationship: relationshipXml(
        'rIdBoundary',
        'https://a3s.dev/internal-mode',
        HYPERLINK_RELATIONSHIP,
        'Internal',
      ),
    },
    {
      name: 'unsafe scheme',
      relationship: relationshipXml(
        'rIdBoundary',
        'javascript:alert(1)',
        HYPERLINK_RELATIONSHIP,
      ),
    },
    {
      name: 'relative target',
      relationship: relationshipXml(
        'rIdBoundary',
        '../unsafe-relative',
        HYPERLINK_RELATIONSHIP,
      ),
    },
    {
      name: 'duplicate relationship ID',
      relationship: relationshipsXml([
        relationshipElement('rIdBoundary', 'https://a3s.dev/first'),
        relationshipElement('rIdBoundary', 'https://a3s.dev/second'),
      ]),
    },
    {
      name: 'unexpected relationship child',
      relationship: relationshipsXml([
        relationshipElement('rIdBoundary', 'https://a3s.dev/valid'),
        `<Unexpected xmlns="${PACKAGE_RELATIONSHIP_NAMESPACE}"/>`,
      ]),
    },
  ])('fails closed for $name', async ({ relationship }) => {
    const content = commentContent('Boundary text');
    const source = await seedArchive(content);
    await wrapCommentRun(source, 'rIdBoundary');
    if (relationship) source.file(COMMENTS_RELATIONSHIPS_PATH, relationship);

    const output = await exportWithSource(content, source);
    const comments = await xmlEntry(output, COMMENTS_PATH);
    expect(descendants(comments, 'hyperlink')).toHaveLength(0);
    expect(
      descendants(
        await xmlEntry(output, COMMENTS_RELATIONSHIPS_PATH),
        'Relationship',
      ),
    ).toHaveLength(0);
  });

  test('rejects combined external and anchor destinations and namespace spoofing', async () => {
    for (const mutation of ['combined', 'spoofed'] as const) {
      const content = commentContent('Ambiguous destination');
      const source = await seedArchive(content);
      const comments = await wrapCommentRun(source, 'rIdBoundary');
      const hyperlink = descendants(comments, 'hyperlink')[0];
      if (mutation === 'combined') {
        hyperlink.setAttributeNS(WORD_NAMESPACE, 'w:anchor', 'local-anchor');
      } else {
        removeNamespacedAttribute(hyperlink, 'id', RELATIONSHIP_NAMESPACE);
        hyperlink.setAttributeNS('urn:a3s:spoof', 'spoof:id', 'rIdBoundary');
      }
      source.file(
        COMMENTS_PATH,
        new XMLSerializer().serializeToString(comments),
      );
      source.file(
        COMMENTS_RELATIONSHIPS_PATH,
        relationshipXml(
          'rIdBoundary',
          'https://a3s.dev/ambiguous',
          HYPERLINK_RELATIONSHIP,
        ),
      );

      const output = await exportWithSource(content, source);
      expect(
        descendants(await xmlEntry(output, COMMENTS_PATH), 'hyperlink'),
      ).toHaveLength(0);
      expect(
        descendants(
          await xmlEntry(output, COMMENTS_RELATIONSHIPS_PATH),
          'Relationship',
        ),
      ).toHaveLength(0);
    }
  });

  test('accepts a strict UTF-16 mailto relationship and rewrites it safely', async () => {
    const content = commentContent('Mail reviewer');
    const source = await seedArchive(content);
    const comments = await wrapCommentRun(source, 'strictMail');
    const hyperlink = descendants(comments, 'hyperlink')[0];
    hyperlink.setAttributeNS(WORD_NAMESPACE, 'w:tooltip', 'Email reviewer');
    const strictComments = new XMLSerializer()
      .serializeToString(comments)
      .replaceAll(WORD_NAMESPACE, STRICT_WORD_NAMESPACE)
      .replaceAll(RELATIONSHIP_NAMESPACE, STRICT_RELATIONSHIP_NAMESPACE);
    source.file(COMMENTS_PATH, utf16Xml(strictComments));
    source.file(
      COMMENTS_RELATIONSHIPS_PATH,
      utf16Xml(
        relationshipXml(
          'strictMail',
          'mailto:reviewer@example.com',
          STRICT_HYPERLINK_RELATIONSHIP,
        ),
      ),
    );

    const output = await exportWithSource(content, source);
    const outputComments = await xmlEntry(output, COMMENTS_PATH);
    const outputHyperlink = descendants(outputComments, 'hyperlink')[0];
    expect(outputHyperlink).toBeDefined();
    expect(wordAttribute(outputHyperlink, 'tooltip')).toBe('Email reviewer');
    const id = relationshipAttribute(outputHyperlink, 'id');
    expect(id).toBeTruthy();
    const relationships = await xmlEntry(output, COMMENTS_RELATIONSHIPS_PATH);
    expect(
      relationshipById(relationships, id ?? '').getAttribute('Target'),
    ).toBe('mailto:reviewer@example.com');
  });

  test('allocates a fresh relationship ID when a stable comment link collides', async () => {
    const generatedArchive = new JSZip();
    const sourceArchive = new JSZip();
    generatedArchive.file(
      COMMENTS_RELATIONSHIPS_PATH,
      relationshipXml(
        'rIdCollision',
        'https://a3s.dev/existing',
        HYPERLINK_RELATIONSHIP,
      ),
    );
    sourceArchive.file(
      COMMENTS_RELATIONSHIPS_PATH,
      relationshipXml(
        'rIdCollision',
        'https://a3s.dev/intended',
        HYPERLINK_RELATIONSHIP,
      ),
    );
    const generated = parseXml(commentsXml('<w:r><w:t>Collision</w:t></w:r>'));
    const source = parseXml(
      commentsXml(
        '<w:hyperlink r:id="rIdCollision"><w:r><w:t>Collision</w:t></w:r></w:hyperlink>',
      ),
    );
    const generatedScope = descendants(generated, 'comment')[0];
    const sourceScope = descendants(source, 'comment')[0];

    await preserveDocxNoteCommentHyperlinks(
      generatedArchive,
      sourceArchive,
      generated,
      source,
      [{ generated: generatedScope, source: sourceScope }],
      'comment',
      COMMENTS_PATH,
    );

    const hyperlink = descendants(generated, 'hyperlink')[0];
    const rewrittenId = relationshipAttribute(hyperlink, 'id');
    expect(rewrittenId).toBeTruthy();
    expect(rewrittenId).not.toBe('rIdCollision');
    const relationships = await xmlEntry(
      generatedArchive,
      COMMENTS_RELATIONSHIPS_PATH,
    );
    expect(
      relationshipById(relationships, 'rIdCollision').getAttribute('Target'),
    ).toBe('https://a3s.dev/existing');
    expect(
      relationshipById(relationships, rewrittenId ?? '').getAttribute('Target'),
    ).toBe('https://a3s.dev/intended');
  });
});

function commentContent(text: string): WorkDocumentContent {
  return {
    type: 'document',
    pageSize: 'a4',
    html: '<section data-document-section="true"><p><span data-document-comment="true" data-comment-id="docx-comment-42">Anchor</span></p></section>',
    comments: [
      {
        id: 'docx-comment-42',
        author: 'Reviewer',
        date: '2026-08-12T00:00:00.000Z',
        text,
        resolved: false,
      },
    ],
  };
}

async function wrapCommentRun(
  archive: JSZip,
  relationshipId: string,
): Promise<Document> {
  const comments = await xmlEntry(archive, COMMENTS_PATH);
  const comment = descendants(comments, 'comment').find(
    (item) => wordAttribute(item, 'id') === '42',
  );
  const run = comment
    ? descendants(comment, 'r').find((item) => textOf(item).length > 0)
    : null;
  if (!run?.parentNode) throw new Error('Missing comment run.');
  const hyperlink = comments.createElementNS(WORD_NAMESPACE, 'w:hyperlink');
  hyperlink.setAttributeNS(RELATIONSHIP_NAMESPACE, 'r:id', relationshipId);
  run.parentNode.insertBefore(hyperlink, run);
  hyperlink.append(run);
  archive.file(COMMENTS_PATH, new XMLSerializer().serializeToString(comments));
  return comments;
}

function commentsXml(content: string): string {
  return [
    `<w:comments xmlns:w="${WORD_NAMESPACE}" xmlns:r="${RELATIONSHIP_NAMESPACE}">`,
    `<w:comment w:id="42"><w:p>${content}</w:p></w:comment>`,
    '</w:comments>',
  ].join('');
}

function relationshipXml(
  id: string,
  target: string,
  type: string,
  targetMode = 'External',
): string {
  return relationshipsXml([relationshipElement(id, target, type, targetMode)]);
}

function relationshipElement(
  id: string,
  target: string,
  type = HYPERLINK_RELATIONSHIP,
  targetMode = 'External',
): string {
  return `<Relationship Id="${id}" Type="${type}" Target="${target}" TargetMode="${targetMode}"/>`;
}

function relationshipsXml(elements: readonly string[]): string {
  return `<Relationships xmlns="${PACKAGE_RELATIONSHIP_NAMESPACE}">${elements.join('')}</Relationships>`;
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

function relationshipById(document: Document, id: string): Element {
  const matches = directChildren(
    document.documentElement,
    'Relationship',
  ).filter((item) => item.getAttribute('Id') === id);
  if (matches.length !== 1) throw new Error(`Missing relationship ${id}.`);
  return matches[0];
}

function removeNamespacedAttribute(
  element: Element,
  localName: string,
  namespace: string,
): void {
  for (const item of Array.from(element.attributes)) {
    if (
      xmlAttributeLocalName(item) === localName &&
      xmlAttributeNamespace(element, item) === namespace
    ) {
      element.removeAttributeNode(item);
    }
  }
}

function wordAttribute(element: Element | undefined, localName: string) {
  return namespacedAttribute(
    element,
    localName,
    new Set([WORD_NAMESPACE, STRICT_WORD_NAMESPACE]),
  );
}

function relationshipAttribute(
  element: Element | undefined,
  localName: string,
) {
  return namespacedAttribute(
    element,
    localName,
    new Set([RELATIONSHIP_NAMESPACE, STRICT_RELATIONSHIP_NAMESPACE]),
  );
}

function namespacedAttribute(
  element: Element | undefined,
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

function textOf(scope: ParentNode): string {
  return descendants(scope, 't')
    .map((item) => item.textContent ?? '')
    .join('');
}

async function xmlEntry(archive: JSZip, path: string): Promise<Document> {
  const entry = archive.file(path);
  if (!entry) throw new Error(`Missing ${path}.`);
  return parseXml(decodeXmlBytes(await entry.async('uint8array'), path), path);
}

function utf16Xml(source: string): Uint8Array {
  const normalized = source
    .replace(/^<\?xml[^>]*\?>/u, '')
    .replace(/^\s+/u, '');
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
