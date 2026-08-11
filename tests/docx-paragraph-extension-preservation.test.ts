import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { importOfficeFile } from '../src/core';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
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

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STRICT_WORD_NAMESPACE =
  'http://purl.oclc.org/ooxml/wordprocessingml/main';
const RELATIONSHIPS_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const STRICT_RELATIONSHIPS_NAMESPACE =
  'http://purl.oclc.org/ooxml/officeDocument/relationships';
const MARKUP_COMPATIBILITY_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';
const WORD_2010_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2010/wordml';
const XMLNS_NAMESPACE = 'http://www.w3.org/2000/xmlns/';
const VENDOR_NAMESPACE = 'urn:a3s:test:paragraph-extension';

describe('DOCX paragraph extension preservation', () => {
  test('retains passive metadata on stable body, list, header, and footer paragraphs', async () => {
    const content = documentContent();
    const seed = await createDocxBlob(content);
    const source = await JSZip.loadAsync(await seed.arrayBuffer());
    const body = await xmlEntry(source, 'word/document.xml');
    declareIgnorableVendor(body.documentElement);
    const bodyParagraph = paragraphByIdentity(body, '1A2B3C4D', '1A2B3C4E');
    decorateParagraph(bodyParagraph, 'body');
    const justification = directChild(
      directChild(bodyParagraph, 'pPr') ?? bodyParagraph,
      'jc',
    );
    justification?.setAttribute('val', 'left');
    decorateParagraph(
      paragraphByIdentity(body, '2A2B3C4D', '2A2B3C4E'),
      'list',
    );
    source.file('word/document.xml', strictUtf16(body));

    const headerPart = await paragraphPartByIdentity(
      source,
      /^word\/header\d*\.xml$/i,
      '3A2B3C4D',
      '3A2B3C4E',
    );
    declareIgnorableVendor(headerPart.document.documentElement);
    decorateParagraph(
      paragraphByIdentity(headerPart.document, '3A2B3C4D', '3A2B3C4E'),
      'header',
    );
    source.file(headerPart.path, serializeXml(headerPart.document));

    const footerPart = await paragraphPartByIdentity(
      source,
      /^word\/footer\d*\.xml$/i,
      '4A2B3C4D',
      '4A2B3C4E',
    );
    declareIgnorableVendor(footerPart.document.documentElement);
    decorateParagraph(
      paragraphByIdentity(footerPart.document, '4A2B3C4D', '4A2B3C4E'),
      'footer',
    );
    source.file(footerPart.path, serializeXml(footerPart.document));

    const first = await createDocxBlob(
      content,
      await source.generateAsync({ type: 'arraybuffer' }),
    );
    const archive = await JSZip.loadAsync(await first.arrayBuffer());
    const outputBody = await xmlEntry(archive, 'word/document.xml');
    assertPreservedParagraph(
      paragraphByIdentity(outputBody, '1A2B3C4D', '1A2B3C4E'),
      'body',
    );
    assertPreservedParagraph(
      paragraphByIdentity(outputBody, '2A2B3C4D', '2A2B3C4E'),
      'list',
    );
    expect(
      attribute(
        directChild(
          directChild(
            paragraphByIdentity(outputBody, '1A2B3C4D', '1A2B3C4E'),
            'pPr',
          ) ?? outputBody.documentElement,
          'jc',
        ) ?? outputBody.documentElement,
        'val',
      ),
    ).toBe('right');
    await expect(
      archive.file('word/document.xml')?.async('text'),
    ).resolves.toContain('encoding="UTF-8"');

    const outputHeader = await paragraphPartByIdentity(
      archive,
      /^word\/header\d*\.xml$/i,
      '3A2B3C4D',
      '3A2B3C4E',
    );
    assertPreservedParagraph(
      paragraphByIdentity(outputHeader.document, '3A2B3C4D', '3A2B3C4E'),
      'header',
    );
    const outputFooter = await paragraphPartByIdentity(
      archive,
      /^word\/footer\d*\.xml$/i,
      '4A2B3C4D',
      '4A2B3C4E',
    );
    assertPreservedParagraph(
      paragraphByIdentity(outputFooter.document, '4A2B3C4D', '4A2B3C4E'),
      'footer',
    );

    const second = await createDocxBlob(content, await first.arrayBuffer());
    const secondArchive = await JSZip.loadAsync(await second.arrayBuffer());
    const secondBody = await xmlEntry(secondArchive, 'word/document.xml');
    expect(
      vendorValue(
        paragraphByIdentity(secondBody, '1A2B3C4D', '1A2B3C4E'),
        'paragraphMeta',
      ),
    ).toBe('body-extension');
  });

  test('drops metadata for changed text versions and ambiguous paragraph IDs', async () => {
    const content = documentContent(false);
    const seed = await createDocxBlob(content);
    const source = await JSZip.loadAsync(await seed.arrayBuffer());
    const body = await xmlEntry(source, 'word/document.xml');
    declareIgnorableVendor(body.documentElement);
    const paragraph = paragraphByIdentity(body, '1A2B3C4D', '1A2B3C4E');
    appendVendor(paragraph, 'paragraphMeta', 'must-not-reattach');
    paragraph.parentNode?.insertBefore(
      paragraph.cloneNode(true),
      paragraph.nextSibling,
    );
    source.file('word/document.xml', serializeXml(body));

    const ambiguous = await createDocxBlob(
      content,
      await source.generateAsync({ type: 'arraybuffer' }),
    );
    const ambiguousArchive = await JSZip.loadAsync(
      await ambiguous.arrayBuffer(),
    );
    expect(
      descendants(
        await xmlEntry(ambiguousArchive, 'word/document.xml'),
        'paragraphMeta',
      ),
    ).toEqual([]);

    const uniqueSource = await JSZip.loadAsync(await seed.arrayBuffer());
    const uniqueBody = await xmlEntry(uniqueSource, 'word/document.xml');
    declareIgnorableVendor(uniqueBody.documentElement);
    appendVendor(
      paragraphByIdentity(uniqueBody, '1A2B3C4D', '1A2B3C4E'),
      'paragraphMeta',
      'must-not-follow-edited-text',
    );
    uniqueSource.file('word/document.xml', serializeXml(uniqueBody));
    const edited = await createDocxBlob(
      documentContent(false, '1A2B3C4F'),
      await uniqueSource.generateAsync({ type: 'arraybuffer' }),
    );
    const editedArchive = await JSZip.loadAsync(await edited.arrayBuffer());
    expect(
      descendants(
        await xmlEntry(editedArchive, 'word/document.xml'),
        'paragraphMeta',
      ),
    ).toEqual([]);
  });

  test('imports native paragraph identities into body and page-chrome HTML', async () => {
    const seed = await createDocxBlob(documentContent());
    const imported = await importOfficeFile(
      new File([seed], 'paragraph-identities.docx', { type: seed.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    expect(imported.content.html).toContain(
      'data-office-paragraph-id="1A2B3C4D"',
    );
    expect(imported.content.html).toContain(
      'data-office-paragraph-text-id="1A2B3C4E"',
    );
    expect(imported.content.html).toContain(
      'data-office-paragraph-id="2A2B3C4D"',
    );
    const header = imported.content.pageChrome?.default.headerHtml ?? '';
    expect(header).toContain('data-office-paragraph-id="3A2B3C4D"');
    expect(header).toContain('data-office-paragraph-text-id="3A2B3C4E"');
    const footer = imported.content.pageChrome?.default.footerHtml ?? '';
    expect(footer).toContain('data-office-paragraph-id="4A2B3C4D"');
    expect(footer).toContain('data-office-paragraph-text-id="4A2B3C4E"');
  });

  test('rejects paragraph identity attributes from spoofed namespaces', async () => {
    const seed = await createDocxBlob(documentContent(false));
    const archive = await JSZip.loadAsync(await seed.arrayBuffer());
    const body = await xmlEntry(archive, 'word/document.xml');
    const paragraph = paragraphByIdentity(body, '1A2B3C4D', '1A2B3C4E');
    for (const item of Array.from(paragraph.attributes)) {
      if (xmlAttributeNamespace(paragraph, item) === WORD_2010_NAMESPACE) {
        paragraph.removeAttributeNode(item);
      }
    }
    paragraph.setAttributeNS(VENDOR_NAMESPACE, 'vnd:paraId', '5A2B3C4D');
    paragraph.setAttributeNS(VENDOR_NAMESPACE, 'vnd:textId', '5A2B3C4E');
    expect(word2010Attribute(paragraph, 'paraId')).toBeNull();
    archive.file('word/document.xml', serializeXml(body));

    const buffer = await archive.generateAsync({ type: 'arraybuffer' });
    const imported = await importOfficeFile(new File([buffer], 'spoofed.docx'));
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    expect(imported.content.html).not.toContain('5A2B3C4D');
    expect(imported.content.html).not.toContain('5A2B3C4E');
  });
});

function documentContent(withChrome = true, bodyTextId = '1A2B3C4E') {
  return {
    type: 'document' as const,
    html: [
      paragraphHtml('Body paragraph', '1A2B3C4D', bodyTextId, 'right'),
      '<ul><li>',
      paragraphHtml('List paragraph', '2A2B3C4D', '2A2B3C4E'),
      '</li></ul>',
    ].join(''),
    pageSize: 'a4' as const,
    ...(withChrome
      ? {
          pageChrome: {
            differentFirstPage: false,
            differentOddEvenPages: false,
            default: {
              headerHtml: paragraphHtml(
                'Header paragraph',
                '3A2B3C4D',
                '3A2B3C4E',
              ),
              footerHtml: paragraphHtml(
                'Footer paragraph',
                '4A2B3C4D',
                '4A2B3C4E',
              ),
              showPageNumber: false,
            },
            first: { headerHtml: '', footerHtml: '', showPageNumber: false },
            even: { headerHtml: '', footerHtml: '', showPageNumber: false },
          },
        }
      : {}),
  };
}

function paragraphHtml(
  text: string,
  paragraphId: string,
  textId: string,
  alignment = '',
): string {
  return `<p data-office-paragraph-id="${paragraphId}" data-office-paragraph-text-id="${textId}"${
    alignment ? ` style="text-align: ${alignment}"` : ''
  }>${text}</p>`;
}

function decorateParagraph(paragraph: Element, label: string): void {
  paragraph.setAttributeNS(VENDOR_NAMESPACE, 'vnd:stable', `${label}-meta`);
  appendVendor(paragraph, 'paragraphMeta', `${label}-extension`);
  const properties = directChild(paragraph, 'pPr');
  if (properties) {
    properties.setAttributeNS(
      VENDOR_NAMESPACE,
      'vnd:propertiesStable',
      `${label}-properties`,
    );
  }
  const run = descendants(paragraph, 'r')[0];
  if (run) appendVendor(run, 'runMeta', 'drop-me');
  const unsafe = appendVendor(paragraph, 'unsafeParagraph', 'drop-me');
  unsafe.setAttributeNS(RELATIONSHIPS_NAMESPACE, 'r:id', 'rIdUnsafe');
  const semanticAttribute = appendVendor(
    paragraph,
    'semanticAttribute',
    'drop-me',
  );
  semanticAttribute.setAttributeNS(WORD_2010_NAMESPACE, 'w14:paraId', '7');
  const semantic = paragraph.ownerDocument.createElementNS(
    WORD_2010_NAMESPACE,
    'w14:collapsed',
  );
  paragraph.append(semantic);
  const passiveAlternate = alternateContent(paragraph.ownerDocument, 'vnd');
  appendVendor(
    directChild(passiveAlternate, 'Choice') ?? passiveAlternate,
    'choiceMeta',
    `${label}-choice`,
  );
  paragraph.append(passiveAlternate);
  const semanticAlternate = alternateContent(paragraph.ownerDocument, 'w14');
  appendVendor(
    directChild(semanticAlternate, 'Choice') ?? semanticAlternate,
    'requiresSemanticNamespace',
    'drop-me',
  );
  paragraph.append(semanticAlternate);
}

function assertPreservedParagraph(paragraph: Element, label: string): void {
  expect(vendorAttribute(paragraph, 'stable')).toBe(`${label}-meta`);
  expect(vendorValue(paragraph, 'paragraphMeta')).toBe(`${label}-extension`);
  const properties = directChild(paragraph, 'pPr');
  if (properties) {
    expect(vendorAttribute(properties, 'propertiesStable')).toBe(
      `${label}-properties`,
    );
  }
  expect(vendorValue(paragraph, 'unsafeParagraph')).toBeNull();
  expect(vendorValue(paragraph, 'semanticAttribute')).toBeNull();
  expect(vendorValue(paragraph, 'requiresSemanticNamespace')).toBeNull();
  expect(descendants(paragraph, 'collapsed')).toEqual([]);
  expect(descendants(paragraph, 'runMeta')).toEqual([]);
  expect(
    descendants(paragraph, 'choiceMeta').map((item) => item.textContent),
  ).toEqual([`${label}-choice`]);
}

function alternateContent(document: Document, requiredPrefix: string): Element {
  const alternate = document.createElementNS(
    MARKUP_COMPATIBILITY_NAMESPACE,
    'mc:AlternateContent',
  );
  const choice = document.createElementNS(
    MARKUP_COMPATIBILITY_NAMESPACE,
    'mc:Choice',
  );
  choice.setAttribute('Requires', requiredPrefix);
  alternate.append(choice);
  alternate.append(
    document.createElementNS(MARKUP_COMPATIBILITY_NAMESPACE, 'mc:Fallback'),
  );
  return alternate;
}

function paragraphByIdentity(
  document: Document,
  paragraphId: string,
  textId: string,
): Element {
  const paragraph = descendants(document, 'p').find(
    (item) =>
      word2010Attribute(item, 'paraId') === paragraphId &&
      word2010Attribute(item, 'textId') === textId,
  );
  if (!paragraph) {
    throw new Error(`Missing paragraph identity: ${paragraphId}/${textId}`);
  }
  return paragraph;
}

async function paragraphPartByIdentity(
  archive: JSZip,
  pathPattern: RegExp,
  paragraphId: string,
  textId: string,
): Promise<{ document: Document; path: string }> {
  for (const path of Object.keys(archive.files).filter((item) =>
    pathPattern.test(item),
  )) {
    const document = await xmlEntry(archive, path);
    if (
      descendants(document, 'p').some(
        (item) =>
          word2010Attribute(item, 'paraId') === paragraphId &&
          word2010Attribute(item, 'textId') === textId,
      )
    ) {
      return { document, path };
    }
  }
  throw new Error(`Missing paragraph part: ${paragraphId}/${textId}`);
}

function word2010Attribute(element: Element, localName: string): string | null {
  return (
    Array.from(element.attributes).find(
      (item) =>
        xmlAttributeLocalName(item) === localName &&
        xmlAttributeNamespace(element, item) === WORD_2010_NAMESPACE,
    )?.value ?? null
  );
}

function declareIgnorableVendor(root: Element): void {
  root.setAttributeNS(XMLNS_NAMESPACE, 'xmlns:vnd', VENDOR_NAMESPACE);
  const current = (attribute(root, 'mc:Ignorable') ?? '').trim();
  root.setAttributeNS(
    MARKUP_COMPATIBILITY_NAMESPACE,
    'mc:Ignorable',
    [current, 'vnd'].filter(Boolean).join(' '),
  );
}

function appendVendor(
  parent: Element,
  localName: string,
  value: string,
): Element {
  const element = parent.ownerDocument.createElementNS(
    VENDOR_NAMESPACE,
    `vnd:${localName}`,
  );
  element.textContent = value;
  parent.append(element);
  return element;
}

function vendorValue(parent: Element, localName: string): string | null {
  return (
    Array.from(parent.children).find(
      (item) =>
        item.namespaceURI === VENDOR_NAMESPACE && item.localName === localName,
    )?.textContent ?? null
  );
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

async function xmlEntry(archive: JSZip, path: string): Promise<Document> {
  const source = await archive.file(path)?.async('text');
  if (!source) throw new Error(`Missing OOXML part: ${path}`);
  return parseXml(source, path);
}

function strictUtf16(document: Document): Uint8Array {
  return utf16LittleEndian(
    serializeXml(document)
      .replaceAll(WORD_NAMESPACE, STRICT_WORD_NAMESPACE)
      .replaceAll(RELATIONSHIPS_NAMESPACE, STRICT_RELATIONSHIPS_NAMESPACE)
      .replace(
        /^\s*<\?xml[^?]*\?>/i,
        '<?xml version="1.0" encoding="UTF-16" standalone="yes"?>',
      ),
  );
}

function serializeXml(document: Document): string {
  return new XMLSerializer().serializeToString(document);
}

function utf16LittleEndian(source: string): Uint8Array {
  const bytes = new Uint8Array(2 + source.length * 2);
  bytes[0] = 0xff;
  bytes[1] = 0xfe;
  for (let index = 0; index < source.length; index += 1) {
    const codeUnit = source.charCodeAt(index);
    bytes[2 + index * 2] = codeUnit & 0xff;
    bytes[3 + index * 2] = codeUnit >>> 8;
  }
  return bytes;
}
