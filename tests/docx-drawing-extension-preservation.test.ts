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
const WORDPROCESSING_DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
const STRICT_WORDPROCESSING_DRAWING_NAMESPACE =
  'http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing';
const DRAWING_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/main';
const STRICT_DRAWING_NAMESPACE = 'http://purl.oclc.org/ooxml/drawingml/main';
const PICTURE_NAMESPACE =
  'http://schemas.openxmlformats.org/drawingml/2006/picture';
const STRICT_PICTURE_NAMESPACE = 'http://purl.oclc.org/ooxml/drawingml/picture';
const RELATIONSHIPS_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const STRICT_RELATIONSHIPS_NAMESPACE =
  'http://purl.oclc.org/ooxml/officeDocument/relationships';
const MARKUP_COMPATIBILITY_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';
const WORDPROCESSING_DRAWING_2010_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing';
const XMLNS_NAMESPACE = 'http://www.w3.org/2000/xmlns/';
const VENDOR_NAMESPACE = 'urn:a3s:test:drawing-extension';
const pixelPng =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwC' +
  'AAAAC0lEQVR42mP8/x8AAusB9Y9Z9WQAAAAASUVORK5CYII=';

describe('DOCX drawing extension preservation', () => {
  test('retains passive vendor metadata on stable body, header, and footer drawing identities', async () => {
    const content = documentContent();
    const seed = await createDocxBlob(content);
    const source = await JSZip.loadAsync(await seed.arrayBuffer());
    const body = await xmlEntry(source, 'word/document.xml');
    const bodyDrawing = drawingByIdentity(body, '1A2B3C4D', '42');
    declareIgnorableVendor(body.documentElement);
    decorateDrawing(bodyDrawing, 'body');
    const sourceExtent = directChild(bodyDrawing, 'extent');
    if (!sourceExtent) throw new Error('Generated body extent is missing.');
    sourceExtent.setAttribute('cx', '1');
    source.file('word/document.xml', strictUtf16(body));

    const sourceHeader = await drawingPartByIdentity(
      source,
      /^word\/header\d*\.xml$/i,
      '2A2B3C4D',
      '43',
    );
    const { document: header, path: headerPath } = sourceHeader;
    const headerDrawing = drawingByIdentity(header, '2A2B3C4D', '43');
    declareIgnorableVendor(header.documentElement);
    decorateDrawing(headerDrawing, 'header');
    source.file(headerPath, serializeXml(header));
    const sourceFooter = await drawingPartByIdentity(
      source,
      /^word\/footer\d*\.xml$/i,
      '3A2B3C4D',
      '44',
    );
    const footerDrawing = drawingByIdentity(
      sourceFooter.document,
      '3A2B3C4D',
      '44',
    );
    declareIgnorableVendor(sourceFooter.document.documentElement);
    decorateDrawing(footerDrawing, 'footer');
    source.file(sourceFooter.path, serializeXml(sourceFooter.document));

    const first = await createDocxBlob(
      content,
      await source.generateAsync({ type: 'arraybuffer' }),
    );
    const firstArchive = await JSZip.loadAsync(await first.arrayBuffer());
    const firstBody = await xmlEntry(firstArchive, 'word/document.xml');
    assertPreservedDrawing(
      drawingByIdentity(firstBody, '1A2B3C4D', '42'),
      'body',
    );
    expect(
      directChild(
        drawingByIdentity(firstBody, '1A2B3C4D', '42'),
        'extent',
      )?.getAttribute('cx'),
    ).not.toBe('1');
    await expect(
      firstArchive.file('word/document.xml')?.async('text'),
    ).resolves.toContain('encoding="UTF-8"');

    const firstHeader = await drawingPartByIdentity(
      firstArchive,
      /^word\/header\d*\.xml$/i,
      '2A2B3C4D',
      '43',
    );
    assertPreservedDrawing(
      drawingByIdentity(firstHeader.document, '2A2B3C4D', '43'),
      'header',
    );
    const firstFooter = await drawingPartByIdentity(
      firstArchive,
      /^word\/footer\d*\.xml$/i,
      '3A2B3C4D',
      '44',
    );
    assertPreservedDrawing(
      drawingByIdentity(firstFooter.document, '3A2B3C4D', '44'),
      'footer',
    );

    const second = await createDocxBlob(content, await first.arrayBuffer());
    const secondArchive = await JSZip.loadAsync(await second.arrayBuffer());
    const secondBody = await xmlEntry(secondArchive, 'word/document.xml');
    expect(
      vendorValue(
        drawingByIdentity(secondBody, '1A2B3C4D', '42'),
        'drawingMeta',
      ),
    ).toBe('body-extension');
  });

  test('drops drawing extensions when a source identity is ambiguous', async () => {
    const content = documentContent(false);
    const seed = await createDocxBlob(content);
    const source = await JSZip.loadAsync(await seed.arrayBuffer());
    const body = await xmlEntry(source, 'word/document.xml');
    const drawing = drawingByIdentity(body, '1A2B3C4D', '42');
    declareIgnorableVendor(body.documentElement);
    appendVendor(drawing, 'drawingMeta', 'must-not-reattach');
    drawing.parentNode?.insertBefore(
      drawing.cloneNode(true),
      drawing.nextSibling,
    );
    source.file('word/document.xml', serializeXml(body));

    const output = await createDocxBlob(
      content,
      await source.generateAsync({ type: 'arraybuffer' }),
    );
    const archive = await JSZip.loadAsync(await output.arrayBuffer());
    const outputBody = await xmlEntry(archive, 'word/document.xml');
    expect(descendants(outputBody, 'drawingMeta')).toEqual([]);
  });

  test('carries native header and footer drawing identities into editable page-chrome HTML', async () => {
    const seed = await createDocxBlob(documentContent());
    const imported = await importOfficeFile(
      new File([seed], 'header-drawing-identity.docx', { type: seed.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const header = imported.content.pageChrome?.default.headerHtml ?? '';
    expect(header).toContain('data-office-image-object-id="2A2B3C4D"');
    expect(header).toContain('data-office-image-doc-properties-id="43"');
    expect(header).toContain('data-office-image-anchor-id="2A2B3C4D"');
    expect(header).toContain('data-office-image-edit-id="1A0B0C0D"');
    const footer = imported.content.pageChrome?.default.footerHtml ?? '';
    expect(footer).toContain('data-office-image-object-id="3A2B3C4D"');
    expect(footer).toContain('data-office-image-doc-properties-id="44"');
    expect(footer).toContain('data-office-image-anchor-id="3A2B3C4D"');
    expect(footer).toContain('data-office-image-edit-id="2A0B0C0D"');
  });

  test('ignores namespace-spoofed page-chrome drawing identities', async () => {
    const seed = await createDocxBlob(documentContent());
    const source = await JSZip.loadAsync(await seed.arrayBuffer());
    const headerPart = await drawingPartByIdentity(
      source,
      /^word\/header\d*\.xml$/i,
      '2A2B3C4D',
      '43',
    );
    const anchor = drawingByIdentity(headerPart.document, '2A2B3C4D', '43');
    for (const item of Array.from(anchor.attributes)) {
      if (
        xmlAttributeNamespace(anchor, item) ===
          WORDPROCESSING_DRAWING_2010_NAMESPACE &&
        (xmlAttributeLocalName(item) === 'anchorId' ||
          xmlAttributeLocalName(item) === 'editId')
      ) {
        anchor.removeAttributeNode(item);
      }
    }
    declareIgnorableVendor(headerPart.document.documentElement);
    anchor.setAttributeNS(VENDOR_NAMESPACE, 'vnd:anchorId', '6A2B3C4D');
    anchor.setAttributeNS(VENDOR_NAMESPACE, 'vnd:editId', '7A0B0C0D');
    const nativeProperties = directChild(anchor, 'docPr');
    if (!nativeProperties)
      throw new Error('Generated drawing properties are missing.');
    const spoofedProperties = headerPart.document.createElementNS(
      VENDOR_NAMESPACE,
      'vnd:docPr',
    );
    spoofedProperties.setAttribute('id', '99');
    anchor.insertBefore(spoofedProperties, nativeProperties);
    source.file(headerPart.path, serializeXml(headerPart.document));

    const buffer = await source.generateAsync({ type: 'arraybuffer' });
    const imported = await importOfficeFile(
      new File([buffer], 'spoofed-header-drawing.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const header = imported.content.pageChrome?.default.headerHtml ?? '';
    expect(header).toContain('data-office-image-doc-properties-id="43"');
    expect(header).toContain('data-office-image-anchor-id="0000002B"');
    expect(header).not.toContain('6A2B3C4D');
    expect(header).not.toContain('7A0B0C0D');
  });
});

function documentContent(withHeader = true) {
  return {
    type: 'document' as const,
    html: `<p>Body ${imageHtml('Body image', '1A2B3C4D', 42, '0A0B0C0D')}</p>`,
    pageSize: 'a4' as const,
    ...(withHeader
      ? {
          pageChrome: {
            differentFirstPage: false,
            differentOddEvenPages: false,
            default: {
              headerHtml: `<p>${imageHtml('Header image', '2A2B3C4D', 43, '1A0B0C0D')}</p>`,
              footerHtml: `<p>${imageHtml('Footer image', '3A2B3C4D', 44, '2A0B0C0D')}</p>`,
              showPageNumber: false,
            },
            first: { headerHtml: '', footerHtml: '', showPageNumber: false },
            even: { headerHtml: '', footerHtml: '', showPageNumber: false },
          },
        }
      : {}),
  };
}

function imageHtml(
  alternativeText: string,
  objectId: string,
  docPropertiesId: number,
  editId: string,
): string {
  return [
    `<img src="${pixelPng}" alt="${alternativeText}" width="40" height="30"`,
    ` data-office-image-object-id="${objectId}"`,
    ` data-office-image-doc-properties-id="${docPropertiesId}"`,
    ` data-office-image-anchor-id="${objectId}"`,
    ` data-office-image-edit-id="${editId}">`,
  ].join('');
}

function decorateDrawing(drawing: Element, label: string): void {
  drawing.setAttributeNS(VENDOR_NAMESPACE, 'vnd:stable', `${label}-metadata`);
  appendVendor(drawing, 'drawingMeta', `${label}-extension`);
  const properties = directChild(drawing, 'docPr');
  if (!properties) throw new Error('Generated drawing properties are missing.');
  properties.setAttributeNS(
    VENDOR_NAMESPACE,
    'vnd:propertiesStable',
    `${label}-properties`,
  );
  const unsafe = appendVendor(drawing, 'unsafeDrawing', 'drop-me');
  unsafe.setAttributeNS(RELATIONSHIPS_NAMESPACE, 'r:id', 'rIdUnsafeDrawing');
  const semanticAttribute = appendVendor(
    drawing,
    'semanticAttribute',
    'drop-me',
  );
  semanticAttribute.setAttributeNS(
    WORDPROCESSING_DRAWING_2010_NAMESPACE,
    'wp14:relativeHeight',
    '7',
  );
  const semantic = drawing.ownerDocument.createElementNS(
    WORDPROCESSING_DRAWING_2010_NAMESPACE,
    'wp14:sizeRelH',
  );
  semantic.setAttribute('relativeFrom', 'page');
  drawing.append(semantic);
  const passiveAlternate = alternateContent(drawing.ownerDocument, 'vnd');
  appendVendor(
    directChild(passiveAlternate, 'Choice') ?? passiveAlternate,
    'choiceMeta',
    `${label}-choice`,
  );
  appendVendor(
    directChild(passiveAlternate, 'Fallback') ?? passiveAlternate,
    'fallbackMeta',
    `${label}-fallback`,
  );
  drawing.append(passiveAlternate);
  const semanticAlternate = alternateContent(drawing.ownerDocument, 'wp14');
  const semanticChoice = directChild(semanticAlternate, 'Choice');
  if (semanticChoice) semanticChoice.append(semantic.cloneNode(true));
  drawing.append(semanticAlternate);
  const semanticRequirement = alternateContent(drawing.ownerDocument, 'wp14');
  appendVendor(
    directChild(semanticRequirement, 'Choice') ?? semanticRequirement,
    'requiresSemanticNamespace',
    'drop-me',
  );
  drawing.append(semanticRequirement);
}

function assertPreservedDrawing(drawing: Element, label: string): void {
  expect(vendorAttribute(drawing, 'stable')).toBe(`${label}-metadata`);
  expect(vendorValue(drawing, 'drawingMeta')).toBe(`${label}-extension`);
  expect(
    vendorAttribute(directChild(drawing, 'docPr'), 'propertiesStable'),
  ).toBe(`${label}-properties`);
  expect(vendorValue(drawing, 'unsafeDrawing')).toBeNull();
  expect(vendorValue(drawing, 'semanticAttribute')).toBeNull();
  expect(vendorValue(drawing, 'requiresSemanticNamespace')).toBeNull();
  expect(descendants(drawing, 'sizeRelH')).toEqual([]);
  expect(
    descendants(drawing, 'choiceMeta').map((item) => item.textContent),
  ).toEqual([`${label}-choice`]);
  expect(
    descendants(drawing, 'fallbackMeta').map((item) => item.textContent),
  ).toEqual([`${label}-fallback`]);
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

function drawingByIdentity(
  document: Document,
  anchorId: string,
  docPropertiesId: string,
): Element {
  const drawing = descendants(document, 'anchor')
    .concat(descendants(document, 'inline'))
    .find(
      (item) =>
        attribute(item, 'anchorId') === anchorId &&
        directChild(item, 'docPr')?.getAttribute('id') === docPropertiesId,
    );
  if (!drawing) {
    throw new Error(`Missing drawing identity: ${anchorId}/${docPropertiesId}`);
  }
  return drawing;
}

async function drawingPartByIdentity(
  archive: JSZip,
  pathPattern: RegExp,
  anchorId: string,
  docPropertiesId: string,
): Promise<{ document: Document; path: string }> {
  const found: string[] = [];
  for (const path of Object.keys(archive.files).filter((item) =>
    pathPattern.test(item),
  )) {
    const document = await xmlEntry(archive, path);
    found.push(
      ...descendants(document, 'anchor')
        .concat(descendants(document, 'inline'))
        .map(
          (item) =>
            `${path}:${attribute(item, 'anchorId')}/${directChild(item, 'docPr')?.getAttribute('id')}`,
        ),
    );
    if (hasDrawingIdentity(document, anchorId, docPropertiesId)) {
      return { document, path };
    }
  }
  throw new Error(
    `Missing drawing part: ${anchorId}/${docPropertiesId}; found ${found.join(', ')}`,
  );
}

function hasDrawingIdentity(
  document: Document,
  anchorId: string,
  docPropertiesId: string,
): boolean {
  return descendants(document, 'anchor')
    .concat(descendants(document, 'inline'))
    .some(
      (item) =>
        attribute(item, 'anchorId') === anchorId &&
        directChild(item, 'docPr')?.getAttribute('id') === docPropertiesId,
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
  const source = await archive.file(path)?.async('text');
  if (!source) throw new Error(`Missing OOXML part: ${path}`);
  return parseXml(source, path);
}

function strictUtf16(document: Document): Uint8Array {
  return utf16LittleEndian(
    serializeXml(document)
      .replaceAll(WORD_NAMESPACE, STRICT_WORD_NAMESPACE)
      .replaceAll(
        WORDPROCESSING_DRAWING_NAMESPACE,
        STRICT_WORDPROCESSING_DRAWING_NAMESPACE,
      )
      .replaceAll(DRAWING_NAMESPACE, STRICT_DRAWING_NAMESPACE)
      .replaceAll(PICTURE_NAMESPACE, STRICT_PICTURE_NAMESPACE)
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
