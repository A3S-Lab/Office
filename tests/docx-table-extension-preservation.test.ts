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
const VENDOR_NAMESPACE = 'urn:a3s:test:table-extension';

describe('DOCX table extension preservation', () => {
  test('retains scoped passive metadata across nested Strict OOXML tables', async () => {
    const content = documentContent();
    const seed = await createDocxBlob(content);
    const source = await JSZip.loadAsync(await seed.arrayBuffer());
    const body = await xmlEntry(source, 'word/document.xml');
    declareIgnorableVendor(body.documentElement);

    const outerRow = rowByIdentity(body, '1A2B3C4D', '1A2B3C4E');
    const outerTable = owningAncestor(outerRow, 'tbl');
    const outerCell = cellByParagraphIdentity(body, '2A2B3C4D');
    decorateScope(outerTable, 'tblPr', 'outer-table');
    decorateScope(outerRow, 'trPr', 'outer-row');
    decorateScope(outerCell, 'tcPr', 'outer-cell');

    const innerRow = rowByIdentity(body, '4A2B3C4D', '4A2B3C4E');
    const innerTable = owningAncestor(innerRow, 'tbl');
    const innerCell = cellByParagraphIdentity(body, '5A2B3C4D');
    decorateScope(innerTable, 'tblPr', 'inner-table');
    decorateScope(innerRow, 'trPr', 'inner-row');
    decorateScope(innerCell, 'tcPr', 'inner-cell');

    const sourceTableWidth = directChild(
      requiredChild(outerTable, 'tblPr'),
      'tblW',
    );
    sourceTableWidth?.setAttribute('w:w', '42');
    const sourceHeight = directChild(
      requiredChild(outerRow, 'trPr'),
      'trHeight',
    );
    sourceHeight?.setAttribute('w:val', '120');
    const sourceShading = directChild(requiredChild(outerCell, 'tcPr'), 'shd');
    sourceShading?.setAttribute('w:fill', '000000');
    source.file('word/document.xml', strictUtf16(body));

    const sourceHeader = await partWithRowIdentity(
      source,
      /^word\/header\d*\.xml$/i,
      '6A2B3C4D',
      '6A2B3C4E',
    );
    declareIgnorableVendor(sourceHeader.document.documentElement);
    const headerRow = rowByIdentity(
      sourceHeader.document,
      '6A2B3C4D',
      '6A2B3C4E',
    );
    decorateScope(owningAncestor(headerRow, 'tbl'), 'tblPr', 'header-table');
    decorateScope(headerRow, 'trPr', 'header-row');
    decorateScope(
      cellByParagraphIdentity(sourceHeader.document, '7A2B3C4D'),
      'tcPr',
      'header-cell',
    );
    source.file(sourceHeader.path, serializeXml(sourceHeader.document));

    const result = await createDocxBlob(
      content,
      await source.generateAsync({ type: 'arraybuffer' }),
    );
    const archive = await JSZip.loadAsync(await result.arrayBuffer());
    const output = await xmlEntry(archive, 'word/document.xml');
    const outputOuterRow = rowByIdentity(output, '1A2B3C4D', '1A2B3C4E');
    const outputOuterTable = owningAncestor(outputOuterRow, 'tbl');
    const outputOuterCell = cellByParagraphIdentity(output, '2A2B3C4D');
    expect(directChildren(outputOuterTable, 'tr')).toHaveLength(1);
    expect(directChildren(outputOuterRow, 'tc')).toHaveLength(2);
    assertPreservedScope(outputOuterTable, 'tblPr', 'outer-table');
    assertPreservedScope(outputOuterRow, 'trPr', 'outer-row');
    assertPreservedScope(outputOuterCell, 'tcPr', 'outer-cell');

    const outputInnerRow = rowByIdentity(output, '4A2B3C4D', '4A2B3C4E');
    assertPreservedScope(
      owningAncestor(outputInnerRow, 'tbl'),
      'tblPr',
      'inner-table',
    );
    assertPreservedScope(outputInnerRow, 'trPr', 'inner-row');
    assertPreservedScope(
      cellByParagraphIdentity(output, '5A2B3C4D'),
      'tcPr',
      'inner-cell',
    );

    expect(
      attribute(
        directChild(requiredChild(outputOuterTable, 'tblPr'), 'tblW') ??
          outputOuterTable,
        'w',
      ),
    ).not.toBe('42');
    expect(
      attribute(
        directChild(requiredChild(outputOuterRow, 'trPr'), 'trHeight') ??
          outputOuterRow,
        'val',
      ),
    ).toBe('600');
    expect(
      attribute(
        directChild(requiredChild(outputOuterCell, 'tcPr'), 'shd') ??
          outputOuterCell,
        'fill',
      ),
    ).toBe('DDEEFF');
    await expect(
      archive.file('word/document.xml')?.async('text'),
    ).resolves.toContain('encoding="UTF-8"');

    const outputHeader = await partWithRowIdentity(
      archive,
      /^word\/header\d*\.xml$/i,
      '6A2B3C4D',
      '6A2B3C4E',
    );
    const outputHeaderRow = rowByIdentity(
      outputHeader.document,
      '6A2B3C4D',
      '6A2B3C4E',
    );
    assertPreservedScope(
      owningAncestor(outputHeaderRow, 'tbl'),
      'tblPr',
      'header-table',
    );
    assertPreservedScope(outputHeaderRow, 'trPr', 'header-row');
    assertPreservedScope(
      cellByParagraphIdentity(outputHeader.document, '7A2B3C4D'),
      'tcPr',
      'header-cell',
    );
  });

  test('drops row metadata after a text-version change without discarding stable table and cell scopes', async () => {
    const content = documentContent();
    const seed = await createDocxBlob(content);
    const source = await JSZip.loadAsync(await seed.arrayBuffer());
    const body = await xmlEntry(source, 'word/document.xml');
    declareIgnorableVendor(body.documentElement);
    const row = rowByIdentity(body, '1A2B3C4D', '1A2B3C4E');
    const table = owningAncestor(row, 'tbl');
    const cell = cellByParagraphIdentity(body, '2A2B3C4D');
    decorateScope(table, 'tblPr', 'stable-table');
    decorateScope(row, 'trPr', 'edited-row');
    decorateScope(cell, 'tcPr', 'stable-cell');
    source.file('word/document.xml', serializeXml(body));

    const result = await createDocxBlob(
      documentContent('1A2B3C4F'),
      await source.generateAsync({ type: 'arraybuffer' }),
    );
    const archive = await JSZip.loadAsync(await result.arrayBuffer());
    const output = await xmlEntry(archive, 'word/document.xml');
    const outputRow = rowByIdentity(output, '1A2B3C4D', '1A2B3C4F');
    expect(vendorValue(outputRow, 'scopeMeta')).toBeNull();
    assertPreservedScope(
      owningAncestor(outputRow, 'tbl'),
      'tblPr',
      'stable-table',
    );
    assertPreservedScope(
      cellByParagraphIdentity(output, '2A2B3C4D'),
      'tcPr',
      'stable-cell',
    );
  });

  test('fails closed when native row and paragraph identities are ambiguous', async () => {
    const content = documentContent();
    const seed = await createDocxBlob(content);
    const source = await JSZip.loadAsync(await seed.arrayBuffer());
    const body = await xmlEntry(source, 'word/document.xml');
    declareIgnorableVendor(body.documentElement);
    const row = rowByIdentity(body, '1A2B3C4D', '1A2B3C4E');
    const table = owningAncestor(row, 'tbl');
    const cell = cellByParagraphIdentity(body, '2A2B3C4D');
    decorateScope(table, 'tblPr', 'ambiguous-table');
    decorateScope(row, 'trPr', 'ambiguous-row');
    decorateScope(cell, 'tcPr', 'ambiguous-cell');
    row.parentNode?.insertBefore(row.cloneNode(true), row.nextSibling);
    source.file('word/document.xml', serializeXml(body));

    const result = await createDocxBlob(
      content,
      await source.generateAsync({ type: 'arraybuffer' }),
    );
    const archive = await JSZip.loadAsync(await result.arrayBuffer());
    const output = await xmlEntry(archive, 'word/document.xml');
    const outputRow = rowByIdentity(output, '1A2B3C4D', '1A2B3C4E');
    expect(
      vendorValue(owningAncestor(outputRow, 'tbl'), 'scopeMeta'),
    ).toBeNull();
    expect(vendorValue(outputRow, 'scopeMeta')).toBeNull();
    expect(
      vendorValue(cellByParagraphIdentity(output, '2A2B3C4D'), 'scopeMeta'),
    ).toBeNull();
  });
});

function documentContent(outerRowTextId = '1A2B3C4E') {
  const headerHtml = [
    '<table><tbody>',
    '<tr data-office-row-id="6A2B3C4D" data-office-row-text-id="6A2B3C4E">',
    '<th colspan="2">',
    paragraphHtml('Header cell', '7A2B3C4D', '7A2B3C4E'),
    '</th></tr></tbody></table>',
  ].join('');
  return {
    type: 'document' as const,
    pageSize: 'a4' as const,
    html: [
      '<table data-office-table-width-type="pixels" data-office-table-width="300"><tbody>',
      `<tr data-office-row-id="1A2B3C4D" data-office-row-text-id="${outerRowTextId}" data-office-row-height="40" data-office-row-height-rule="exact">`,
      '<td data-office-cell-fill="#DDEEFF">',
      paragraphHtml('Outer alpha', '2A2B3C4D', '2A2B3C4E'),
      '<table><tbody>',
      '<tr data-office-row-id="4A2B3C4D" data-office-row-text-id="4A2B3C4E" data-office-cant-split="true"><td data-office-cell-fill="#F0F0F0">',
      paragraphHtml('Nested', '5A2B3C4D', '5A2B3C4E'),
      '</td></tr></tbody></table>',
      '</td><td>',
      paragraphHtml('Outer beta', '3A2B3C4D', '3A2B3C4E'),
      '</td></tr></tbody></table>',
    ].join(''),
    pageChrome: {
      differentFirstPage: false,
      differentOddEvenPages: false,
      default: { headerHtml, footerHtml: '', showPageNumber: false },
      first: { headerHtml: '', footerHtml: '', showPageNumber: false },
      even: { headerHtml: '', footerHtml: '', showPageNumber: false },
    },
  };
}

function paragraphHtml(
  text: string,
  paragraphId: string,
  textId: string,
): string {
  return `<p data-office-paragraph-id="${paragraphId}" data-office-paragraph-text-id="${textId}">${text}</p>`;
}

function decorateScope(
  element: Element,
  propertyName: 'tblPr' | 'tcPr' | 'trPr',
  label: string,
): void {
  element.setAttributeNS(VENDOR_NAMESPACE, 'vnd:stable', `${label}-attribute`);
  appendVendor(element, 'scopeMeta', `${label}-element`);
  const properties = requiredChild(element, propertyName);
  properties.setAttributeNS(
    VENDOR_NAMESPACE,
    'vnd:stableProperties',
    `${label}-properties-attribute`,
  );
  appendVendor(properties, 'propertiesMeta', `${label}-properties-element`);
  const unsafe = appendVendor(element, 'unsafeRelationship', 'drop-me');
  unsafe.setAttributeNS(RELATIONSHIPS_NAMESPACE, 'r:id', 'rIdUnsafe');
  const semanticAttribute = appendVendor(
    element,
    'semanticAttribute',
    'drop-me',
  );
  semanticAttribute.setAttributeNS(WORD_2010_NAMESPACE, 'w14:paraId', '7');
  const semantic = element.ownerDocument.createElementNS(
    WORD_2010_NAMESPACE,
    'w14:collapsed',
  );
  element.append(semantic);
  const alternate = alternateContent(element.ownerDocument);
  appendVendor(
    directChild(alternate, 'Choice') ?? alternate,
    'choiceMeta',
    `${label}-choice`,
  );
  element.append(alternate);
}

function assertPreservedScope(
  element: Element,
  propertyName: 'tblPr' | 'tcPr' | 'trPr',
  label: string,
): void {
  expect(vendorAttribute(element, 'stable')).toBe(`${label}-attribute`);
  expect(vendorValue(element, 'scopeMeta')).toBe(`${label}-element`);
  const properties = requiredChild(element, propertyName);
  expect(vendorAttribute(properties, 'stableProperties')).toBe(
    `${label}-properties-attribute`,
  );
  expect(vendorValue(properties, 'propertiesMeta')).toBe(
    `${label}-properties-element`,
  );
  expect(vendorValue(element, 'unsafeRelationship')).toBeNull();
  expect(vendorValue(element, 'semanticAttribute')).toBeNull();
  expect(descendants(element, 'collapsed')).toEqual([]);
  expect(
    descendants(element, 'choiceMeta').map((item) => item.textContent),
  ).toContain(`${label}-choice`);
}

function rowByIdentity(
  document: Document,
  rowId: string,
  rowTextId: string,
): Element {
  const row = descendants(document, 'tr').find(
    (item) =>
      word2010Attribute(item, 'paraId') === rowId &&
      word2010Attribute(item, 'textId') === rowTextId,
  );
  if (!row) throw new Error(`Missing table row: ${rowId}/${rowTextId}`);
  return row;
}

function cellByParagraphIdentity(document: Document, paragraphId: string) {
  const paragraph = descendants(document, 'p').find(
    (item) => word2010Attribute(item, 'paraId') === paragraphId,
  );
  if (!paragraph) throw new Error(`Missing cell paragraph: ${paragraphId}`);
  return owningAncestor(paragraph, 'tc');
}

function owningAncestor(element: Element, localName: string): Element {
  let current = element.parentElement;
  while (current) {
    if (current.localName === localName) return current;
    current = current.parentElement;
  }
  throw new Error(`Missing ${localName} ancestor.`);
}

function requiredChild(element: Element, localName: string): Element {
  const child = directChild(element, localName);
  if (!child) throw new Error(`Missing ${localName} child.`);
  return child;
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

function appendVendor(parent: Element, localName: string, value: string) {
  const element = parent.ownerDocument.createElementNS(
    VENDOR_NAMESPACE,
    `vnd:${localName}`,
  );
  element.textContent = value;
  parent.append(element);
  return element;
}

function alternateContent(document: Document): Element {
  const alternate = document.createElementNS(
    MARKUP_COMPATIBILITY_NAMESPACE,
    'mc:AlternateContent',
  );
  const choice = document.createElementNS(
    MARKUP_COMPATIBILITY_NAMESPACE,
    'mc:Choice',
  );
  choice.setAttribute('Requires', 'vnd');
  alternate.append(choice);
  alternate.append(
    document.createElementNS(MARKUP_COMPATIBILITY_NAMESPACE, 'mc:Fallback'),
  );
  return alternate;
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

async function partWithRowIdentity(
  archive: JSZip,
  pattern: RegExp,
  rowId: string,
  rowTextId: string,
): Promise<{ document: Document; path: string }> {
  for (const path of Object.keys(archive.files).filter((item) =>
    pattern.test(item),
  )) {
    const document = await xmlEntry(archive, path);
    if (
      descendants(document, 'tr').some(
        (row) =>
          word2010Attribute(row, 'paraId') === rowId &&
          word2010Attribute(row, 'textId') === rowTextId,
      )
    ) {
      return { document, path };
    }
  }
  throw new Error(`Missing table-row part: ${rowId}/${rowTextId}`);
}

function serializeXml(document: Document): string {
  return new XMLSerializer().serializeToString(document);
}

function strictUtf16(document: Document): Uint8Array {
  return utf16LittleEndian(
    new XMLSerializer()
      .serializeToString(document)
      .replaceAll(WORD_NAMESPACE, STRICT_WORD_NAMESPACE)
      .replaceAll(RELATIONSHIPS_NAMESPACE, STRICT_RELATIONSHIPS_NAMESPACE)
      .replace(
        /^\s*<\?xml[^?]*\?>/i,
        '<?xml version="1.0" encoding="UTF-16" standalone="yes"?>',
      ),
  );
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
