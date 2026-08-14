import { expect, test } from '@rstest/core';
import * as Y from 'yjs';
import {
  createOfficeCollaborationSession,
  readOfficeDocumentCollaboration,
  type WorkDocumentNode,
} from '../src/core';

const COMPLEX_DOCUMENT_FIXTURE_BASE64 =
  'AS2z8hkAKAETYTNzLm9mZmljZS5tZXRhZGF0YQhwcm90b2NvbAF3GGEzcy5vZmZpY2UuY29sbGFib3JhdGlvbigBE2Ezcy5vZmZpY2UubWV0YWRhdGEHdmVyc2lvbgF9ASgBE2Ezcy5vZmZpY2UubWV0YWRhdGEKYXJ0aWZhY3RJZAF3GGZpeHR1cmUtY29tcGxleC1kb2N1bWVudCgBE2Ezcy5vZmZpY2UubWV0YWRhdGEEa2luZAF3CGRvY3VtZW50KAETYTNzLm9mZmljZS5tZXRhZGF0YQtpbml0aWFsaXplZAF4CAEhYTNzLm9mZmljZS5ib290c3RyYXAuaW5pdGlhbGl6ZXJzAXceNDI0MjQzOmJyb3dzZXItY29tcGxleC1maXh0dXJlBwEbYTNzLm9mZmljZS5kb2N1bWVudC5jb250ZW50Aw9kb2N1bWVudFNlY3Rpb24HALPyGQYDCmJ1bGxldExpc3QHALPyGQcDCGxpc3RJdGVtBwCz8hkIAwlwYXJhZ3JhcGgHALPyGQkGBACz8hkKC0xpc3QgYW5jaG9yKACz8hkJC3BhcmFncmFwaElkAXcIMDAwMDAxMDEoALPyGQkGdGV4dElkAXcIMDAwMDAxMDKHs/IZCQMJcGFyYWdyYXBoBwCz8hkYBgQAs/IZGQlMaXN0IHRhaWwoALPyGRgLcGFyYWdyYXBoSWQBdwgwMDAwMDExMCgAs/IZGAZ0ZXh0SWQBdwgwMDAwMDExMYez8hkHAwV0YWJsZQcAs/IZJQMIdGFibGVSb3cHALPyGSYDCXRhYmxlQ2VsbAcAs/IZJwMJcGFyYWdyYXBoBwCz8hkoBgQAs/IZKQpPdXRlciBjZWxsKACz8hkoC3BhcmFncmFwaElkAXcIMDAwMDAyMTEoALPyGSgGdGV4dElkAXcIMDAwMDAyMTKHs/IZKAMFdGFibGUHALPyGTYDCHRhYmxlUm93BwCz8hk3Awl0YWJsZUNlbGwHALPyGTgDCXBhcmFncmFwaAcAs/IZOQYEALPyGToNTmVzdGVkIHRhcmdldCgAs/IZOQtwYXJhZ3JhcGhJZAF3CDAwMDAwMzExKACz8hk5BnRleHRJZAF3CDAwMDAwMzEyh7PyGTkDCXBhcmFncmFwaAcAs/IZSgYEALPyGUsLTmVzdGVkIHRhaWwoALPyGUoLcGFyYWdyYXBoSWQBdwgwMDAwMDMyMCgAs/IZSgZ0ZXh0SWQBdwgwMDAwMDMyMSgAs/IZNwVyb3dJZAF3CDAwMDAwMzAxKACz8hk3CXJvd1RleHRJZAF3CDAwMDAwMzAyKACz8hkmBXJvd0lkAXcIMDAwMDAyMDEoALPyGSYJcm93VGV4dElkAXcIMDAwMDAyMDIoALPyGQYCaWQBdxhkb2N1bWVudC1zZWN0aW9uLWNvbXBsZXgA';
const NATIVE_NESTED_TEXT_REPLACE_BASE64 =
  'AQSj9zYAhLPyGUcNTmF0aXZlIG5lc3RlZKiz8hlJAXcIMDAwMDAzMTOos/IZWgF3CDAwMDAwMzAzqLPyGVwBdwgwMDAwMDIwMwGz8hkEOw1JAVoBXAE=';
const NATIVE_LIST_PARAGRAPH_INSERT_BASE64 =
  'AQWj9zYQx7PyGQmz8hkYAwlwYXJhZ3JhcGgoAKP3NhALcGFyYWdyYXBoSWQBdwgwMDAwMDEwNSgAo/c2EAZ0ZXh0SWQBdwgwMDAwMDEwNgcAo/c2EAYEAKP3NhMLTmF0aXZlIGxpc3QBs/IZBDsNSQFaAVwB';
const NATIVE_NESTED_CELL_PARAGRAPH_INSERT_BASE64 =
  'AQej9zYfx7PyGTmz8hlKAwlwYXJhZ3JhcGgoAKP3Nh8LcGFyYWdyYXBoSWQBdwgwMDAwMDMxNSgAo/c2HwZ0ZXh0SWQBdwgwMDAwMDMxNgcAo/c2HwYEAKP3NiILTmF0aXZlIGNlbGyoo/c2DgF3CDAwMDAwMzA0qKP3Ng8BdwgwMDAwMDIwNAKz8hkEOw1JAVoBXAGj9zYBDgI=';
const NATIVE_NESTED_CELL_PARAGRAPH_DELETE_BASE64 =
  'AQKj9zYwqKP3Ni4BdwgwMDAwMDMwNaij9zYvAXcIMDAwMDAyMDUCs/IZBDsNSRBaAVwBo/c2Ag4CLgI=';

test('converges concurrent browser and native edits in nested list and table structures', () => {
  const browserDocument = documentFromFixture();
  const nativeDocument = documentFromFixture();
  addBrowserNestedCellParagraph(browserDocument);
  for (const update of [
    NATIVE_NESTED_TEXT_REPLACE_BASE64,
    NATIVE_LIST_PARAGRAPH_INSERT_BASE64,
    NATIVE_NESTED_CELL_PARAGRAPH_INSERT_BASE64,
    NATIVE_NESTED_CELL_PARAGRAPH_DELETE_BASE64,
  ]) {
    Y.applyUpdate(nativeDocument, decodeBase64(update));
  }

  expect(rowIdentity(nativeDocument, '00000201').rowTextId).toBe('00000205');
  expect(rowIdentity(nativeDocument, '00000301').rowTextId).toBe('00000305');
  exchangeUpdates(browserDocument, nativeDocument);

  const contents = [browserDocument, nativeDocument].map((document) =>
    readOfficeDocumentCollaboration(
      createOfficeCollaborationSession({
        artifactId: 'fixture-complex-document',
        document,
        kind: 'document',
      }),
    ),
  );
  for (const content of contents) {
    const root = content.model?.root;
    expect(root).toBeDefined();
    if (!root) throw new Error('Expected a shared Document model.');
    expect(collectNodes(root, 'bulletList')).toHaveLength(1);
    expect(collectNodes(root, 'listItem')).toHaveLength(1);
    expect(collectNodes(root, 'table')).toHaveLength(2);
    expect(collectNodes(root, 'tableRow')).toHaveLength(2);
    expect(documentTexts(root)).toEqual(
      expect.arrayContaining([
        'List anchor',
        'Native list',
        'List tail',
        'Outer cell',
        'Native nested',
        'Native cell',
        'Browser cell',
      ]),
    );
    expect(documentTexts(root)).not.toEqual(
      expect.arrayContaining(['Nested target', 'Nested tail']),
    );
    expect(paragraph(root, '00000311')).toMatchObject({
      attrs: { paragraphId: '00000311', textId: '00000313' },
    });
    expect(paragraph(root, '00000315')).toMatchObject({
      attrs: { paragraphId: '00000315', textId: '00000316' },
    });
    expect(paragraph(root, '00000330')).toMatchObject({
      attrs: { paragraphId: '00000330', textId: '00000331' },
    });
  }
  expect(contents[0]?.model?.root).toEqual(contents[1]?.model?.root);
  expect(contents[0]?.html).toBe(contents[1]?.html);
  expect(rowIdentity(browserDocument, '00000201')).toEqual(
    rowIdentity(nativeDocument, '00000201'),
  );
  expect(rowIdentity(browserDocument, '00000301')).toEqual(
    rowIdentity(nativeDocument, '00000301'),
  );
  expect(['00000205', '00000240']).toContain(
    rowIdentity(browserDocument, '00000201').rowTextId,
  );
  expect(['00000305', '00000340']).toContain(
    rowIdentity(browserDocument, '00000301').rowTextId,
  );
});

function documentFromFixture(): Y.Doc {
  const document = new Y.Doc();
  Y.applyUpdate(document, decodeBase64(COMPLEX_DOCUMENT_FIXTURE_BASE64));
  return document;
}

function addBrowserNestedCellParagraph(document: Y.Doc): void {
  const fragment = document.getXmlFragment('a3s.office.document.content');
  const outerRow = xmlElementByAttribute(fragment, 'rowId', '00000201');
  const nestedRow = xmlElementByAttribute(fragment, 'rowId', '00000301');
  const nestedCell = nestedRow?.get(0);
  if (
    !outerRow ||
    !nestedRow ||
    !(nestedCell instanceof Y.XmlElement) ||
    nestedCell.nodeName !== 'tableCell'
  ) {
    throw new Error('Expected the nested table fixture structure.');
  }
  document.transact(() => {
    const browserParagraph = new Y.XmlElement('paragraph');
    browserParagraph.setAttribute('paragraphId', '00000330');
    browserParagraph.setAttribute('textId', '00000331');
    const browserText = new Y.XmlText();
    browserText.insert(0, 'Browser cell');
    browserParagraph.insert(0, [browserText]);
    nestedCell.insert(1, [browserParagraph]);
    nestedRow.setAttribute('rowTextId', '00000340');
    outerRow.setAttribute('rowTextId', '00000240');
  }, 'browser-editor');
}

function rowIdentity(
  document: Y.Doc,
  rowId: string,
): { rowId: string; rowTextId: string } {
  const row = xmlElementByAttribute(
    document.getXmlFragment('a3s.office.document.content'),
    'rowId',
    rowId,
  );
  if (!row) throw new Error(`Expected table row '${rowId}'.`);
  return {
    rowId: row.getAttribute('rowId') ?? '',
    rowTextId: row.getAttribute('rowTextId') ?? '',
  };
}

function xmlElementByAttribute(
  root: Y.XmlFragment | Y.XmlElement,
  name: string,
  value: string,
): Y.XmlElement | null {
  for (const child of root.toArray()) {
    if (!(child instanceof Y.XmlElement)) continue;
    if (child.getAttribute(name) === value) return child;
    const descendant = xmlElementByAttribute(child, name, value);
    if (descendant) return descendant;
  }
  return null;
}

function collectNodes(
  root: WorkDocumentNode,
  type: string,
): WorkDocumentNode[] {
  const nodes = root.type === type ? [root] : [];
  for (const child of root.content ?? [])
    nodes.push(...collectNodes(child, type));
  return nodes;
}

function documentTexts(root: WorkDocumentNode): string[] {
  return collectNodes(root, 'paragraph').map((node) => nodeText(node));
}

function nodeText(node: WorkDocumentNode): string {
  if (node.text !== undefined) return node.text;
  return (node.content ?? []).map((child) => nodeText(child)).join('');
}

function paragraph(
  root: WorkDocumentNode,
  paragraphId: string,
): WorkDocumentNode {
  const result = collectNodes(root, 'paragraph').find(
    (node) => node.attrs?.paragraphId === paragraphId,
  );
  if (!result) throw new Error(`Expected paragraph '${paragraphId}'.`);
  return result;
}

function decodeBase64(value: string): Uint8Array {
  return Uint8Array.from(atob(value), (character) => character.charCodeAt(0));
}

function exchangeUpdates(first: Y.Doc, second: Y.Doc): void {
  const firstUpdate = Y.encodeStateAsUpdate(first, Y.encodeStateVector(second));
  const secondUpdate = Y.encodeStateAsUpdate(
    second,
    Y.encodeStateVector(first),
  );
  Y.applyUpdate(first, secondUpdate, 'test-network');
  Y.applyUpdate(second, firstUpdate, 'test-network');
}
