import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { importOfficeFile } from '../src/core';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import {
  descendants,
  parseXml,
} from '../src/internal/features/work/work-ooxml-package';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from '../src/internal/features/work/work-docx-settings-xml';

const WORD_2010_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2010/wordml';
const VENDOR_NAMESPACE = 'urn:a3s:test:spoofed-row-identity';

describe('DOCX table-row identity', () => {
  test('round-trips native row identities in body and page chrome tables', async () => {
    const content = {
      type: 'document' as const,
      html: tableHtml('Body row', '1A2B3C4D', '1A2B3C4E'),
      pageSize: 'a4' as const,
      pageChrome: {
        differentFirstPage: false,
        differentOddEvenPages: false,
        default: {
          headerHtml: tableHtml('Header row', '2A2B3C4D', '2A2B3C4E'),
          footerHtml: '',
          showPageNumber: false,
        },
        first: { headerHtml: '', footerHtml: '', showPageNumber: false },
        even: { headerHtml: '', footerHtml: '', showPageNumber: false },
      },
    };
    const blob = await createDocxBlob(content);
    const archive = await JSZip.loadAsync(await blob.arrayBuffer());
    const body = await xmlEntry(archive, 'word/document.xml');
    expect(rowIdentities(body)).toContainEqual({
      rowId: '1A2B3C4D',
      rowTextId: '1A2B3C4E',
    });
    const header = await partWithRowIdentity(
      archive,
      /^word\/header\d*\.xml$/i,
      '2A2B3C4D',
    );
    expect(rowIdentities(header)).toContainEqual({
      rowId: '2A2B3C4D',
      rowTextId: '2A2B3C4E',
    });

    const imported = await importOfficeFile(
      new File([blob], 'table-row-identities.docx', { type: blob.type }),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    expect(imported.content.html).toContain('data-office-row-id="1A2B3C4D"');
    expect(imported.content.html).toContain(
      'data-office-row-text-id="1A2B3C4E"',
    );
    expect(imported.content.pageChrome?.default.headerHtml).toContain(
      'data-office-row-id="2A2B3C4D"',
    );
  });

  test('rejects spoofed namespaces and duplicate native row IDs', async () => {
    const content = {
      type: 'document' as const,
      html: tableHtml('Body row', '3A2B3C4D', '3A2B3C4E'),
      pageSize: 'a4' as const,
    };
    const seed = await createDocxBlob(content);
    const spoofedArchive = await JSZip.loadAsync(await seed.arrayBuffer());
    const spoofedBody = await xmlEntry(spoofedArchive, 'word/document.xml');
    const spoofedRow = descendants(spoofedBody, 'tr')[0];
    for (const item of Array.from(spoofedRow.attributes)) {
      if (xmlAttributeNamespace(spoofedRow, item) === WORD_2010_NAMESPACE) {
        spoofedRow.removeAttributeNode(item);
      }
    }
    spoofedRow.setAttributeNS(VENDOR_NAMESPACE, 'vnd:paraId', '3A2B3C4D');
    spoofedRow.setAttributeNS(VENDOR_NAMESPACE, 'vnd:textId', '3A2B3C4E');
    spoofedArchive.file(
      'word/document.xml',
      new XMLSerializer().serializeToString(spoofedBody),
    );
    const spoofed = await importDocument(
      await spoofedArchive.generateAsync({ type: 'arraybuffer' }),
      'spoofed-row.docx',
    );
    expect(spoofed.html).not.toContain('data-office-row-id="3A2B3C4D"');

    const duplicateArchive = await JSZip.loadAsync(await seed.arrayBuffer());
    const duplicateBody = await xmlEntry(duplicateArchive, 'word/document.xml');
    const duplicateRow = descendants(duplicateBody, 'tr')[0];
    duplicateRow.parentNode?.insertBefore(
      duplicateRow.cloneNode(true),
      duplicateRow.nextSibling,
    );
    duplicateArchive.file(
      'word/document.xml',
      new XMLSerializer().serializeToString(duplicateBody),
    );
    const duplicated = await importDocument(
      await duplicateArchive.generateAsync({ type: 'arraybuffer' }),
      'duplicate-row.docx',
    );
    expect(duplicated.html).not.toContain('data-office-row-id="3A2B3C4D"');

    const collisionArchive = await JSZip.loadAsync(await seed.arrayBuffer());
    const collisionBody = await xmlEntry(collisionArchive, 'word/document.xml');
    const cellParagraph = descendants(collisionBody, 'p')[0];
    cellParagraph.setAttributeNS(WORD_2010_NAMESPACE, 'w14:paraId', '3A2B3C4D');
    for (const item of Array.from(cellParagraph.attributes)) {
      if (
        xmlAttributeLocalName(item) === 'textId' &&
        xmlAttributeNamespace(cellParagraph, item) === WORD_2010_NAMESPACE
      ) {
        cellParagraph.removeAttributeNode(item);
      }
    }
    collisionArchive.file(
      'word/document.xml',
      new XMLSerializer().serializeToString(collisionBody),
    );
    const collided = await importDocument(
      await collisionArchive.generateAsync({ type: 'arraybuffer' }),
      'cross-kind-row-collision.docx',
    );
    expect(collided.html).not.toContain('data-office-row-id="3A2B3C4D"');
  });
});

function tableHtml(text: string, rowId: string, rowTextId: string): string {
  return [
    '<table><tbody>',
    `<tr data-office-row-id="${rowId}" data-office-row-text-id="${rowTextId}">`,
    `<td><p>${text}</p></td>`,
    '</tr></tbody></table>',
  ].join('');
}

function rowIdentities(document: Document): Array<{
  rowId: string | null;
  rowTextId: string | null;
}> {
  return descendants(document, 'tr').map((row) => ({
    rowId: word2010Attribute(row, 'paraId'),
    rowTextId: word2010Attribute(row, 'textId'),
  }));
}

function word2010Attribute(element: Element, name: string): string | null {
  return (
    Array.from(element.attributes).find(
      (item) =>
        xmlAttributeLocalName(item) === name &&
        xmlAttributeNamespace(element, item) === WORD_2010_NAMESPACE,
    )?.value ?? null
  );
}

async function partWithRowIdentity(
  archive: JSZip,
  pathPattern: RegExp,
  rowId: string,
): Promise<Document> {
  for (const path of Object.keys(archive.files).filter((item) =>
    pathPattern.test(item),
  )) {
    const document = await xmlEntry(archive, path);
    if (rowIdentities(document).some((identity) => identity.rowId === rowId)) {
      return document;
    }
  }
  throw new Error(`Missing row identity: ${rowId}`);
}

async function xmlEntry(archive: JSZip, path: string): Promise<Document> {
  const source = await archive.file(path)?.async('text');
  if (!source) throw new Error(`Missing OOXML part: ${path}`);
  return parseXml(source, path);
}

async function importDocument(buffer: ArrayBuffer, name: string) {
  const imported = await importOfficeFile(new File([buffer], name));
  if (imported.content.type !== 'document') {
    throw new Error('Expected an imported document artifact.');
  }
  return imported.content;
}
