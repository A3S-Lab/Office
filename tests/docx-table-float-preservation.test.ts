import { describe, expect, test } from '@rstest/core';
import { Editor } from '@tiptap/core';
import JSZip from 'jszip';
import {
  createArtifact,
  createArtifactBlob,
  importOfficeFile,
} from '../src/core';
import { createWorkDocumentExtensions } from '../src/internal/features/work/work-document-extensions';
import {
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
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

describe('DOCX floating table preservation', () => {
  test('preserves relationship-free w:tblpPr as opaque table metadata on untouched round-trip', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html =
      '<table><tbody><tr><td><p>Cell</p></td></tr></tbody></table>';
    const seed = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await seed.arrayBuffer());
    const document = await xmlEntry(archive, 'word/document.xml');
    const table = descendants(document, 'tbl')[0];
    const properties =
      directChild(table, 'tblPr') ??
      (() => {
        const created = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
        table.insertBefore(created, table.firstChild);
        return created;
      })();
    const float = document.createElementNS(WORD_NAMESPACE, 'w:tblpPr');
    float.setAttributeNS(WORD_NAMESPACE, 'w:horzAnchor', 'page');
    float.setAttributeNS(WORD_NAMESPACE, 'w:vertAnchor', 'text');
    float.setAttributeNS(WORD_NAMESPACE, 'w:tblpX', '720');
    float.setAttributeNS(WORD_NAMESPACE, 'w:tblpY', '360');
    float.setAttributeNS(WORD_NAMESPACE, 'w:leftFromText', '120');
    float.setAttributeNS(WORD_NAMESPACE, 'w:rightFromText', '120');
    float.setAttributeNS(WORD_NAMESPACE, 'w:topFromText', '0');
    float.setAttributeNS(WORD_NAMESPACE, 'w:bottomFromText', '0');
    float.setAttributeNS(WORD_NAMESPACE, 'w:tblOverlap', 'never');
    properties.append(float);
    archive.file(
      'word/document.xml',
      new XMLSerializer().serializeToString(document),
    );
    const source = await archive.generateAsync({ type: 'arraybuffer' });
    const imported = await importOfficeFile(
      new File([source], 'floating-table.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const importedTable = html.body.querySelector('table');
    expect(importedTable?.dataset.officeTableFloatOmml).toBeTruthy();

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      const edited = new DOMParser().parseFromString(
        editor.getHTML(),
        'text/html',
      );
      expect(
        edited.body.querySelector('table')?.dataset.officeTableFloatOmml,
      ).toBeTruthy();
      imported.content.html = editor.getHTML();
    } finally {
      editor.destroy();
    }

    const output = await JSZip.loadAsync(
      await (await createArtifactBlob(imported)).arrayBuffer(),
    );
    const exported = await xmlEntry(output, 'word/document.xml');
    const exportedFloat = directChild(
      directChild(descendants(exported, 'tbl')[0], 'tblPr'),
      'tblpPr',
    );
    expect(exportedFloat).toBeTruthy();
    expect(floatAttributes(exportedFloat!)).toMatchObject({
      horzAnchor: 'page',
      vertAnchor: 'text',
      tblpX: '720',
      tblpY: '360',
      leftFromText: '120',
      rightFromText: '120',
      topFromText: '0',
      bottomFromText: '0',
      tblOverlap: 'never',
    });
  });

  test('fails closed for relationship-bound w:tblpPr instead of preserving it', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html =
      '<table><tbody><tr><td><p>Cell</p></td></tr></tbody></table>';
    const seed = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await seed.arrayBuffer());
    const document = await xmlEntry(archive, 'word/document.xml');
    const table = descendants(document, 'tbl')[0];
    const properties =
      directChild(table, 'tblPr') ??
      (() => {
        const created = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
        table.insertBefore(created, table.firstChild);
        return created;
      })();
    const float = document.createElementNS(WORD_NAMESPACE, 'w:tblpPr');
    float.setAttributeNS(WORD_NAMESPACE, 'w:horzAnchor', 'page');
    float.setAttributeNS(RELATIONSHIP_NAMESPACE, 'r:id', 'rIdUnsafe');
    properties.append(float);
    archive.file(
      'word/document.xml',
      new XMLSerializer().serializeToString(document),
    );
    const source = await archive.generateAsync({ type: 'arraybuffer' });
    const imported = await importOfficeFile(
      new File([source], 'unsafe-floating-table.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    expect(html.body.querySelector('table')?.dataset.officeTableFloatOmml).toBe(
      undefined,
    );
    const output = await JSZip.loadAsync(
      await (await createArtifactBlob(imported)).arrayBuffer(),
    );
    const exported = await xmlEntry(output, 'word/document.xml');
    expect(
      directChild(
        directChild(descendants(exported, 'tbl')[0], 'tblPr'),
        'tblpPr',
      ),
    ).toBeUndefined();
  });
});

async function xmlEntry(archive: JSZip, path: string): Promise<Document> {
  const entry = archive.file(path);
  if (!entry) throw new Error(`Missing ${path}`);
  return parseXml(await entry.async('string'), path);
}

function floatAttributes(element: Element): Record<string, string> {
  return Object.fromEntries(
    Array.from(element.attributes)
      .filter(
        (item) =>
          xmlAttributeNamespace(element, item) === WORD_NAMESPACE ||
          item.namespaceURI === WORD_NAMESPACE,
      )
      .map((item) => [xmlAttributeLocalName(item), item.value]),
  );
}
