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

describe('DOCX table property-revision preservation', () => {
  test('preserves relationship-free w:tblPrChange as opaque table metadata on untouched round-trip', async () => {
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
    const change = document.createElementNS(WORD_NAMESPACE, 'w:tblPrChange');
    change.setAttributeNS(WORD_NAMESPACE, 'w:id', '7');
    change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
    change.setAttributeNS(WORD_NAMESPACE, 'w:date', '2026-09-08T00:00:00Z');
    const prior = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
    const look = document.createElementNS(WORD_NAMESPACE, 'w:tblLook');
    look.setAttributeNS(WORD_NAMESPACE, 'w:val', '04A0');
    look.setAttributeNS(WORD_NAMESPACE, 'w:firstRow', '1');
    look.setAttributeNS(WORD_NAMESPACE, 'w:lastRow', '0');
    look.setAttributeNS(WORD_NAMESPACE, 'w:firstColumn', '1');
    look.setAttributeNS(WORD_NAMESPACE, 'w:lastColumn', '0');
    look.setAttributeNS(WORD_NAMESPACE, 'w:noHBand', '0');
    look.setAttributeNS(WORD_NAMESPACE, 'w:noVBand', '1');
    prior.append(look);
    change.append(prior);
    properties.append(change);
    archive.file(
      'word/document.xml',
      new XMLSerializer().serializeToString(document),
    );
    const source = await archive.generateAsync({ type: 'arraybuffer' });
    const imported = await importOfficeFile(
      new File([source], 'table-property-revision.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    const importedTable = html.body.querySelector('table');
    expect(importedTable?.dataset.officeTablePropertyRevisionOmml).toBeTruthy();

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
        edited.body.querySelector('table')?.dataset
          .officeTablePropertyRevisionOmml,
      ).toBeTruthy();
      imported.content.html = editor.getHTML();
    } finally {
      editor.destroy();
    }

    const output = await JSZip.loadAsync(
      await (await createArtifactBlob(imported)).arrayBuffer(),
    );
    const exported = await xmlEntry(output, 'word/document.xml');
    const exportedChange = directChild(
      directChild(descendants(exported, 'tbl')[0], 'tblPr'),
      'tblPrChange',
    );
    expect(exportedChange).toBeTruthy();
    expect(changeAttributes(exportedChange!)).toMatchObject({
      id: '7',
      author: 'Reviewer',
    });
    const priorLook = directChild(
      directChild(exportedChange!, 'tblPr'),
      'tblLook',
    );
    expect(
      priorLook?.getAttributeNS(WORD_NAMESPACE, 'val') ??
        priorLook?.getAttribute('w:val') ??
        priorLook?.getAttribute('val'),
    ).toBe('04A0');
  });

  test('drops relationship-bound w:tblPrChange instead of inventing opaque metadata', async () => {
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
    const change = document.createElementNS(WORD_NAMESPACE, 'w:tblPrChange');
    change.setAttributeNS(WORD_NAMESPACE, 'w:id', '3');
    change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
    const prior = document.createElementNS(WORD_NAMESPACE, 'w:tblPr');
    const linked = document.createElementNS(WORD_NAMESPACE, 'w:tblStyle');
    linked.setAttributeNS(RELATIONSHIP_NAMESPACE, 'r:id', 'rId9');
    prior.append(linked);
    change.append(prior);
    properties.append(change);
    archive.file(
      'word/document.xml',
      new XMLSerializer().serializeToString(document),
    );
    const source = await archive.generateAsync({ type: 'arraybuffer' });
    const imported = await importOfficeFile(
      new File([source], 'table-property-revision-bound.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    expect(
      html.body.querySelector('table')?.dataset.officeTablePropertyRevisionOmml,
    ).toBeFalsy();
  });

  test('preserves relationship-free w:trPrChange as opaque row metadata on untouched round-trip', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html =
      '<table><tbody><tr><td><p>Cell</p></td></tr></tbody></table>';
    const seed = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await seed.arrayBuffer());
    const document = await xmlEntry(archive, 'word/document.xml');
    const row = descendants(document, 'tr')[0];
    const properties =
      directChild(row, 'trPr') ??
      (() => {
        const created = document.createElementNS(WORD_NAMESPACE, 'w:trPr');
        row.insertBefore(created, row.firstChild);
        return created;
      })();
    const change = document.createElementNS(WORD_NAMESPACE, 'w:trPrChange');
    change.setAttributeNS(WORD_NAMESPACE, 'w:id', '11');
    change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
    const prior = document.createElementNS(WORD_NAMESPACE, 'w:trPr');
    const cantSplit = document.createElementNS(WORD_NAMESPACE, 'w:cantSplit');
    const header = document.createElementNS(WORD_NAMESPACE, 'w:tblHeader');
    // Extra property keeps this on the opaque path; cantSplit/tblHeader/trHeight/
    // hidden/jc alone are reviewable as row-formatting.
    const height = document.createElementNS(WORD_NAMESPACE, 'w:trHeight');
    height.setAttributeNS(WORD_NAMESPACE, 'w:val', '240');
    const gridBefore = document.createElementNS(WORD_NAMESPACE, 'w:gridBefore');
    gridBefore.setAttributeNS(WORD_NAMESPACE, 'w:val', '1');
    prior.append(cantSplit, header, height, gridBefore);
    change.append(prior);
    properties.append(change);
    archive.file(
      'word/document.xml',
      new XMLSerializer().serializeToString(document),
    );
    const imported = await importOfficeFile(
      new File(
        [await archive.generateAsync({ type: 'arraybuffer' })],
        'row-property-revision.docx',
      ),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    expect(
      html.body.querySelector('tr')?.dataset.officeRowPropertyRevisionOmml,
    ).toBeTruthy();

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      imported.content.html = editor.getHTML();
    } finally {
      editor.destroy();
    }

    const exported = await xmlEntry(
      await JSZip.loadAsync(
        await (await createArtifactBlob(imported)).arrayBuffer(),
      ),
      'word/document.xml',
    );
    const exportedChange = directChild(
      directChild(descendants(exported, 'tr')[0], 'trPr'),
      'trPrChange',
    );
    expect(exportedChange).toBeTruthy();
    expect(changeAttributes(exportedChange!)).toMatchObject({
      id: '11',
      author: 'Reviewer',
    });
    expect(
      directChild(directChild(exportedChange!, 'trPr'), 'cantSplit'),
    ).toBeTruthy();
    expect(
      directChild(directChild(exportedChange!, 'trPr'), 'tblHeader'),
    ).toBeTruthy();
    expect(
      directChild(directChild(exportedChange!, 'trPr'), 'trHeight'),
    ).toBeTruthy();
    expect(
      directChild(directChild(exportedChange!, 'trPr'), 'gridBefore'),
    ).toBeTruthy();
  });

  test('preserves relationship-free w:tcPrChange as opaque cell metadata on untouched round-trip', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    artifact.content.html =
      '<table><tbody><tr><td><p>Cell</p></td></tr></tbody></table>';
    const seed = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await seed.arrayBuffer());
    const document = await xmlEntry(archive, 'word/document.xml');
    const cell = descendants(document, 'tc')[0];
    const properties =
      directChild(cell, 'tcPr') ??
      (() => {
        const created = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
        cell.insertBefore(created, cell.firstChild);
        return created;
      })();
    const change = document.createElementNS(WORD_NAMESPACE, 'w:tcPrChange');
    change.setAttributeNS(WORD_NAMESPACE, 'w:id', '13');
    change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
    const prior = document.createElementNS(WORD_NAMESPACE, 'w:tcPr');
    const vAlign = document.createElementNS(WORD_NAMESPACE, 'w:vAlign');
    vAlign.setAttributeNS(WORD_NAMESPACE, 'w:val', 'center');
    prior.append(vAlign);
    // Multi-property prior outside the reviewable
    // vAlign/solid-shd/tcMar/tcW/noWrap/textDirection/tcFitText subset
    // (hideMark stays opaque).
    const shading = document.createElementNS(WORD_NAMESPACE, 'w:shd');
    shading.setAttributeNS(WORD_NAMESPACE, 'w:val', 'clear');
    shading.setAttributeNS(WORD_NAMESPACE, 'w:fill', 'FFFF00');
    prior.append(shading);
    const hideMark = document.createElementNS(WORD_NAMESPACE, 'w:hideMark');
    prior.append(hideMark);
    change.append(prior);
    properties.append(change);
    archive.file(
      'word/document.xml',
      new XMLSerializer().serializeToString(document),
    );
    const imported = await importOfficeFile(
      new File(
        [await archive.generateAsync({ type: 'arraybuffer' })],
        'cell-property-revision.docx',
      ),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    const html = new DOMParser().parseFromString(
      imported.content.html,
      'text/html',
    );
    expect(
      html.body.querySelector('td')?.dataset.officeCellPropertyRevisionOmml,
    ).toBeTruthy();

    const editor = new Editor({
      extensions: createWorkDocumentExtensions(),
      content: imported.content.html,
    });
    try {
      imported.content.html = editor.getHTML();
    } finally {
      editor.destroy();
    }

    const exported = await xmlEntry(
      await JSZip.loadAsync(
        await (await createArtifactBlob(imported)).arrayBuffer(),
      ),
      'word/document.xml',
    );
    const exportedChange = directChild(
      directChild(descendants(exported, 'tc')[0], 'tcPr'),
      'tcPrChange',
    );
    expect(exportedChange).toBeTruthy();
    expect(changeAttributes(exportedChange!)).toMatchObject({
      id: '13',
      author: 'Reviewer',
    });
    const priorProperties = directChild(exportedChange!, 'tcPr');
    const priorAlign = directChild(priorProperties, 'vAlign');
    expect(
      priorAlign?.getAttributeNS(WORD_NAMESPACE, 'val') ??
        priorAlign?.getAttribute('w:val') ??
        priorAlign?.getAttribute('val'),
    ).toBe('center');
    expect(directChild(priorProperties, 'shd')).toBeTruthy();
    expect(directChild(priorProperties, 'hideMark')).toBeTruthy();
  });
});

async function xmlEntry(archive: JSZip, path: string): Promise<Document> {
  const entry = archive.file(path);
  if (!entry) throw new Error(`Missing ${path}`);
  return parseXml(await entry.async('string'), path);
}

function changeAttributes(element: Element): Record<string, string> {
  const result: Record<string, string> = {};
  for (const attribute of Array.from(element.attributes)) {
    const local = xmlAttributeLocalName(attribute);
    const namespace =
      attribute.namespaceURI || xmlAttributeNamespace(element, attribute);
    if (namespace && namespace !== WORD_NAMESPACE) continue;
    result[local] = attribute.value;
  }
  return result;
}
