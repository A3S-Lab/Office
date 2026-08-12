import { describe, expect, test } from '@rstest/core';
import JSZip from 'jszip';
import { createArtifactBlob, importOfficeFile } from '../src/core';
import { createDocxBlob } from '../src/internal/features/work/work-docx-export';
import { patchDocxNoteImageRelationships } from '../src/internal/features/work/work-docx-note-image-relationships';
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
import type { WorkDocumentContent } from '../src/internal/features/work/work-types';

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
const RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const STRICT_RELATIONSHIP_NAMESPACE =
  'http://purl.oclc.org/ooxml/officeDocument/relationships';
const MARKUP_COMPATIBILITY_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';
const WORDPROCESSING_DRAWING_2010_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing';
const XMLNS_NAMESPACE = 'http://www.w3.org/2000/xmlns/';
const VENDOR_NAMESPACE = 'urn:a3s:test:note-drawing';
const pixelPng =
  'data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwC' +
  'AAAAC0lEQVR42mP8/x8AAusB9Y9Z9WQAAAAASUVORK5CYII=';

describe('DOCX note drawing preservation', () => {
  test('exports and publicly round-trips native footnote and endnote images', async () => {
    for (const [kind, identity] of [
      [
        'footnote',
        { anchorId: '4A2B3C4D', docPropertiesId: 45, editId: '3A0B0C0D' },
      ],
      [
        'endnote',
        { anchorId: '5A2B3C4D', docPropertiesId: 46, editId: '4A0B0C0D' },
      ],
    ] as const) {
      const content = noteContent(kind, imageHtml(`${kind} image`, identity));
      const seed = await createDocxBlob(content);
      const archive = await JSZip.loadAsync(await seed.arrayBuffer());
      const path = `word/${kind}s.xml`;
      const note = noteById(
        await xmlEntry(archive, path),
        kind,
        kind === 'footnote' ? '27' : '41',
      );
      expect(
        drawingByIdentity(
          note,
          identity.anchorId,
          String(identity.docPropertiesId),
        ),
      ).toBeDefined();
      await expectImageRelationship(archive, path, note);

      const imported = await importOfficeFile(
        new File([seed], `${kind}-drawing.docx`, { type: seed.type }),
      );
      if (imported.content.type !== 'document') {
        throw new Error('Expected an imported document artifact.');
      }
      const html = new DOMParser().parseFromString(
        imported.content.html,
        'text/html',
      );
      const importedImage = html.body.querySelector<HTMLImageElement>(
        `aside[data-note-kind="${kind}"] img`,
      );
      expect(importedImage?.dataset.officeImageDocPropertiesId).toBe(
        String(identity.docPropertiesId),
      );
      expect(importedImage?.dataset.officeImageAnchorId).toBe(
        identity.anchorId,
      );
      expect(importedImage?.dataset.officeImageEditId).toBe(identity.editId);
      expect(importedImage?.dataset.officeImageLayout).toBe('square');
      expect(importedImage?.dataset.officeImageAlignment).toBe('right');
      expect(importedImage?.dataset.officeImageWrapDistance).toBe('3');
      expect(importedImage?.dataset.officeImageCropTop).toBe('12.5');

      const output = await JSZip.loadAsync(
        await (await createArtifactBlob(imported)).arrayBuffer(),
      );
      const outputNote = noteById(
        await xmlEntry(output, path),
        kind,
        kind === 'footnote' ? '27' : '41',
      );
      expect(
        drawingByIdentity(
          outputNote,
          identity.anchorId,
          String(identity.docPropertiesId),
        ),
      ).toBeDefined();
    }
  });

  test('retains passive metadata on a stable strict UTF-16 note drawing', async () => {
    const identity = {
      anchorId: '6A2B3C4D',
      docPropertiesId: 47,
      editId: '5A0B0C0D',
    };
    const content = noteContent(
      'footnote',
      imageHtml('Strict note image', identity),
    );
    const source = await archiveFromContent(content);
    const footnotes = await xmlEntry(source, 'word/footnotes.xml');
    const note = noteById(footnotes, 'footnote', '27');
    const drawing = drawingByIdentity(
      note,
      identity.anchorId,
      String(identity.docPropertiesId),
    );
    decorateDrawing(drawing, 'strict-note');
    const extent = directChild(drawing, 'extent');
    if (!extent) throw new Error('Generated note extent is missing.');
    extent.setAttribute('cx', '1');
    source.file('word/footnotes.xml', strictUtf16(footnotes));

    const output = await exportWithSource(content, source);
    const outputNote = noteById(
      await xmlEntry(output, 'word/footnotes.xml'),
      'footnote',
      '27',
    );
    const outputDrawing = drawingByIdentity(
      outputNote,
      identity.anchorId,
      String(identity.docPropertiesId),
    );
    expect(vendorAttribute(outputDrawing, 'stable')).toBe(
      'strict-note-metadata',
    );
    expect(vendorValue(outputDrawing, 'drawingMeta')).toBe(
      'strict-note-extension',
    );
    expect(
      vendorAttribute(directChild(outputDrawing, 'docPr'), 'propertiesStable'),
    ).toBe('strict-note-properties');
    expect(vendorValue(outputDrawing, 'unsafeDrawing')).toBeNull();
    expect(descendants(outputDrawing, 'sizeRelH')).toHaveLength(0);
    expect(directChild(outputDrawing, 'extent')?.getAttribute('cx')).not.toBe(
      '1',
    );
    await expect(
      output.file('word/footnotes.xml')?.async('text'),
    ).resolves.toContain('encoding="UTF-8"');
  });

  test('preserves a strict UTF-16 source drawing through the public artifact boundary', async () => {
    const identity = {
      anchorId: '7A2B3C4D',
      docPropertiesId: 48,
      editId: '6A0B0C0D',
    };
    const content = noteContent(
      'endnote',
      imageHtml('Imported endnote image', identity),
    );
    const source = await archiveFromContent(content);
    const endnotes = await xmlEntry(source, 'word/endnotes.xml');
    const drawing = drawingByIdentity(
      noteById(endnotes, 'endnote', '41'),
      identity.anchorId,
      String(identity.docPropertiesId),
    );
    declareIgnorableVendor(endnotes.documentElement);
    appendVendor(drawing, 'drawingMeta', 'public-round-trip');
    source.file('word/endnotes.xml', strictUtf16(endnotes));
    await expectImageRelationship(
      source,
      'word/endnotes.xml',
      noteById(endnotes, 'endnote', '41'),
    );
    const buffer = await source.generateAsync({ type: 'arraybuffer' });

    const imported = await importOfficeFile(
      new File([buffer], 'source-backed-note-drawing.docx'),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    expect(imported.content.html).toContain(
      `data-office-image-anchor-id="${identity.anchorId}"`,
    );
    const output = await JSZip.loadAsync(
      await (await createArtifactBlob(imported)).arrayBuffer(),
    );
    const outputDrawing = drawingByIdentity(
      noteById(await xmlEntry(output, 'word/endnotes.xml'), 'endnote', '41'),
      identity.anchorId,
      String(identity.docPropertiesId),
    );
    expect(vendorValue(outputDrawing, 'drawingMeta')).toBe('public-round-trip');
  });

  test('fails closed for ambiguous, changed, spoofed, and relationship-bound drawing metadata', async () => {
    for (const boundary of [
      'ambiguous',
      'changed',
      'relationship',
      'spoofed',
    ] as const) {
      const identity = {
        anchorId: '1B2C3D4E',
        docPropertiesId: 91,
        editId: '2B3C4D5E',
      };
      const content = noteContent(
        'footnote',
        imageHtml(`${boundary} note image`, identity),
      );
      const source = await archiveFromContent(content);
      const footnotes = await xmlEntry(source, 'word/footnotes.xml');
      const note = noteById(footnotes, 'footnote', '27');
      const drawing = drawingByIdentity(
        note,
        identity.anchorId,
        String(identity.docPropertiesId),
      );
      declareIgnorableVendor(footnotes.documentElement);
      const extension = appendVendor(drawing, 'mustNotReattach', boundary);
      if (boundary === 'ambiguous') {
        const run = closestAncestor(drawing, 'r');
        run?.parentNode?.insertBefore(run.cloneNode(true), run.nextSibling);
      } else if (boundary === 'changed') {
        setDrawing2010Attribute(drawing, 'anchorId', '3B2C3D4E');
      } else if (boundary === 'relationship') {
        extension.setAttributeNS(
          RELATIONSHIP_NAMESPACE,
          'r:id',
          'rIdUnsafeDrawing',
        );
      } else {
        removeDrawing2010Attribute(drawing, 'anchorId');
        drawing.setAttributeNS(
          VENDOR_NAMESPACE,
          'vnd:anchorId',
          identity.anchorId,
        );
      }
      source.file('word/footnotes.xml', serializeXml(footnotes));

      const output = await exportWithSource(content, source);
      const outputNote = noteById(
        await xmlEntry(output, 'word/footnotes.xml'),
        'footnote',
        '27',
      );
      expect(descendants(outputNote, 'mustNotReattach'), boundary).toHaveLength(
        0,
      );
      expect(descendants(outputNote, 'drawing'), boundary).toHaveLength(1);
    }
  });

  test('rejects ambiguous, external, and missing note image relationships', async () => {
    const content = noteContent(
      'endnote',
      imageHtml('Relationship boundary', {
        anchorId: '4B2C3D4E',
        docPropertiesId: 92,
        editId: '5B3C4D5E',
      }),
    );
    const seed = await createDocxBlob(content);
    const seedBuffer = await seed.arrayBuffer();

    const duplicate = await JSZip.loadAsync(seedBuffer);
    const duplicateRelationships = await xmlEntry(
      duplicate,
      'word/_rels/endnotes.xml.rels',
    );
    const duplicateImageRelationship = imageRelationship(
      duplicateRelationships,
    );
    duplicateRelationships.documentElement.append(
      duplicateImageRelationship.cloneNode(true),
    );
    duplicate.file(
      'word/_rels/endnotes.xml.rels',
      serializeXml(duplicateRelationships),
    );
    await expect(
      patchDocxNoteImageRelationships(
        await duplicate.generateAsync({ type: 'arraybuffer' }),
      ),
    ).rejects.toThrow('duplicate IDs');

    const missing = await JSZip.loadAsync(seedBuffer);
    const missingRelationships = await xmlEntry(
      missing,
      'word/_rels/endnotes.xml.rels',
    );
    const target =
      imageRelationship(missingRelationships).getAttribute('Target');
    if (!target) throw new Error('Generated image relationship has no target.');
    missing.remove(`word/${target}`);
    await expect(
      patchDocxNoteImageRelationships(
        await missing.generateAsync({ type: 'arraybuffer' }),
      ),
    ).rejects.toThrow('relationship is invalid');

    const external = await JSZip.loadAsync(seedBuffer);
    const externalRelationships = await xmlEntry(
      external,
      'word/_rels/endnotes.xml.rels',
    );
    imageRelationship(externalRelationships).setAttribute(
      'TargetMode',
      'External',
    );
    external.file(
      'word/_rels/endnotes.xml.rels',
      serializeXml(externalRelationships),
    );
    await expect(
      patchDocxNoteImageRelationships(
        await external.generateAsync({ type: 'arraybuffer' }),
      ),
    ).rejects.toThrow('relationship is invalid');
  });
});

function noteContent(
  kind: 'endnote' | 'footnote',
  noteHtml: string,
): WorkDocumentContent {
  const nativeId = kind === 'footnote' ? '27' : '41';
  return {
    type: 'document',
    pageSize: 'a4',
    html: [
      '<section data-document-section="true">',
      `<p>Body<sup data-document-note-reference="true" data-note-kind="${kind}" data-note-id="docx-${kind}-${nativeId}" data-note-number="1">1</sup></p>`,
      `<aside data-document-note="true" data-note-kind="${kind}" data-note-id="docx-${kind}-${nativeId}" data-note-number="1"><p>Before ${noteHtml} After</p></aside>`,
      '</section>',
    ].join(''),
  };
}

function imageHtml(
  alternativeText: string,
  identity: {
    anchorId: string;
    docPropertiesId: number;
    editId: string;
  },
): string {
  return [
    `<img src="${pixelPng}" alt="${alternativeText}" width="40" height="30"`,
    ` data-office-image-object-id="${identity.anchorId}"`,
    ` data-office-image-doc-properties-id="${identity.docPropertiesId}"`,
    ` data-office-image-anchor-id="${identity.anchorId}"`,
    ` data-office-image-edit-id="${identity.editId}"`,
    ' data-office-image-layout="square"',
    ' data-office-image-alignment="right"',
    ' data-office-image-wrap-distance="3"',
    ' data-office-image-crop-top="12.5">',
  ].join('');
}

async function archiveFromContent(
  content: WorkDocumentContent,
): Promise<JSZip> {
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

async function expectImageRelationship(
  archive: JSZip,
  partPath: string,
  scope: Element,
): Promise<void> {
  const blip = descendants(scope, 'blip')[0];
  const relationshipId = relationshipAttribute(blip, 'embed');
  expect(relationshipId).toMatch(/^rId/u);
  const fileName = partPath.slice(partPath.lastIndexOf('/') + 1);
  const relationshipsPath = `word/_rels/${fileName}.rels`;
  const relationships = await xmlEntry(archive, relationshipsPath);
  const relationship = descendants(relationships, 'Relationship').find(
    (item) => attribute(item, 'Id') === relationshipId,
  );
  if (!relationship) {
    throw new Error(
      `Missing ${partPath} image relationship ${relationshipId}: ${new XMLSerializer().serializeToString(relationships)}`,
    );
  }
  expect(attribute(relationship, 'Type')).toMatch(/\/image$/u);
  expect(attribute(relationship, 'Target')).toMatch(/^media\//u);
}

function imageRelationship(document: Document): Element {
  const relationship = descendants(document, 'Relationship').find((item) =>
    item.getAttribute('Type')?.endsWith('/image'),
  );
  if (!relationship)
    throw new Error('Generated image relationship is missing.');
  return relationship;
}

function noteById(
  document: Document,
  kind: 'endnote' | 'footnote',
  id: string,
): Element {
  const note = descendants(document, kind).find(
    (item) => wordAttribute(item, 'id') === id,
  );
  if (!note) throw new Error(`Missing ${kind} ${id}.`);
  return note;
}

function drawingByIdentity(
  scope: ParentNode,
  anchorId: string,
  docPropertiesId: string,
): Element {
  const drawing = descendants(scope, 'anchor')
    .concat(descendants(scope, 'inline'))
    .find(
      (item) =>
        drawing2010Attribute(item, 'anchorId') === anchorId &&
        directChild(item, 'docPr')?.getAttribute('id') === docPropertiesId,
    );
  if (!drawing) {
    throw new Error(`Missing drawing identity ${anchorId}/${docPropertiesId}.`);
  }
  return drawing;
}

function decorateDrawing(drawing: Element, label: string): void {
  declareIgnorableVendor(drawing.ownerDocument.documentElement);
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
  unsafe.setAttributeNS(RELATIONSHIP_NAMESPACE, 'r:id', 'rIdUnsafeDrawing');
  const semantic = drawing.ownerDocument.createElementNS(
    WORDPROCESSING_DRAWING_2010_NAMESPACE,
    'wp14:sizeRelH',
  );
  drawing.append(semantic);
}

function declareIgnorableVendor(root: Element): void {
  root.setAttributeNS(XMLNS_NAMESPACE, 'xmlns:vnd', VENDOR_NAMESPACE);
  if (!root.lookupNamespaceURI?.('mc')) {
    root.setAttributeNS(
      XMLNS_NAMESPACE,
      'xmlns:mc',
      MARKUP_COMPATIBILITY_NAMESPACE,
    );
  }
  const current = (attribute(root, 'mc:Ignorable') ?? '').trim();
  const ignorable = Array.from(root.attributes).find(
    (item) =>
      xmlAttributeLocalName(item) === 'Ignorable' &&
      xmlAttributeNamespace(root, item) === MARKUP_COMPATIBILITY_NAMESPACE,
  );
  const value = Array.from(new Set([current, 'vnd'].filter(Boolean))).join(' ');
  if (ignorable) ignorable.value = value;
  else {
    root.setAttributeNS(MARKUP_COMPATIBILITY_NAMESPACE, 'mc:Ignorable', value);
  }
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

function setDrawing2010Attribute(
  element: Element,
  name: string,
  value: string,
): void {
  removeDrawing2010Attribute(element, name);
  element.setAttributeNS(
    WORDPROCESSING_DRAWING_2010_NAMESPACE,
    `wp14:${name}`,
    value,
  );
}

function removeDrawing2010Attribute(element: Element, name: string): void {
  for (const item of Array.from(element.attributes)) {
    if (
      xmlAttributeLocalName(item) === name &&
      xmlAttributeNamespace(element, item) ===
        WORDPROCESSING_DRAWING_2010_NAMESPACE
    ) {
      element.removeAttributeNode(item);
    }
  }
}

function closestAncestor(element: Element, localName: string): Element | null {
  let current: Element | null = element;
  while (current) {
    if (current.localName === localName) return current;
    current = current.parentElement;
  }
  return null;
}

function wordAttribute(element: Element, name: string): string | null {
  return (
    Array.from(element.attributes).find(
      (item) =>
        xmlAttributeLocalName(item) === name &&
        (xmlAttributeNamespace(element, item) === WORD_NAMESPACE ||
          xmlAttributeNamespace(element, item) === STRICT_WORD_NAMESPACE),
    )?.value ?? null
  );
}

function drawing2010Attribute(element: Element, name: string): string | null {
  return (
    Array.from(element.attributes).find(
      (item) =>
        xmlAttributeLocalName(item) === name &&
        xmlAttributeNamespace(element, item) ===
          WORDPROCESSING_DRAWING_2010_NAMESPACE,
    )?.value ?? null
  );
}

function relationshipAttribute(
  element: Element | undefined,
  name: string,
): string | null {
  if (!element) return null;
  return (
    Array.from(element.attributes).find(
      (item) =>
        xmlAttributeLocalName(item) === name &&
        (xmlAttributeNamespace(element, item) === RELATIONSHIP_NAMESPACE ||
          xmlAttributeNamespace(element, item) ===
            STRICT_RELATIONSHIP_NAMESPACE),
    )?.value ?? null
  );
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
      .replaceAll(RELATIONSHIP_NAMESPACE, STRICT_RELATIONSHIP_NAMESPACE)
      .replace(
        /^\s*<\?xml[^?]*\?>/iu,
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
