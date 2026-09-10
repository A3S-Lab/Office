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

describe('DOCX section property-revision preservation', () => {
  test('preserves relationship-free w:sectPrChange as opaque section metadata on untouched round-trip', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    const seed = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await seed.arrayBuffer());
    const document = await xmlEntry(archive, 'word/document.xml');
    const section = descendants(document, 'sectPr').find(
      (element) => element.parentElement?.localName !== 'sectPrChange',
    );
    if (!section) throw new Error('Expected body sectPr.');
    const change = document.createElementNS(WORD_NAMESPACE, 'w:sectPrChange');
    change.setAttributeNS(WORD_NAMESPACE, 'w:id', '21');
    change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
    const prior = document.createElementNS(WORD_NAMESPACE, 'w:sectPr');
    // printerSettings is CT_Rel (required r:id); keep empty relationship-free fixture opaque.
    const printerSettings = document.createElementNS(
      WORD_NAMESPACE,
      'w:printerSettings',
    );
    prior.append(printerSettings);
    change.append(prior);
    section.append(change);
    archive.file(
      'word/document.xml',
      new XMLSerializer().serializeToString(document),
    );
    const imported = await importOfficeFile(
      new File(
        [await archive.generateAsync({ type: 'arraybuffer' })],
        'section-property-revision.docx',
      ),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    expect(
      imported.content.sections?.[0]?.layout.propertyRevisionOmml ||
        imported.content.html.includes('data-section-property-revision-omml'),
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
    const exportedSection = descendants(exported, 'sectPr').find(
      (element) => element.parentElement?.localName !== 'sectPrChange',
    );
    const exportedChange = directChild(exportedSection!, 'sectPrChange');
    expect(exportedChange).toBeTruthy();
    expect(changeAttributes(exportedChange!)).toMatchObject({
      id: '21',
      author: 'Reviewer',
    });
    const priorPrinterSettings = directChild(
      directChild(exportedChange!, 'sectPr'),
      'printerSettings',
    );
    expect(priorPrinterSettings).toBeTruthy();
  });

  test('drops relationship-bound w:sectPrChange instead of inventing opaque metadata', async () => {
    const artifact = createArtifact('blank-document');
    if (artifact.content.type !== 'document') {
      throw new Error('Expected a document artifact.');
    }
    const seed = await createArtifactBlob(artifact);
    const archive = await JSZip.loadAsync(await seed.arrayBuffer());
    const document = await xmlEntry(archive, 'word/document.xml');
    const section = descendants(document, 'sectPr').find(
      (element) => element.parentElement?.localName !== 'sectPrChange',
    );
    if (!section) throw new Error('Expected body sectPr.');
    const change = document.createElementNS(WORD_NAMESPACE, 'w:sectPrChange');
    change.setAttributeNS(WORD_NAMESPACE, 'w:id', '22');
    change.setAttributeNS(WORD_NAMESPACE, 'w:author', 'Reviewer');
    const prior = document.createElementNS(WORD_NAMESPACE, 'w:sectPr');
    const headerRef = document.createElementNS(
      WORD_NAMESPACE,
      'w:headerReference',
    );
    headerRef.setAttributeNS(RELATIONSHIP_NAMESPACE, 'r:id', 'rIdHeader');
    prior.append(headerRef);
    change.append(prior);
    section.append(change);
    archive.file(
      'word/document.xml',
      new XMLSerializer().serializeToString(document),
    );
    const imported = await importOfficeFile(
      new File(
        [await archive.generateAsync({ type: 'arraybuffer' })],
        'section-property-revision-bound.docx',
      ),
    );
    if (imported.content.type !== 'document') {
      throw new Error('Expected an imported document artifact.');
    }
    expect(
      imported.content.sections?.[0]?.layout.propertyRevisionOmml,
    ).toBeFalsy();
    expect(
      imported.content.html.includes('data-section-property-revision-omml'),
    ).toBe(false);
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
