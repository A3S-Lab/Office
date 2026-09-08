import JSZip from 'jszip';
import { importDocxScopedPropertyRevisionElement } from './work-document-table-property-revision';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import { descendants, directChild, parseXml } from './work-ooxml-package';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';
import type { WorkDocumentSectionLayout } from './work-types';

const MAX_SECTION_PROPERTY_REVISION_PATCHES = 512;

export class DocxSectionPropertyRevisionPatchCollector {
  readonly patches: Array<string | null> = [];

  record(layout: WorkDocumentSectionLayout): void {
    if (layout.formattingChange?.kind === 'section-formatting') {
      this.patches.push(null);
      return;
    }
    if (this.patches.length >= MAX_SECTION_PROPERTY_REVISION_PATCHES) {
      throw new Error('Document exceeds the section property-revision limit.');
    }
    this.patches.push(layout.propertyRevisionOmml?.trim() || null);
  }
}

export async function patchDocxSectionPropertyRevisions(
  buffer: ArrayBuffer,
  patches: readonly (string | null)[],
): Promise<ArrayBuffer> {
  if (!patches.some(Boolean)) return buffer;
  if (patches.length > MAX_SECTION_PROPERTY_REVISION_PATCHES) {
    throw new Error('Document exceeds the section property-revision limit.');
  }
  const archive = await JSZip.loadAsync(buffer);
  const entry = archive.file('word/document.xml');
  if (!entry) return buffer;
  const document = parseXml(
    decodeXmlBytes(
      await entry.async('uint8array'),
      'generated DOCX word/document.xml',
    ),
    'generated DOCX word/document.xml',
  );
  let patchIndex = 0;
  let changed = false;
  for (const section of descendants(document, 'sectPr')) {
    if (!DOCX_WORDPROCESSING_NAMESPACES.has(section.namespaceURI ?? '')) {
      continue;
    }
    // Prior snapshots live inside sectPrChange; skip those nested sectPr nodes.
    if (section.parentElement?.localName === 'sectPrChange') continue;
    const omml = patches[patchIndex] ?? null;
    patchIndex += 1;
    if (!omml) continue;
    const revision = importDocxScopedPropertyRevisionElement(
      document,
      omml,
      'sectPrChange',
    );
    if (!revision) continue;
    const existing = directChild(section, 'sectPrChange');
    if (existing) existing.replaceWith(revision);
    else section.append(revision);
    changed = true;
  }
  if (changed) {
    archive.file('word/document.xml', serializeUtf8Xml(document));
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}
