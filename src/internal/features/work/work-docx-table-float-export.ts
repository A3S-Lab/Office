import JSZip from 'jszip';
import {
  documentTableFloatOmmlFromElement,
  importDocxTableFloatElement,
} from './work-document-table-float';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import { descendants, directChild, parseXml } from './work-ooxml-package';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';

const MAX_TABLE_FLOAT_PATCHES = 4_096;

export class DocxTableFloatPatchCollector {
  readonly patches: Array<string | null> = [];

  record(element: HTMLTableElement): void {
    if (this.patches.length >= MAX_TABLE_FLOAT_PATCHES) {
      throw new Error('Document exceeds the floating-table limit.');
    }
    this.patches.push(documentTableFloatOmmlFromElement(element));
  }
}

export async function patchDocxTableFloats(
  buffer: ArrayBuffer,
  patches: readonly (string | null)[],
): Promise<ArrayBuffer> {
  if (!patches.some(Boolean)) return buffer;
  if (patches.length > MAX_TABLE_FLOAT_PATCHES) {
    throw new Error('Document exceeds the floating-table limit.');
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
  for (const table of descendants(document, 'tbl')) {
    if (!DOCX_WORDPROCESSING_NAMESPACES.has(table.namespaceURI ?? '')) {
      continue;
    }
    const omml = patches[patchIndex] ?? null;
    patchIndex += 1;
    if (!omml) continue;
    const float = importDocxTableFloatElement(document, omml);
    if (!float) continue;
    let properties = directChild(table, 'tblPr');
    if (!properties) {
      const namespace =
        table.namespaceURI ?? [...DOCX_WORDPROCESSING_NAMESPACES][0];
      const prefix = table.prefix || 'w';
      properties = document.createElementNS(namespace, `${prefix}:tblPr`);
      table.insertBefore(properties, table.firstChild);
    }
    const existing = directChild(properties, 'tblpPr');
    if (existing) existing.replaceWith(float);
    else properties.append(float);
    changed = true;
  }
  if (changed) {
    archive.file('word/document.xml', serializeUtf8Xml(document));
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}
