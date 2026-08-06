import JSZip from 'jszip';
import type { WorkDocumentSection } from './work-document-section';
import { descendants, directChildren, parseXml } from './work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

export async function patchDocxDocumentLayout(
  buffer: ArrayBuffer,
  sections: readonly WorkDocumentSection[],
): Promise<ArrayBuffer> {
  const archive = await JSZip.loadAsync(buffer);
  const entry = archive.file('word/document.xml');
  if (!entry) return buffer;
  const document = parseXml(await entry.async('string'), 'word/document.xml');
  patchSectionDocumentGrids(document, sections);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

function patchSectionDocumentGrids(
  document: Document,
  sections: readonly WorkDocumentSection[],
): void {
  const sectionProperties = descendants(document, 'sectPr').filter(
    (element) => !hasAncestor(element, 'sectPrChange'),
  );
  for (const [index, properties] of sectionProperties.entries()) {
    for (const grid of directChildren(properties, 'docGrid')) grid.remove();
    const value = sections[index]?.layout.documentGrid;
    if (!value) continue;
    const grid = document.createElementNS(WORD_NAMESPACE, 'w:docGrid');
    grid.setAttributeNS(WORD_NAMESPACE, 'w:type', value.type);
    grid.setAttributeNS(
      WORD_NAMESPACE,
      'w:linePitch',
      String(Math.max(1, Math.round(value.linePitch * 20))),
    );
    properties.append(grid);
  }
}

function hasAncestor(element: Element, localName: string): boolean {
  let ancestor = element.parentElement;
  while (ancestor) {
    if (ancestor.localName === localName) return true;
    ancestor = ancestor.parentElement;
  }
  return false;
}
