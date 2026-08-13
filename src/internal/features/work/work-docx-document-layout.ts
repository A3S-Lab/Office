import JSZip from 'jszip';
import {
  DOCUMENT_PAGE_BORDER_EDGES,
  normalizeDocumentPageBorders,
} from './work-document-page-borders';
import type { WorkDocumentSection } from './work-document-section';
import { setDocxBorderAttributes } from './work-docx-paragraph-borders-export';
import {
  descendants,
  directChildren,
  parseXml,
  xmlNamespacePrefix,
} from './work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const SECTION_PROPERTY_ORDER = [
  'headerReference',
  'footerReference',
  'footnotePr',
  'endnotePr',
  'type',
  'pgSz',
  'pgMar',
  'paperSrc',
  'pgBorders',
  'lnNumType',
  'pgNumType',
  'cols',
  'formProt',
  'vAlign',
  'noEndnote',
  'titlePg',
  'textDirection',
  'bidi',
  'rtlGutter',
  'docGrid',
  'printerSettings',
  'sectPrChange',
] as const;

export async function patchDocxDocumentLayout(
  buffer: ArrayBuffer,
  sections: readonly WorkDocumentSection[],
): Promise<ArrayBuffer> {
  const archive = await JSZip.loadAsync(buffer);
  const entry = archive.file('word/document.xml');
  if (!entry) return buffer;
  const document = parseXml(await entry.async('string'), 'word/document.xml');
  patchSectionPageBorders(document, sections);
  patchSectionDocumentGrids(document, sections);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  return archive.generateAsync({ type: 'arraybuffer' });
}

function patchSectionPageBorders(
  document: Document,
  sections: readonly WorkDocumentSection[],
): void {
  const sectionProperties = effectiveSectionProperties(document);
  if (
    sections.some((section) => section.layout.pageBorders) &&
    sectionProperties.length !== sections.length
  ) {
    throw new Error(
      `Generated DOCX has ${sectionProperties.length} section properties for ${sections.length} page-border section(s).`,
    );
  }
  const prefix =
    xmlNamespacePrefix(document.documentElement, WORD_NAMESPACE) ?? 'w';
  for (const [index, properties] of sectionProperties.entries()) {
    for (const element of directChildren(properties, 'pgBorders').filter(
      (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
    )) {
      element.remove();
    }
    const value = normalizeDocumentPageBorders(
      sections[index]?.layout.pageBorders,
    );
    if (!value) continue;
    const container = document.createElementNS(
      WORD_NAMESPACE,
      `${prefix}:pgBorders`,
    );
    if (value.zOrder)
      setWordAttribute(document, container, 'zOrder', value.zOrder);
    if (value.display)
      setWordAttribute(document, container, 'display', value.display);
    if (value.offsetFrom)
      setWordAttribute(document, container, 'offsetFrom', value.offsetFrom);
    for (const edge of DOCUMENT_PAGE_BORDER_EDGES) {
      const border = value.edges[edge];
      if (!border) continue;
      const element = document.createElementNS(
        WORD_NAMESPACE,
        `${prefix}:${edge}`,
      );
      setDocxBorderAttributes(document, element, border);
      container.append(element);
    }
    insertSectionProperty(properties, container);
  }
}

function patchSectionDocumentGrids(
  document: Document,
  sections: readonly WorkDocumentSection[],
): void {
  const sectionProperties = effectiveSectionProperties(document);
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
    insertSectionProperty(properties, grid);
  }
}

function effectiveSectionProperties(document: Document): Element[] {
  return descendants(document, 'sectPr').filter(
    (element) => !hasAncestor(element, 'sectPrChange'),
  );
}

function insertSectionProperty(parent: Element, element: Element): void {
  const targetIndex = SECTION_PROPERTY_ORDER.indexOf(
    element.localName as (typeof SECTION_PROPERTY_ORDER)[number],
  );
  const next = directChildren(parent).find((candidate) => {
    const index = SECTION_PROPERTY_ORDER.indexOf(
      candidate.localName as (typeof SECTION_PROPERTY_ORDER)[number],
    );
    return index >= 0 && index > targetIndex;
  });
  parent.insertBefore(element, next ?? null);
}

function setWordAttribute(
  document: Document,
  element: Element,
  name: string,
  value: string,
): void {
  const prefix =
    xmlNamespacePrefix(document.documentElement, WORD_NAMESPACE) ?? 'w';
  element.setAttributeNS(WORD_NAMESPACE, `${prefix}:${name}`, value);
}

function hasAncestor(element: Element, localName: string): boolean {
  let ancestor = element.parentElement;
  while (ancestor) {
    if (ancestor.localName === localName) return true;
    ancestor = ancestor.parentElement;
  }
  return false;
}
