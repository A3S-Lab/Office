import JSZip from 'jszip';
import {
  DOCUMENT_PAGE_BORDER_EDGES,
  normalizeDocumentPageBorders,
} from './work-document-page-borders';
import {
  DOCUMENT_PAGE_MARGIN_KEYS,
  type WorkDocumentPageMargins,
  normalizeDocumentPageMargins,
} from './work-document-page-margins';
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
const SETTINGS_ORDER = [
  'mirrorMargins',
  'alignBordersAndEdges',
  'bordersDoNotSurroundHeader',
  'bordersDoNotSurroundFooter',
  'gutterAtTop',
  'hideSpellingErrors',
  'hideGrammaticalErrors',
  'activeWritingStyle',
  'proofState',
  'formsDesign',
  'attachedTemplate',
  'linkStyles',
  'trackRevisions',
  'evenAndOddHeaders',
  'updateFields',
  'defaultTabStop',
  'compat',
] as const;

export async function patchDocxDocumentLayout(
  buffer: ArrayBuffer,
  sections: readonly WorkDocumentSection[],
): Promise<ArrayBuffer> {
  const archive = await JSZip.loadAsync(buffer);
  const entry = archive.file('word/document.xml');
  if (!entry) return buffer;
  const document = parseXml(await entry.async('string'), 'word/document.xml');
  patchSectionPageMargins(document, sections);
  patchSectionPageBorders(document, sections);
  patchSectionDocumentGrids(document, sections);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  await patchDocumentPageMarginSettings(archive, sections);
  return archive.generateAsync({ type: 'arraybuffer' });
}

function patchSectionPageMargins(
  document: Document,
  sections: readonly WorkDocumentSection[],
): void {
  const sectionProperties = effectiveSectionProperties(document);
  if (
    sections.some((section) => section.layout.pageMargins) &&
    sectionProperties.length !== sections.length
  ) {
    throw new Error(
      `Generated DOCX has ${sectionProperties.length} section properties for ${sections.length} exact page-margin section(s).`,
    );
  }
  const prefix =
    xmlNamespacePrefix(document.documentElement, WORD_NAMESPACE) ?? 'w';
  for (const [index, properties] of sectionProperties.entries()) {
    const pageMargins = normalizeDocumentPageMargins(
      sections[index]?.layout.pageMargins,
    );
    if (pageMargins) {
      for (const element of directChildren(properties, 'pgMar').filter(
        (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
      )) {
        element.remove();
      }
      const element = document.createElementNS(
        WORD_NAMESPACE,
        `${prefix}:pgMar`,
      );
      for (const key of DOCUMENT_PAGE_MARGIN_KEYS) {
        setWordAttribute(document, element, key, String(pageMargins[key]));
      }
      insertSectionProperty(properties, element);
    }
    for (const element of directChildren(properties, 'rtlGutter').filter(
      (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
    )) {
      element.remove();
    }
    if (pageMargins?.gutterOnRight !== undefined) {
      const element = document.createElementNS(
        WORD_NAMESPACE,
        `${prefix}:rtlGutter`,
      );
      setWordAttribute(
        document,
        element,
        'val',
        pageMargins.gutterOnRight ? '1' : '0',
      );
      insertSectionProperty(properties, element);
    }
  }
}

async function patchDocumentPageMarginSettings(
  archive: JSZip,
  sections: readonly WorkDocumentSection[],
): Promise<void> {
  const pageMargins = sections.flatMap((section) => {
    const value = normalizeDocumentPageMargins(section.layout.pageMargins);
    return value ? [value] : [];
  });
  if (!pageMargins.length) return;
  assertConsistentPageMarginSettings(pageMargins);
  const entry = archive.file('word/settings.xml');
  if (!entry) {
    throw new Error(
      'Generated DOCX lacks word/settings.xml for exact page-margin settings.',
    );
  }
  const document = parseXml(await entry.async('string'), 'word/settings.xml');
  const root = document.documentElement;
  if (root.localName !== 'settings' || root.namespaceURI !== WORD_NAMESPACE) {
    throw new Error(
      'Generated DOCX word/settings.xml is not WordprocessingML.',
    );
  }
  replaceSettingsOnOff(
    document,
    root,
    'mirrorMargins',
    pageMargins[0]?.mirrorMargins,
  );
  replaceSettingsOnOff(
    document,
    root,
    'gutterAtTop',
    pageMargins[0]?.gutterAtTop,
  );
  archive.file(
    'word/settings.xml',
    new XMLSerializer().serializeToString(document),
  );
}

function assertConsistentPageMarginSettings(
  pageMargins: readonly WorkDocumentPageMargins[],
): void {
  const first = pageMargins[0];
  if (!first) return;
  if (
    pageMargins.some(
      (value) =>
        value.mirrorMargins !== first.mirrorMargins ||
        value.gutterAtTop !== first.gutterAtTop,
    )
  ) {
    throw new Error(
      'Document-wide mirrorMargins and gutterAtTop values must match across all sections.',
    );
  }
}

function replaceSettingsOnOff(
  document: Document,
  root: Element,
  localName: 'gutterAtTop' | 'mirrorMargins',
  value: boolean | undefined,
): void {
  for (const element of directChildren(root, localName).filter(
    (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
  )) {
    element.remove();
  }
  if (value === undefined) return;
  const prefix = xmlNamespacePrefix(root, WORD_NAMESPACE) ?? 'w';
  const element = document.createElementNS(
    WORD_NAMESPACE,
    `${prefix}:${localName}`,
  );
  setWordAttribute(document, element, 'val', value ? '1' : '0');
  insertOrderedProperty(root, element, SETTINGS_ORDER);
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
  insertOrderedProperty(parent, element, SECTION_PROPERTY_ORDER);
}

function insertOrderedProperty<const T extends readonly string[]>(
  parent: Element,
  element: Element,
  order: T,
): void {
  const targetIndex = order.indexOf(element.localName as T[number]);
  const next = directChildren(parent).find((candidate) => {
    const index = order.indexOf(candidate.localName as T[number]);
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
