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
import {
  documentPageGeometryForLayout,
  normalizeDocumentPaperSource,
} from './work-document-page-size';
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
  patchSectionPageSetup(document, sections);
  patchSectionPageMargins(document, sections);
  patchSectionPageBorders(document, sections);
  patchSectionDocumentGrids(document, sections);
  patchSectionLnNumTypes(document, sections);
  patchSectionPgNumTypes(document, sections);
  patchSectionFormProts(document, sections);
  patchSectionVerticalAligns(document, sections);
  patchSectionNoEndnotes(document, sections);
  patchSectionTextDirections(document, sections);
  patchSectionBidis(document, sections);
  patchSectionFootnotePrs(document, sections);
  patchSectionEndnotePrs(document, sections);
  archive.file(
    'word/document.xml',
    new XMLSerializer().serializeToString(document),
  );
  await patchDocumentPageMarginSettings(archive, sections);
  return archive.generateAsync({ type: 'arraybuffer' });
}

function patchSectionPageSetup(
  document: Document,
  sections: readonly WorkDocumentSection[],
): void {
  const sectionProperties = effectiveSectionProperties(document);
  if (sectionProperties.length !== sections.length) {
    throw new Error(
      `Generated DOCX has ${sectionProperties.length} section properties for ${sections.length} exact page-setup section(s).`,
    );
  }
  const prefix =
    xmlNamespacePrefix(document.documentElement, WORD_NAMESPACE) ?? 'w';
  for (const [index, properties] of sectionProperties.entries()) {
    const layout = sections[index]?.layout;
    if (!layout) continue;
    for (const element of directChildren(properties, 'pgSz').filter(
      (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
    )) {
      element.remove();
    }
    const geometry = documentPageGeometryForLayout(layout);
    const pageSize = document.createElementNS(WORD_NAMESPACE, `${prefix}:pgSz`);
    setWordAttribute(document, pageSize, 'w', String(geometry.width));
    setWordAttribute(document, pageSize, 'h', String(geometry.height));
    if (geometry.orientation) {
      setWordAttribute(document, pageSize, 'orient', geometry.orientation);
    }
    if (geometry.code !== undefined) {
      setWordAttribute(document, pageSize, 'code', String(geometry.code));
    }
    insertSectionProperty(properties, pageSize);

    const paperSource = normalizeDocumentPaperSource(layout.paperSource);
    if (!paperSource) continue;
    for (const element of directChildren(properties, 'paperSrc').filter(
      (candidate) => candidate.namespaceURI === WORD_NAMESPACE,
    )) {
      element.remove();
    }
    const source = document.createElementNS(
      WORD_NAMESPACE,
      `${prefix}:paperSrc`,
    );
    if (paperSource.first !== undefined) {
      setWordAttribute(document, source, 'first', String(paperSource.first));
    }
    if (paperSource.other !== undefined) {
      setWordAttribute(document, source, 'other', String(paperSource.other));
    }
    insertSectionProperty(properties, source);
  }
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

function patchSectionLnNumTypes(
  document: Document,
  sections: readonly WorkDocumentSection[],
): void {
  const sectionProperties = effectiveSectionProperties(document);
  for (const [index, properties] of sectionProperties.entries()) {
    for (const existing of directChildren(properties, 'lnNumType')) {
      existing.remove();
    }
    const value = sections[index]?.layout.lnNumType;
    if (!value) continue;
    const lnNumType = document.createElementNS(WORD_NAMESPACE, 'w:lnNumType');
    if (value.countBy !== undefined) {
      lnNumType.setAttributeNS(
        WORD_NAMESPACE,
        'w:countBy',
        String(value.countBy),
      );
    }
    if (value.start !== undefined) {
      lnNumType.setAttributeNS(WORD_NAMESPACE, 'w:start', String(value.start));
    }
    if (value.distance !== undefined) {
      lnNumType.setAttributeNS(
        WORD_NAMESPACE,
        'w:distance',
        String(value.distance),
      );
    }
    if (value.restart !== undefined) {
      lnNumType.setAttributeNS(WORD_NAMESPACE, 'w:restart', value.restart);
    }
    insertSectionProperty(properties, lnNumType);
  }
}

function patchSectionPgNumTypes(
  document: Document,
  sections: readonly WorkDocumentSection[],
): void {
  const sectionProperties = effectiveSectionProperties(document);
  for (const [index, properties] of sectionProperties.entries()) {
    for (const existing of directChildren(properties, 'pgNumType')) {
      existing.remove();
    }
    const layout = sections[index]?.layout;
    const value =
      layout?.pgNumType ??
      (layout?.pageNumberStart !== undefined
        ? { start: layout.pageNumberStart }
        : undefined);
    if (!value) continue;
    const pgNumType = document.createElementNS(WORD_NAMESPACE, 'w:pgNumType');
    if (value.fmt !== undefined) {
      pgNumType.setAttributeNS(WORD_NAMESPACE, 'w:fmt', value.fmt);
    }
    if (value.start !== undefined) {
      pgNumType.setAttributeNS(WORD_NAMESPACE, 'w:start', String(value.start));
    }
    insertSectionProperty(properties, pgNumType);
  }
}

function patchSectionFormProts(
  document: Document,
  sections: readonly WorkDocumentSection[],
): void {
  const sectionProperties = effectiveSectionProperties(document);
  for (const [index, properties] of sectionProperties.entries()) {
    for (const existing of directChildren(properties, 'formProt')) {
      existing.remove();
    }
    const formProt = sections[index]?.layout.formProt;
    if (formProt === undefined) continue;
    const element = document.createElementNS(WORD_NAMESPACE, 'w:formProt');
    if (!formProt) {
      element.setAttributeNS(WORD_NAMESPACE, 'w:val', '0');
    }
    insertSectionProperty(properties, element);
  }
}

function patchSectionVerticalAligns(
  document: Document,
  sections: readonly WorkDocumentSection[],
): void {
  const sectionProperties = effectiveSectionProperties(document);
  for (const [index, properties] of sectionProperties.entries()) {
    for (const existing of directChildren(properties, 'vAlign')) {
      existing.remove();
    }
    const verticalAlign = sections[index]?.layout.verticalAlign;
    if (verticalAlign === undefined) continue;
    const element = document.createElementNS(WORD_NAMESPACE, 'w:vAlign');
    element.setAttributeNS(WORD_NAMESPACE, 'w:val', verticalAlign);
    insertSectionProperty(properties, element);
  }
}

function patchSectionNoEndnotes(
  document: Document,
  sections: readonly WorkDocumentSection[],
): void {
  const sectionProperties = effectiveSectionProperties(document);
  for (const [index, properties] of sectionProperties.entries()) {
    for (const existing of directChildren(properties, 'noEndnote')) {
      existing.remove();
    }
    const noEndnote = sections[index]?.layout.noEndnote;
    if (noEndnote === undefined) continue;
    const element = document.createElementNS(WORD_NAMESPACE, 'w:noEndnote');
    if (!noEndnote) {
      element.setAttributeNS(WORD_NAMESPACE, 'w:val', '0');
    }
    insertSectionProperty(properties, element);
  }
}

function patchSectionTextDirections(
  document: Document,
  sections: readonly WorkDocumentSection[],
): void {
  const sectionProperties = effectiveSectionProperties(document);
  for (const [index, properties] of sectionProperties.entries()) {
    for (const existing of directChildren(properties, 'textDirection')) {
      existing.remove();
    }
    const textDirection = sections[index]?.layout.textDirection;
    if (textDirection === undefined) continue;
    const element = document.createElementNS(WORD_NAMESPACE, 'w:textDirection');
    element.setAttributeNS(WORD_NAMESPACE, 'w:val', textDirection);
    insertSectionProperty(properties, element);
  }
}

function patchSectionBidis(
  document: Document,
  sections: readonly WorkDocumentSection[],
): void {
  const sectionProperties = effectiveSectionProperties(document);
  for (const [index, properties] of sectionProperties.entries()) {
    for (const existing of directChildren(properties, 'bidi')) {
      existing.remove();
    }
    const bidi = sections[index]?.layout.bidi;
    if (bidi === undefined) continue;
    const element = document.createElementNS(WORD_NAMESPACE, 'w:bidi');
    if (!bidi) {
      element.setAttributeNS(WORD_NAMESPACE, 'w:val', '0');
    }
    insertSectionProperty(properties, element);
  }
}

function patchSectionFootnotePrs(
  document: Document,
  sections: readonly WorkDocumentSection[],
): void {
  const sectionProperties = effectiveSectionProperties(document);
  for (const [index, properties] of sectionProperties.entries()) {
    for (const existing of directChildren(properties, 'footnotePr')) {
      existing.remove();
    }
    const value = sections[index]?.layout.footnotePr;
    if (value === undefined) continue;
    const footnotePr = document.createElementNS(WORD_NAMESPACE, 'w:footnotePr');
    if (value.pos !== undefined) {
      const pos = document.createElementNS(WORD_NAMESPACE, 'w:pos');
      pos.setAttributeNS(WORD_NAMESPACE, 'w:val', value.pos);
      footnotePr.append(pos);
    }
    if (value.numFmt !== undefined) {
      const numFmt = document.createElementNS(WORD_NAMESPACE, 'w:numFmt');
      numFmt.setAttributeNS(WORD_NAMESPACE, 'w:val', value.numFmt);
      footnotePr.append(numFmt);
    }
    if (value.numStart !== undefined) {
      const numStart = document.createElementNS(WORD_NAMESPACE, 'w:numStart');
      numStart.setAttributeNS(WORD_NAMESPACE, 'w:val', String(value.numStart));
      footnotePr.append(numStart);
    }
    if (value.numRestart !== undefined) {
      const numRestart = document.createElementNS(
        WORD_NAMESPACE,
        'w:numRestart',
      );
      numRestart.setAttributeNS(WORD_NAMESPACE, 'w:val', value.numRestart);
      footnotePr.append(numRestart);
    }
    insertSectionProperty(properties, footnotePr);
  }
}

function patchSectionEndnotePrs(
  document: Document,
  sections: readonly WorkDocumentSection[],
): void {
  const sectionProperties = effectiveSectionProperties(document);
  for (const [index, properties] of sectionProperties.entries()) {
    for (const existing of directChildren(properties, 'endnotePr')) {
      existing.remove();
    }
    const value = sections[index]?.layout.endnotePr;
    if (value === undefined) continue;
    const endnotePr = document.createElementNS(WORD_NAMESPACE, 'w:endnotePr');
    if (value.pos !== undefined) {
      const pos = document.createElementNS(WORD_NAMESPACE, 'w:pos');
      pos.setAttributeNS(WORD_NAMESPACE, 'w:val', value.pos);
      endnotePr.append(pos);
    }
    if (value.numFmt !== undefined) {
      const numFmt = document.createElementNS(WORD_NAMESPACE, 'w:numFmt');
      numFmt.setAttributeNS(WORD_NAMESPACE, 'w:val', value.numFmt);
      endnotePr.append(numFmt);
    }
    if (value.numStart !== undefined) {
      const numStart = document.createElementNS(WORD_NAMESPACE, 'w:numStart');
      numStart.setAttributeNS(WORD_NAMESPACE, 'w:val', String(value.numStart));
      endnotePr.append(numStart);
    }
    if (value.numRestart !== undefined) {
      const numRestart = document.createElementNS(
        WORD_NAMESPACE,
        'w:numRestart',
      );
      numRestart.setAttributeNS(WORD_NAMESPACE, 'w:val', value.numRestart);
      endnotePr.append(numRestart);
    }
    insertSectionProperty(properties, endnotePr);
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
