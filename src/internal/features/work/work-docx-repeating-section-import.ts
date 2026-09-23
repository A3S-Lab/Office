import {
  normalizeRepeatingSectionItemProperties,
  normalizeRepeatingSectionProperties,
  repeatingSectionDomAttributes,
  repeatingSectionItemDomAttributes,
  type WorkDocumentRepeatingSectionItemProperties,
  type WorkDocumentRepeatingSectionProperties,
} from './work-document-repeating-section';
import {
  CONTENT_CONTROL_WORD_2012_NAMESPACE,
  hasContentControlRelationshipReference,
  isDocxWordElement,
} from './work-docx-note-comment-content-control-xml';
import {
  XML_NAMESPACE,
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { descendants, xmlNamespacePrefix } from './work-ooxml-package';

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const STRICT_WORD_NAMESPACE =
  'http://purl.oclc.org/ooxml/wordprocessingml/main';
const WORD_NAMESPACES = new Set([WORD_NAMESPACE, STRICT_WORD_NAMESPACE]);
const CONTENT_CONTROL_WORD_2010_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2010/wordml';
const MARKUP_COMPATIBILITY_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';
const MAX_IMPORTED_REPEATING_SECTIONS = 256;
const MAX_ITEMS_PER_SECTION = 64;
const MAX_ITEM_TEXT = 100_000;
const RUN_CONTENT_NAMES = new Set([
  'br',
  'cr',
  'noBreakHyphen',
  'softHyphen',
  'tab',
  't',
]);
const SECTION_PROPERTY_NAMES = new Set(['alias', 'id', 'lock', 'tag']);
const LOCK_VALUES = new Set([
  'contentLocked',
  'sdtContentLocked',
  'sdtLocked',
  'unlocked',
]);
const APPEARANCE_VALUES = new Set(['boundingBox', 'hidden', 'tags']);

export interface ImportedDocxRepeatingSectionItemMarker {
  start: string;
  end: string;
  properties: WorkDocumentRepeatingSectionItemProperties;
}

export interface ImportedDocxRepeatingSectionMarker {
  start: string;
  end: string;
  properties: WorkDocumentRepeatingSectionProperties;
  items: ImportedDocxRepeatingSectionItemMarker[];
}

export interface ImportedDocxRepeatingSectionMarkers {
  sections: ImportedDocxRepeatingSectionMarker[];
  unsupported: number;
}

export interface DocxRepeatingSectionInspection {
  supported: number;
  unsupported: number;
}

/**
 * Inspects body-level repeating sections without mutating the package.
 * Nested or multi-paragraph items stay unsupported.
 */
export function inspectDocxRepeatingSections(
  document: Document,
): DocxRepeatingSectionInspection {
  let supported = 0;
  let unsupported = 0;
  for (const control of repeatingSectionCandidates(document)) {
    if (readRepeatingSection(control, supported + 1)) supported += 1;
    else unsupported += 1;
  }
  return { supported, unsupported };
}

/**
 * Replaces admitted repeating sections with paragraph markers so Mammoth can
 * convert item bodies, then wraps them back into editable HTML nodes.
 */
export function markDocxRepeatingSections(
  document: Document,
): ImportedDocxRepeatingSectionMarkers {
  const sections: ImportedDocxRepeatingSectionMarker[] = [];
  let unsupported = 0;
  for (const control of repeatingSectionCandidates(document)) {
    const parsed = readRepeatingSection(control, sections.length + 1);
    if (!parsed) {
      unsupported += 1;
      continue;
    }
    if (sections.length >= MAX_IMPORTED_REPEATING_SECTIONS) {
      throw new Error('Imported DOCX exceeds the repeating-section limit.');
    }
    const markers = allocateSectionMarkers(
      document,
      sections.length + 1,
      parsed.items.length,
    );
    if (!replaceRepeatingSectionWithMarkers(document, control, markers, parsed)) {
      unsupported += 1;
      continue;
    }
    sections.push({
      start: markers.start,
      end: markers.end,
      properties: parsed.properties,
      items: parsed.items.map((item, index) => ({
        start: markers.itemStarts[index]!,
        end: markers.itemEnds[index]!,
        properties: item.properties,
      })),
    });
  }
  return { sections, unsupported };
}

export function applyImportedDocxRepeatingSectionMarkers(
  document: Document,
  markers: ImportedDocxRepeatingSectionMarkers,
): void {
  if (!markers.sections.length) return;
  for (const section of markers.sections) {
    const sectionStart = findMarkerElement(document.body, section.start);
    const sectionEnd = findMarkerElement(document.body, section.end);
    if (!sectionStart || !sectionEnd) continue;

    const sectionWrapper = document.createElement('div');
    for (const [name, value] of Object.entries(
      repeatingSectionDomAttributes(section.properties),
    )) {
      if (value !== undefined) sectionWrapper.setAttribute(name, value);
    }
    sectionWrapper.className = 'work-document-repeating-section';
    sectionWrapper.setAttribute('role', 'group');
    sectionWrapper.setAttribute(
      'aria-label',
      section.properties.alias || section.properties.tag || '重复部分',
    );

    for (const item of section.items) {
      const itemStart = findMarkerElement(document.body, item.start);
      const itemEnd = findMarkerElement(document.body, item.end);
      if (!itemStart || !itemEnd) continue;
      const itemWrapper = document.createElement('div');
      for (const [name, value] of Object.entries(
        repeatingSectionItemDomAttributes(item.properties),
      )) {
        if (value !== undefined) itemWrapper.setAttribute(name, value);
      }
      itemWrapper.className = 'work-document-repeating-section-item';
      moveNodesBetween(itemStart, itemEnd, itemWrapper);
      removeMarkerElement(itemStart);
      removeMarkerElement(itemEnd);
      sectionWrapper.append(itemWrapper);
    }

    const parent = sectionStart.parentNode;
    if (!parent) continue;
    parent.insertBefore(sectionWrapper, sectionStart);
    removeMarkerElement(sectionStart);
    removeMarkerElement(sectionEnd);
  }
  document.body.normalize();
}

function findMarkerElement(
  root: ParentNode,
  marker: string,
): Element | null {
  const textNodes: Text[] = [];
  const collect = (node: Node) => {
    if (node.nodeType === Node.TEXT_NODE) textNodes.push(node as Text);
    for (const child of Array.from(node.childNodes)) collect(child);
  };
  collect(root);
  for (const text of textNodes) {
    if (!text.data.includes(marker)) continue;
    let current: Element | null =
      text.parentElement instanceof Element ? text.parentElement : null;
    while (current && current !== root) {
      if (current.tagName === 'P' || current.tagName === 'DIV') return current;
      current = current.parentElement;
    }
    return text.parentElement;
  }
  return null;
}

function moveNodesBetween(
  start: Element,
  end: Element,
  target: HTMLElement,
): void {
  let current = start.nextSibling;
  while (current && current !== end) {
    const next = current.nextSibling;
    target.append(current);
    current = next;
  }
}

function removeMarkerElement(element: Element): void {
  element.remove();
}

export function hasImportedDocxRepeatingSectionMarkers(
  markers: ImportedDocxRepeatingSectionMarkers,
): boolean {
  return markers.sections.length > 0;
}

interface ParsedRepeatingSectionItem {
  properties: WorkDocumentRepeatingSectionItemProperties;
  paragraph: Element;
}

interface ParsedRepeatingSection {
  properties: WorkDocumentRepeatingSectionProperties;
  items: ParsedRepeatingSectionItem[];
}

interface SectionMarkerSet {
  start: string;
  end: string;
  itemStarts: string[];
  itemEnds: string[];
}

function repeatingSectionCandidates(document: Document): Element[] {
  return descendants(document, 'sdt').filter((control) => {
    if (!isDocxWordElement(control) || hasContentControlAncestor(control)) {
      return false;
    }
    return isBodyLevelControl(control) && hasRepeatingSectionProperty(control);
  });
}

function hasContentControlAncestor(control: Element): boolean {
  let current = control.parentElement;
  while (current) {
    if (isDocxWordElement(current) && current.localName === 'sdt') return true;
    current = current.parentElement;
  }
  return false;
}

function isBodyLevelControl(control: Element): boolean {
  const parent = control.parentElement;
  return Boolean(
    parent &&
      isDocxWordElement(parent) &&
      (parent.localName === 'body' || parent.localName === 'tc'),
  );
}

function hasRepeatingSectionProperty(control: Element): boolean {
  const properties = Array.from(control.children).find(
    (child) => isDocxWordElement(child) && child.localName === 'sdtPr',
  );
  if (!properties) return false;
  return Array.from(properties.children).some(
    (child) =>
      child.namespaceURI === CONTENT_CONTROL_WORD_2012_NAMESPACE &&
      child.localName === 'repeatingSection',
  );
}

function readRepeatingSection(
  control: Element,
  sequence: number,
): ParsedRepeatingSection | null {
  if (
    !isDocxWordElement(control) ||
    control.localName !== 'sdt' ||
    hasContentControlRelationshipReference(control) ||
    !validContainerAttributes(control)
  ) {
    return null;
  }
  const children = Array.from(control.children);
  const propertiesElements = children.filter(
    (child) => isDocxWordElement(child) && child.localName === 'sdtPr',
  );
  const contentElements = children.filter(
    (child) => isDocxWordElement(child) && child.localName === 'sdtContent',
  );
  if (
    propertiesElements.length !== 1 ||
    contentElements.length !== 1 ||
    children.some(
      (child) =>
        !isDocxWordElement(child) ||
        (child.localName !== 'sdtPr' &&
          child.localName !== 'sdtContent' &&
          child.localName !== 'sdtEndPr'),
    )
  ) {
    return null;
  }
  const properties = readSectionProperties(propertiesElements[0]!, sequence);
  if (!properties) return null;
  const items: ParsedRepeatingSectionItem[] = [];
  for (const child of Array.from(contentElements[0]!.children)) {
    if (!isDocxWordElement(child) || child.localName !== 'sdt') return null;
    const item = readSectionItem(child, items.length + 1);
    if (!item) return null;
    items.push(item);
    if (items.length > MAX_ITEMS_PER_SECTION) return null;
  }
  if (!items.length) return null;
  return { properties, items };
}

function readSectionProperties(
  properties: Element,
  sequence: number,
): WorkDocumentRepeatingSectionProperties | null {
  if (!validContainerAttributes(properties)) return null;
  const groups = new Map<string, Element[]>();
  let appearance: WorkDocumentRepeatingSectionProperties['appearance'] =
    'boundingBox';
  let color: string | null = null;
  let hasRepeating = false;
  for (const child of Array.from(properties.children)) {
    const namespace = child.namespaceURI ?? '';
    if (isDocxWordElement(child)) {
      if (!SECTION_PROPERTY_NAMES.has(child.localName)) return null;
      const matches = groups.get(child.localName) ?? [];
      matches.push(child);
      groups.set(child.localName, matches);
      continue;
    }
    if (namespace === CONTENT_CONTROL_WORD_2012_NAMESPACE) {
      if (child.localName === 'repeatingSection') {
        if (hasRepeating || child.children.length > 0) return null;
        if (!validNamespacedLeaf(child, new Set(), CONTENT_CONTROL_WORD_2012_NAMESPACE)) {
          return null;
        }
        hasRepeating = true;
        continue;
      }
      if (child.localName === 'appearance') {
        const matches = groups.get('w15:appearance') ?? [];
        matches.push(child);
        groups.set('w15:appearance', matches);
        continue;
      }
      if (child.localName === 'color') {
        const matches = groups.get('w15:color') ?? [];
        matches.push(child);
        groups.set('w15:color', matches);
        continue;
      }
      return null;
    }
    if (namespace === MARKUP_COMPATIBILITY_NAMESPACE) continue;
    return null;
  }
  if (!hasRepeating) return null;
  if (Array.from(groups.values()).some((items) => items.length > 1)) return null;
  const id = groups.get('id')?.[0];
  if (id && !validLeaf(id, new Set(['val']))) return null;
  const nativeId = id ? parseNativeId(wordAttribute(id, 'val')) : null;
  if (id && nativeId === null) return null;
  const alias = readStringProperty(groups.get('alias')?.[0]);
  const tag = readStringProperty(groups.get('tag')?.[0]);
  if (alias === null || tag === null) return null;
  const lockElement = groups.get('lock')?.[0];
  if (lockElement && !validLeaf(lockElement, new Set(['val']))) return null;
  const lockValue = lockElement ? wordAttribute(lockElement, 'val') : null;
  const lock = lockValue ?? 'unlocked';
  if (!LOCK_VALUES.has(lock)) return null;
  const appearanceElement = groups.get('w15:appearance')?.[0];
  if (appearanceElement) {
    if (
      !validNamespacedLeaf(
        appearanceElement,
        new Set(['val']),
        CONTENT_CONTROL_WORD_2012_NAMESPACE,
      )
    ) {
      return null;
    }
    const value = namespacedAttribute(
      appearanceElement,
      'val',
      CONTENT_CONTROL_WORD_2012_NAMESPACE,
    );
    if (!value || !APPEARANCE_VALUES.has(value)) return null;
    appearance = value as WorkDocumentRepeatingSectionProperties['appearance'];
  }
  const colorElement = groups.get('w15:color')?.[0];
  if (colorElement) {
    if (
      !validNamespacedLeaf(
        colorElement,
        new Set(['val']),
        CONTENT_CONTROL_WORD_2012_NAMESPACE,
      )
    ) {
      return null;
    }
    const value = namespacedAttribute(
      colorElement,
      'val',
      CONTENT_CONTROL_WORD_2012_NAMESPACE,
    );
    if (!value || !/^#[0-9a-fA-F]{6}$/.test(value) && !/^[0-9a-fA-F]{6}$/.test(value)) {
      return null;
    }
    color = value.startsWith('#') ? value.toLowerCase() : `#${value.toLowerCase()}`;
  }
  return normalizeRepeatingSectionProperties({
    id: `docx-repeating-section-${sequence}`,
    nativeId,
    alias,
    tag,
    lock: lock as WorkDocumentRepeatingSectionProperties['lock'],
    appearance,
    color,
  });
}

function readSectionItem(
  control: Element,
  sequence: number,
): ParsedRepeatingSectionItem | null {
  if (
    !isDocxWordElement(control) ||
    control.localName !== 'sdt' ||
    hasContentControlRelationshipReference(control) ||
    !validContainerAttributes(control)
  ) {
    return null;
  }
  const children = Array.from(control.children);
  const propertiesElements = children.filter(
    (child) => isDocxWordElement(child) && child.localName === 'sdtPr',
  );
  const contentElements = children.filter(
    (child) => isDocxWordElement(child) && child.localName === 'sdtContent',
  );
  if (propertiesElements.length !== 1 || contentElements.length !== 1) {
    return null;
  }
  let hasItem = false;
  let nativeId: number | null = null;
  for (const child of Array.from(propertiesElements[0]!.children)) {
    if (isDocxWordElement(child) && child.localName === 'id') {
      if (!validLeaf(child, new Set(['val']))) return null;
      nativeId = parseNativeId(wordAttribute(child, 'val'));
      if (nativeId === null) return null;
      continue;
    }
    if (
      child.namespaceURI === CONTENT_CONTROL_WORD_2012_NAMESPACE &&
      child.localName === 'repeatingSectionItem'
    ) {
      if (hasItem || child.children.length > 0) return null;
      if (
        !validNamespacedLeaf(
          child,
          new Set(),
          CONTENT_CONTROL_WORD_2012_NAMESPACE,
        )
      ) {
        return null;
      }
      hasItem = true;
      continue;
    }
    if (child.namespaceURI === MARKUP_COMPATIBILITY_NAMESPACE) continue;
    return null;
  }
  if (!hasItem) return null;
  const contentChildren = Array.from(contentElements[0]!.children);
  if (
    contentChildren.length !== 1 ||
    !isDocxWordElement(contentChildren[0]!) ||
    contentChildren[0]!.localName !== 'p'
  ) {
    return null;
  }
  const paragraph = contentChildren[0]!;
  if (!validParagraphRuns(paragraph)) return null;
  const text = Array.from(paragraph.children)
    .filter((child) => isDocxWordElement(child) && child.localName === 'r')
    .map(runText)
    .join('');
  if (text.length > MAX_ITEM_TEXT) return null;
  return {
    properties: normalizeRepeatingSectionItemProperties({
      id: `docx-repeating-section-item-${sequence}`,
      nativeId,
    }),
    paragraph,
  };
}

function replaceRepeatingSectionWithMarkers(
  document: Document,
  control: Element,
  markers: SectionMarkerSet,
  parsed: ParsedRepeatingSection,
): boolean {
  const parent = control.parentNode;
  if (!parent) return false;
  const fragment = document.createDocumentFragment();
  fragment.append(createMarkerParagraph(document, markers.start));
  for (const [index, item] of parsed.items.entries()) {
    fragment.append(createMarkerParagraph(document, markers.itemStarts[index]!));
    fragment.append(item.paragraph.cloneNode(true));
    fragment.append(createMarkerParagraph(document, markers.itemEnds[index]!));
  }
  fragment.append(createMarkerParagraph(document, markers.end));
  parent.replaceChild(fragment, control);
  return true;
}

function createMarkerParagraph(document: Document, marker: string): Element {
  const root = document.documentElement;
  const namespace = WORD_NAMESPACES.has(root.namespaceURI ?? '')
    ? (root.namespaceURI ?? WORD_NAMESPACE)
    : WORD_NAMESPACE;
  const prefix = xmlNamespacePrefix(root, namespace) ?? 'w';
  const paragraph = document.createElementNS(namespace, `${prefix}:p`);
  const run = document.createElementNS(namespace, `${prefix}:r`);
  const text = document.createElementNS(namespace, `${prefix}:t`);
  text.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  text.textContent = marker;
  run.append(text);
  paragraph.append(run);
  return paragraph;
}

function allocateSectionMarkers(
  document: Document,
  initialIndex: number,
  itemCount: number,
): SectionMarkerSet {
  const occupiedText = document.documentElement.textContent ?? '';
  let index = initialIndex;
  for (let attempts = 0; attempts < 1_000_000; attempts += 1) {
    const start = `__A3S_WORK_REPEATING_SECTION_START_${index}__`;
    const end = `__A3S_WORK_REPEATING_SECTION_END_${index}__`;
    const itemStarts = Array.from(
      { length: itemCount },
      (_, itemIndex) =>
        `__A3S_WORK_REPEATING_SECTION_ITEM_START_${index}_${itemIndex + 1}__`,
    );
    const itemEnds = Array.from(
      { length: itemCount },
      (_, itemIndex) =>
        `__A3S_WORK_REPEATING_SECTION_ITEM_END_${index}_${itemIndex + 1}__`,
    );
    const candidates = [start, end, ...itemStarts, ...itemEnds];
    if (candidates.every((marker) => !occupiedText.includes(marker))) {
      return { start, end, itemStarts, itemEnds };
    }
    index += 1;
  }
  throw new Error('Imported DOCX repeating-section marker space is exhausted.');
}

function validParagraphRuns(paragraph: Element): boolean {
  if (!validContainerAttributes(paragraph)) return false;
  let hasRun = false;
  for (const child of Array.from(paragraph.children)) {
    if (!isDocxWordElement(child)) return false;
    if (child.localName === 'pPr') {
      if (child.children.length > 0) return false;
      continue;
    }
    if (child.localName !== 'r' || !validRun(child)) return false;
    hasRun = true;
  }
  return hasRun;
}

function validRun(run: Element): boolean {
  if (!validContainerAttributes(run)) return false;
  for (const child of Array.from(run.children)) {
    if (!isDocxWordElement(child)) return false;
    if (child.localName === 'rPr') {
      if (
        Array.from(child.children).some(
          (item) =>
            !isDocxWordElement(item) ||
            !['b', 'i', 'u', 'strike', 'color', 'sz', 'rFonts'].includes(
              item.localName,
            ),
        )
      ) {
        return false;
      }
      continue;
    }
    if (!RUN_CONTENT_NAMES.has(child.localName)) return false;
  }
  return true;
}

function runText(run: Element): string {
  let value = '';
  for (const child of Array.from(run.children)) {
    if (!isDocxWordElement(child)) continue;
    if (child.localName === 't') value += child.textContent ?? '';
    else if (child.localName === 'tab') value += '\t';
    else if (child.localName === 'br' || child.localName === 'cr')
      value += '\n';
    else if (child.localName === 'noBreakHyphen') value += '\u2011';
    else if (child.localName === 'softHyphen') value += '\u00ad';
  }
  return value;
}

function parseNativeId(value: string | null): number | null {
  if (value === null || !/^-?\d+$/.test(value)) return null;
  const parsed = Number(value);
  if (
    !Number.isSafeInteger(parsed) ||
    parsed < -2_147_483_648 ||
    parsed > 2_147_483_647
  ) {
    return null;
  }
  return parsed;
}

function readStringProperty(element: Element | undefined): string | null {
  if (!element) return '';
  if (!validLeaf(element, new Set(['val']))) return null;
  const value = wordAttribute(element, 'val');
  if (value === null || /[\u0000-\u001f\u007f]/u.test(value)) return null;
  return value.slice(0, 255);
}

function wordAttribute(element: Element, localName: string): string | null {
  const matches = Array.from(element.attributes).filter(
    (item) =>
      xmlAttributeLocalName(item) === localName &&
      xmlAttributeNamespace(element, item) === element.namespaceURI &&
      WORD_NAMESPACES.has(element.namespaceURI ?? ''),
  );
  return matches.length === 1 ? matches[0].value : null;
}

function namespacedAttribute(
  element: Element,
  localName: string,
  namespace: string,
): string | null {
  const matches = Array.from(element.attributes).filter(
    (item) =>
      xmlAttributeLocalName(item) === localName &&
      xmlAttributeNamespace(element, item) === namespace,
  );
  return matches.length === 1 ? matches[0].value : null;
}

function validLeaf(element: Element, allowed: Set<string>): boolean {
  if (element.children.length > 0) return false;
  return Array.from(element.attributes).every((item) => {
    const namespace = xmlAttributeNamespace(element, item);
    const localName = xmlAttributeLocalName(item);
    if (namespace === XML_NAMESPACE && localName === 'space') return true;
    return (
      namespace === element.namespaceURI &&
      WORD_NAMESPACES.has(namespace ?? '') &&
      allowed.has(localName)
    );
  });
}

function validNamespacedLeaf(
  element: Element,
  allowed: Set<string>,
  namespace: string,
): boolean {
  if (element.children.length > 0) return false;
  return Array.from(element.attributes).every((item) => {
    const attributeNamespace = xmlAttributeNamespace(element, item);
    const localName = xmlAttributeLocalName(item);
    if (attributeNamespace === XML_NAMESPACE && localName === 'space')
      return true;
    return attributeNamespace === namespace && allowed.has(localName);
  });
}

function validContainerAttributes(element: Element): boolean {
  return Array.from(element.attributes).every((item) => {
    const namespace = xmlAttributeNamespace(element, item);
    return (
      namespace === element.namespaceURI ||
      namespace === XML_NAMESPACE ||
      namespace === MARKUP_COMPATIBILITY_NAMESPACE ||
      namespace === CONTENT_CONTROL_WORD_2010_NAMESPACE ||
      namespace === CONTENT_CONTROL_WORD_2012_NAMESPACE
    );
  });
}
