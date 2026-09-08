import {
  normalizeDocumentRowGridBefore,
  serializeDocumentRowFormatting,
  type DocumentRowFormattingHeight,
} from './work-document-row-format-changes';
import {
  normalizeDocumentTableAlignment,
  type DocumentTableAlignment,
} from './work-document-table-geometry';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { attribute, directChildren } from './work-ooxml-package';

const MAX_REVISION_DATE_LENGTH = 64;
const REVISION_ATTRIBUTES = new Set(['id', 'author', 'date']);
const RELATIONSHIP_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  'http://purl.oclc.org/ooxml/officeDocument/relationships',
  'http://schemas.openxmlformats.org/package/2006/relationships',
]);
const SUPPORTED_PRIOR_CHILDREN = new Set([
  'cantSplit',
  'tblHeader',
  'trHeight',
  'hidden',
  'jc',
  'gridBefore',
]);
const PIXELS_PER_TWIP = 96 / 1440;

export interface SupportedDocxRowFormattingChange {
  id: string;
  author: string;
  date: string;
  before: string;
}

/**
 * Relationship-free `w:trPrChange` whose prior snapshot contains only
 * `w:cantSplit`, `w:tblHeader`, `w:trHeight`, `w:hidden`, `w:jc`, and/or
 * `w:gridBefore`. Broader row property sets stay on the opaque OMML path.
 */
export function isSupportedDocxRowFormattingChange(change: Element): boolean {
  return supportedRowFormattingChange(change) !== null;
}

export function supportedDocxRowFormattingChangeFromProperties(
  rowProperties: Element | null | undefined,
): SupportedDocxRowFormattingChange | null {
  if (!rowProperties) return null;
  const changes = directChildren(rowProperties, 'trPrChange').filter(
    (element) => element.namespaceURI === rowProperties.namespaceURI,
  );
  if (changes.length !== 1) return null;
  return supportedRowFormattingChange(changes[0] ?? null);
}

export function applyDocumentRowFormattingChangeToElement(
  row: HTMLTableRowElement,
  change: SupportedDocxRowFormattingChange,
): void {
  row.dataset.documentChange = 'true';
  row.dataset.changeKind = 'row-formatting';
  row.dataset.changeId = change.id;
  row.dataset.changeAuthor = change.author;
  row.dataset.changeDate = change.date;
  row.dataset.changeBefore = change.before;
}

function supportedRowFormattingChange(
  change: Element | null,
): SupportedDocxRowFormattingChange | null {
  if (
    !change ||
    change.localName !== 'trPrChange' ||
    !DOCX_WORDPROCESSING_NAMESPACES.has(change.namespaceURI ?? '')
  ) {
    return null;
  }
  if (hasRelationshipBindings(change) || hasUnsupportedRevisionAttributes(change)) {
    return null;
  }
  const id = wordAttribute(change, 'id')?.trim() ?? '';
  const author = wordAttribute(change, 'author')?.trim() ?? '';
  const rawDate = wordAttribute(change, 'date');
  if (
    !/^\+?\d{1,10}$/.test(id) ||
    !author ||
    author.length > 255 ||
    /[\u0000-\u001f\u007f]/.test(author) ||
    (rawDate !== null && rawDate.length > MAX_REVISION_DATE_LENGTH) ||
    (rawDate !== null && !Number.isFinite(Date.parse(rawDate)))
  ) {
    return null;
  }
  const priors = Array.from(change.children).filter(
    (child) =>
      child.localName === 'trPr' &&
      child.namespaceURI === change.namespaceURI,
  );
  if (priors.length !== 1) return null;
  const prior = priors[0];
  if (!prior || hasRelationshipBindings(prior)) return null;
  const children = Array.from(prior.children);
  if (!children.length) return null;
  if (
    children.some(
      (child) =>
        child.namespaceURI !== change.namespaceURI ||
        child.children.length > 0 ||
        !SUPPORTED_PRIOR_CHILDREN.has(child.localName),
    )
  ) {
    return null;
  }
  const names = children.map((child) => child.localName);
  if (new Set(names).size !== names.length) return null;
  const snapshot: {
    cantSplit?: boolean;
    repeatHeader?: boolean;
    height?: DocumentRowFormattingHeight;
    hidden?: boolean;
    alignment?: DocumentTableAlignment;
    gridBefore?: number;
  } = {};
  for (const child of children) {
    if (child.localName === 'cantSplit') {
      snapshot.cantSplit = onOffValue(child);
      continue;
    }
    if (child.localName === 'tblHeader') {
      snapshot.repeatHeader = onOffValue(child);
      continue;
    }
    if (child.localName === 'trHeight') {
      const height = importedRowHeight(child);
      if (!height) return null;
      snapshot.height = height;
      continue;
    }
    if (child.localName === 'hidden') {
      snapshot.hidden = onOffValue(child);
      continue;
    }
    if (child.localName === 'jc') {
      const alignment = normalizeDocumentTableAlignment(
        mapJcValue(attribute(child, 'val')),
      );
      if (!alignment) return null;
      snapshot.alignment = alignment;
      continue;
    }
    if (child.localName === 'gridBefore') {
      const gridBefore = normalizeDocumentRowGridBefore(
        Number(attribute(child, 'val')),
      );
      if (gridBefore === null) return null;
      snapshot.gridBefore = gridBefore;
    }
  }
  return {
    id: `docx-row-format-change-${id}`,
    author,
    date: normalizeRevisionDate(rawDate),
    before: serializeDocumentRowFormatting(snapshot),
  };
}

function importedRowHeight(
  element: Element,
): DocumentRowFormattingHeight | null {
  const value = twipsToPixels(Number(attribute(element, 'val')));
  if (value === null) return null;
  const ruleAttribute = attribute(element, 'hRule');
  const rule =
    ruleAttribute === 'exact'
      ? 'exact'
      : ruleAttribute === null ||
          ruleAttribute === '' ||
          ruleAttribute === 'atLeast'
        ? 'atLeast'
        : null;
  if (!rule) return null;
  return { value, rule };
}

function twipsToPixels(value: number): number | null {
  if (!Number.isFinite(value) || value <= 0) return null;
  return Math.round(value * PIXELS_PER_TWIP * 100) / 100;
}

function onOffValue(element: Element): boolean {
  const value = attribute(element, 'val');
  if (value === null || value === '') return true;
  return !(
    value === '0' ||
    value === 'false' ||
    value === 'off' ||
    value === 'False'
  );
}

function mapJcValue(value: string | null): string | null {
  if (!value) return null;
  if (value === 'start') return 'left';
  if (value === 'end') return 'right';
  return value;
}

function wordAttribute(element: Element, localName: string): string | null {
  const namespace = element.namespaceURI;
  if (!namespace) return null;
  const matches = Array.from(element.attributes).filter(
    (candidate) =>
      xmlAttributeLocalName(candidate) === localName &&
      xmlAttributeNamespace(element, candidate) === namespace,
  );
  return matches.length === 1 ? (matches[0]?.value ?? null) : null;
}

function hasUnsupportedRevisionAttributes(element: Element): boolean {
  const namespace = element.namespaceURI;
  return Array.from(element.attributes).some(
    (candidate) =>
      xmlAttributeNamespace(element, candidate) === namespace &&
      !REVISION_ATTRIBUTES.has(xmlAttributeLocalName(candidate)),
  );
}

function hasRelationshipBindings(element: Element): boolean {
  for (const attributeNode of Array.from(element.attributes)) {
    const namespace =
      attributeNode.namespaceURI ||
      xmlAttributeNamespace(element, attributeNode) ||
      '';
    if (RELATIONSHIP_NAMESPACES.has(namespace)) return true;
  }
  for (const descendant of Array.from(element.querySelectorAll('*'))) {
    for (const attributeNode of Array.from(descendant.attributes)) {
      const namespace =
        attributeNode.namespaceURI ||
        xmlAttributeNamespace(descendant, attributeNode) ||
        '';
      if (RELATIONSHIP_NAMESPACES.has(namespace)) return true;
    }
  }
  return false;
}

function normalizeRevisionDate(value: string | null): string {
  if (!value) return '';
  const time = Date.parse(value);
  return Number.isFinite(time) ? new Date(time).toISOString() : '';
}
