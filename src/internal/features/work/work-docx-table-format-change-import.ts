import {
  normalizeDocumentTableAlignment,
  normalizeDocumentTableCellMarginOverrides,
  normalizeDocumentTableLayoutAlgorithm,
  type DocumentTableAlignment,
  type DocumentTableCellMarginOverrides,
  type DocumentTableCellMarginSide,
  type DocumentTableLayoutAlgorithm,
  type DocumentTablePreferredWidth,
} from './work-document-table-geometry';
import {
  normalizeDocumentTableFormattingSnapshot,
  serializeDocumentTableFormatting,
} from './work-document-table-format-changes';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { attribute, directChildren } from './work-ooxml-package';

const MAX_REVISION_DATE_LENGTH = 64;
const REVISION_ATTRIBUTES = new Set(['id', 'author', 'date']);
const SUPPORTED_PRIOR_CHILDREN = new Set([
  'jc',
  'tblW',
  'tblInd',
  'tblCellMar',
  'tblLayout',
  'bidiVisual',
]);
const MARGIN_SIDES = new Set(['top', 'right', 'bottom', 'left', 'start', 'end']);
const RELATIONSHIP_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  'http://purl.oclc.org/ooxml/officeDocument/relationships',
  'http://schemas.openxmlformats.org/package/2006/relationships',
]);
const PIXELS_PER_TWIP = 96 / 1440;

export interface SupportedDocxTableFormattingChange {
  id: string;
  author: string;
  date: string;
  before: string;
}

/**
 * Relationship-free `w:tblPrChange` whose prior snapshot contains only
 * `w:jc`, `w:tblW`, `w:tblInd`, `w:tblCellMar`, `w:tblLayout`, and/or
 * `w:bidiVisual`. Broader
 * property sets stay on the opaque OMML path.
 */
export function isSupportedDocxTableFormattingChange(
  change: Element,
): boolean {
  return supportedTableFormattingChange(change) !== null;
}

export function supportedDocxTableFormattingChangeFromProperties(
  tableProperties: Element | null | undefined,
): SupportedDocxTableFormattingChange | null {
  if (!tableProperties) return null;
  const changes = directChildren(tableProperties, 'tblPrChange').filter(
    (element) => element.namespaceURI === tableProperties.namespaceURI,
  );
  if (changes.length !== 1) return null;
  return supportedTableFormattingChange(changes[0] ?? null);
}

export function applyDocumentTableFormattingChangeToElement(
  table: HTMLTableElement,
  change: SupportedDocxTableFormattingChange,
): void {
  table.dataset.documentChange = 'true';
  table.dataset.changeKind = 'table-formatting';
  table.dataset.changeId = change.id;
  table.dataset.changeAuthor = change.author;
  table.dataset.changeDate = change.date;
  table.dataset.changeBefore = change.before;
}

function supportedTableFormattingChange(
  change: Element | null,
): SupportedDocxTableFormattingChange | null {
  if (
    !change ||
    change.localName !== 'tblPrChange' ||
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
      child.localName === 'tblPr' &&
      child.namespaceURI === change.namespaceURI,
  );
  if (priors.length !== 1) return null;
  const prior = priors[0];
  if (!prior || hasRelationshipBindings(prior)) return null;
  const children = Array.from(prior.children);
  if (!children.length) return null;
  if (
    children.some((child) => {
      if (
        child.namespaceURI !== change.namespaceURI ||
        !SUPPORTED_PRIOR_CHILDREN.has(child.localName)
      ) {
        return true;
      }
      if (child.localName === 'tblCellMar') return false;
      return child.children.length > 0;
    })
  ) {
    return null;
  }
  const localNames = children.map((child) => child.localName);
  if (new Set(localNames).size !== localNames.length) return null;

  let layout: DocumentTableLayoutAlgorithm | undefined;
  let alignment: DocumentTableAlignment | undefined;
  let width: DocumentTablePreferredWidth | undefined;
  let indent: number | undefined;
  let cellMargins: DocumentTableCellMarginOverrides | undefined;
  let bidiVisual: boolean | undefined;
  for (const child of children) {
    if (child.localName === 'tblLayout') {
      const value = normalizeDocumentTableLayoutAlgorithm(
        attribute(child, 'type'),
      );
      if (!value) return null;
      layout = value;
      continue;
    }
    if (child.localName === 'jc') {
      const value = normalizeDocumentTableAlignment(
        mapJcValue(attribute(child, 'val')),
      );
      if (!value) return null;
      alignment = value;
      continue;
    }
    if (child.localName === 'tblW') {
      const value = importedPreferredWidth(child);
      if (!value) return null;
      width = value;
      continue;
    }
    if (child.localName === 'tblInd') {
      const value = importedIndent(child);
      if (value === null) return null;
      indent = value;
      continue;
    }
    if (child.localName === 'tblCellMar') {
      const value = importedTableCellMargins(child);
      if (!value) return null;
      cellMargins = value;
      continue;
    }
    if (child.localName === 'bidiVisual') {
      bidiVisual = onOffValue(child);
    }
  }
  const snapshot = normalizeDocumentTableFormattingSnapshot({
    ...(layout ? { layout } : {}),
    ...(alignment ? { alignment } : {}),
    ...(width ? { width } : {}),
    ...(indent !== undefined ? { indent } : {}),
    ...(cellMargins ? { cellMargins } : {}),
    ...(bidiVisual !== undefined ? { bidiVisual } : {}),
  });
  if (!snapshot) return null;
  const date = normalizeRevisionDate(rawDate);
  return {
    id: `docx-table-format-change-${id}`,
    author,
    date,
    before: serializeDocumentTableFormatting(snapshot),
  };
}

function importedTableCellMargins(
  element: Element,
): DocumentTableCellMarginOverrides | null {
  const children = Array.from(element.children);
  if (!children.length) return null;
  if (
    children.some(
      (child) =>
        child.namespaceURI !== element.namespaceURI ||
        !MARGIN_SIDES.has(child.localName) ||
        child.children.length > 0,
    )
  ) {
    return null;
  }
  const names = children.map((child) => child.localName);
  if (new Set(names).size !== names.length) return null;
  const margins: DocumentTableCellMarginOverrides = {};
  for (const child of children) {
    const side = marginSide(child.localName);
    if (!side) return null;
    const type = attribute(child, 'type');
    if (type !== null && type !== 'dxa') return null;
    const pixels = twipsToPixels(Number(attribute(child, 'w')), true);
    if (pixels === null) return null;
    if (margins[side] !== undefined) return null;
    margins[side] = pixels;
  }
  return normalizeDocumentTableCellMarginOverrides(margins);
}

function marginSide(localName: string): DocumentTableCellMarginSide | null {
  if (localName === 'top') return 'top';
  if (localName === 'bottom') return 'bottom';
  if (localName === 'left' || localName === 'start') return 'left';
  if (localName === 'right' || localName === 'end') return 'right';
  return null;
}

function importedPreferredWidth(
  width: Element,
): DocumentTablePreferredWidth | null {
  const type = attribute(width, 'type');
  if (type === 'auto' || type === 'nil') return { type: 'auto', value: null };
  if (type === 'pct') {
    const value = percentageValue(attribute(width, 'w'));
    return value === null ? null : { type: 'percent', value };
  }
  if (type === 'dxa') {
    const value = twipsToPixels(Number(attribute(width, 'w')), false);
    return value === null ? null : { type: 'pixels', value };
  }
  return null;
}

function importedIndent(indent: Element): number | null {
  const type = attribute(indent, 'type');
  if (type !== null && type !== 'dxa') return null;
  return twipsToPixels(Number(attribute(indent, 'w')), true);
}

function percentageValue(value: string | null): number | null {
  const normalized = value?.trim();
  if (!normalized) return null;
  const percentage = normalized.endsWith('%')
    ? Number(normalized.slice(0, -1))
    : Number(normalized) / 50;
  if (!Number.isFinite(percentage) || percentage <= 0) return null;
  return Math.round(percentage * 100) / 100;
}

function twipsToPixels(value: number, allowZero: boolean): number | null {
  if (!Number.isFinite(value) || (!allowZero && value <= 0) || value < 0) {
    return null;
  }
  return Math.round(value * PIXELS_PER_TWIP * 100) / 100;
}

function mapJcValue(value: string | null): string | null {
  if (!value) return null;
  if (value === 'start') return 'left';
  if (value === 'end') return 'right';
  return value;
}

function onOffValue(element: Element): boolean {
  const value = attribute(element, 'val');
  if (value === null || value === '') return true;
  const normalized = value.trim().toLowerCase();
  return (
    normalized === '1' ||
    normalized === 'true' ||
    normalized === 'on' ||
    normalized === 'yes'
  );
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
