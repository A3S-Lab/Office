import { serializeDocumentSectionFormatting } from './work-document-section-format-changes';
import {
  type WorkDocumentPageMarginKey,
  type WorkDocumentPageMargins,
  normalizeDocumentPageMargins,
} from './work-document-page-margins';
import {
  MAX_WORD_PAGE_TWIPS,
  MAX_WORD_PAPER_CODE,
  MAX_WORD_PAPER_SOURCE_CODE,
  normalizeDocumentPageGeometry,
  normalizeDocumentPaperSource,
  type WorkDocumentPageGeometry,
  type WorkDocumentPaperSource,
} from './work-document-page-size';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import {
  parseBoundedDocxInteger,
  parseDocxTwipsMeasure,
} from './work-docx-twips';
import { directChildren } from './work-ooxml-package';

const MAX_REVISION_DATE_LENGTH = 64;
const REVISION_ATTRIBUTES = new Set(['id', 'author', 'date']);
const SUPPORTED_PRIOR_CHILDREN = new Set([
  'pgSz',
  'pgMar',
  'paperSrc',
  'cols',
  'titlePg',
  'rtlGutter',
]);
const PAGE_SIZE_ATTRIBUTE_SET = new Set(['w', 'h', 'orient', 'code']);
const PAPER_SOURCE_ATTRIBUTE_SET = new Set(['first', 'other']);
const COLUMNS_ATTRIBUTE_SET = new Set(['num', 'space', 'sep', 'equalWidth']);
const TWIPS_PER_MILLIMETER = 1440 / 25.4;
const PAGE_MARGIN_KEYS = [
  'top',
  'right',
  'bottom',
  'left',
  'header',
  'footer',
  'gutter',
] as const satisfies readonly WorkDocumentPageMarginKey[];
const RELATIONSHIP_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  'http://purl.oclc.org/ooxml/officeDocument/relationships',
  'http://schemas.openxmlformats.org/package/2006/relationships',
]);
const MAX_UNSIGNED_WORD_TWIPS = 31_680;
const MIN_SIGNED_WORD_TWIPS = -2_147_483_648;
const MAX_SIGNED_WORD_TWIPS = 2_147_483_647;

export interface SupportedDocxSectionFormattingChange {
  id: string;
  author: string;
  date: string;
  before: string;
}

/**
 * Relationship-free `w:sectPrChange` whose prior snapshot contains only
 * orientation-only or complete `w:pgSz` (w/h with optional orient/code), a
 * complete seven-edge `w:pgMar`, `w:paperSrc`, equal-width `w:cols`,
 * and/or `w:titlePg` and/or `w:rtlGutter`. Broader section property sets stay on the opaque OMML
 * path.
 */
export function isSupportedDocxSectionFormattingChange(
  change: Element,
): boolean {
  return supportedSectionFormattingChange(change) !== null;
}

export function supportedDocxSectionFormattingChangeFromProperties(
  sectionProperties: Element | null | undefined,
): SupportedDocxSectionFormattingChange | null {
  if (!sectionProperties) return null;
  const changes = directChildren(sectionProperties, 'sectPrChange').filter(
    (element) => element.namespaceURI === sectionProperties.namespaceURI,
  );
  if (changes.length !== 1) return null;
  return supportedSectionFormattingChange(changes[0] ?? null);
}

function supportedSectionFormattingChange(
  change: Element | null,
): SupportedDocxSectionFormattingChange | null {
  if (
    !change ||
    change.localName !== 'sectPrChange' ||
    !DOCX_WORDPROCESSING_NAMESPACES.has(change.namespaceURI ?? '')
  ) {
    return null;
  }
  if (
    hasRelationshipBindings(change) ||
    hasUnsupportedRevisionAttributes(change)
  ) {
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
      child.localName === 'sectPr' &&
      child.namespaceURI === change.namespaceURI,
  );
  if (priors.length !== 1) return null;
  const prior = priors[0];
  if (!prior || hasRelationshipBindings(prior)) return null;
  const children = Array.from(prior.children);
  if (!children.length || children.length > 6) return null;
  if (
    children.some(
      (child) =>
        child.namespaceURI !== change.namespaceURI ||
        !SUPPORTED_PRIOR_CHILDREN.has(child.localName) ||
        child.children.length > 0,
    )
  ) {
    return null;
  }
  const localNames = children.map((child) => child.localName);
  if (new Set(localNames).size !== localNames.length) return null;

  let orientation: 'portrait' | 'landscape' | undefined;
  let pageGeometry: WorkDocumentPageGeometry | undefined;
  let pageMargins: WorkDocumentPageMargins | undefined;
  let paperSource: WorkDocumentPaperSource | undefined;
  let columns:
    | {
        count: number;
        spacing: number;
        separator: boolean;
      }
    | undefined;
  let differentFirstPage: boolean | undefined;
  let rtlGutter: boolean | undefined;
  for (const child of children) {
    if (child.localName === 'pgSz') {
      const value = importedPageSize(child);
      if (!value) return null;
      if (value.orientation) orientation = value.orientation;
      if (value.pageGeometry) pageGeometry = value.pageGeometry;
      continue;
    }
    if (child.localName === 'pgMar') {
      const value = importedPageMargins(child);
      if (!value) return null;
      pageMargins = value;
      continue;
    }
    if (child.localName === 'paperSrc') {
      const value = importedPaperSource(child);
      if (!value) return null;
      paperSource = value;
      continue;
    }
    if (child.localName === 'cols') {
      const value = importedEqualColumns(child);
      if (!value) return null;
      columns = value;
      continue;
    }
    if (child.localName === 'titlePg') {
      differentFirstPage = onOffValue(child);
      continue;
    }
    if (child.localName === 'rtlGutter') {
      rtlGutter = onOffValue(child);
    }
  }
  const before = serializeDocumentSectionFormatting({
    ...(orientation ? { orientation } : {}),
    ...(pageGeometry ? { pageGeometry } : {}),
    ...(pageMargins ? { pageMargins } : {}),
    ...(paperSource ? { paperSource } : {}),
    ...(columns ? { columns } : {}),
    ...(differentFirstPage !== undefined ? { differentFirstPage } : {}),
    ...(rtlGutter !== undefined ? { rtlGutter } : {}),
  });
  return {
    id: `docx-section-format-change-${id}`,
    author,
    date: normalizeRevisionDate(rawDate),
    before,
  };
}

function importedPageSize(element: Element): {
  orientation?: 'portrait' | 'landscape';
  pageGeometry?: WorkDocumentPageGeometry;
} | null {
  const attributes = Array.from(element.attributes).filter(
    (candidate) =>
      xmlAttributeNamespace(element, candidate) === element.namespaceURI,
  );
  if (!attributes.length) return null;
  const names = new Set(
    attributes.map((candidate) => xmlAttributeLocalName(candidate)),
  );
  if ([...names].some((name) => !PAGE_SIZE_ATTRIBUTE_SET.has(name))) {
    return null;
  }
  if (names.size !== attributes.length) return null;

  if (names.size === 1 && names.has('orient')) {
    const value = attributes[0]?.value.trim();
    if (value === 'landscape') return { orientation: 'landscape' };
    if (value === 'portrait') return { orientation: 'portrait' };
    return null;
  }

  if (!names.has('w') || !names.has('h')) return null;
  const byName = new Map(
    attributes.map((candidate) => [
      xmlAttributeLocalName(candidate),
      candidate.value.trim(),
    ]),
  );
  const width = parseDocxTwipsMeasure(byName.get('w') ?? '', {
    minimum: 1,
    maximum: MAX_WORD_PAGE_TWIPS,
    signed: false,
    strict: false,
  });
  const height = parseDocxTwipsMeasure(byName.get('h') ?? '', {
    minimum: 1,
    maximum: MAX_WORD_PAGE_TWIPS,
    signed: false,
    strict: false,
  });
  const orientation = byName.get('orient');
  const codeSource = byName.get('code');
  const code =
    codeSource === undefined
      ? undefined
      : parseBoundedDocxInteger(codeSource, {
          minimum: 0,
          maximum: MAX_WORD_PAPER_CODE,
        });
  if (
    width === null ||
    height === null ||
    (orientation !== undefined &&
      orientation !== 'portrait' &&
      orientation !== 'landscape') ||
    (codeSource !== undefined && code === null)
  ) {
    return null;
  }
  const pageGeometry = normalizeDocumentPageGeometry({
    width,
    height,
    ...(orientation !== undefined ? { orientation } : {}),
    ...(code !== undefined ? { code } : {}),
  });
  return pageGeometry ? { pageGeometry } : null;
}

function importedEqualColumns(element: Element): {
  count: number;
  spacing: number;
  separator: boolean;
} | null {
  if (Array.from(element.children).length > 0) return null;
  const attributes = Array.from(element.attributes).filter(
    (candidate) =>
      xmlAttributeNamespace(element, candidate) === element.namespaceURI,
  );
  const names = new Set(
    attributes.map((candidate) => xmlAttributeLocalName(candidate)),
  );
  if ([...names].some((name) => !COLUMNS_ATTRIBUTE_SET.has(name))) {
    return null;
  }
  if (names.size !== attributes.length) return null;
  const byName = new Map(
    attributes.map((candidate) => [
      xmlAttributeLocalName(candidate),
      candidate.value.trim(),
    ]),
  );
  if (byName.has('equalWidth')) {
    const equalWidth = byName.get('equalWidth')?.toLowerCase();
    if (equalWidth !== '1' && equalWidth !== 'true' && equalWidth !== 'on') {
      return null;
    }
  }
  const countSource = byName.get('num');
  const count =
    countSource === undefined
      ? 1
      : parseBoundedDocxInteger(countSource, { minimum: 1, maximum: 6 });
  if (count === null) return null;
  const spaceSource = byName.get('space');
  let spacing = 12;
  if (spaceSource !== undefined) {
    const twips = parseDocxTwipsMeasure(spaceSource, {
      minimum: 0,
      maximum: MAX_UNSIGNED_WORD_TWIPS,
      signed: false,
      strict: false,
    });
    if (twips === null) return null;
    spacing = Math.round((twips / TWIPS_PER_MILLIMETER) * 10) / 10;
  }
  const sepSource = byName.get('sep')?.toLowerCase();
  const separator =
    sepSource === '1' || sepSource === 'true' || sepSource === 'on';
  return { count, spacing, separator };
}

function importedPaperSource(element: Element): WorkDocumentPaperSource | null {
  const attributes = Array.from(element.attributes).filter(
    (candidate) =>
      xmlAttributeNamespace(element, candidate) === element.namespaceURI,
  );
  if (!attributes.length) return null;
  const names = new Set(
    attributes.map((candidate) => xmlAttributeLocalName(candidate)),
  );
  if ([...names].some((name) => !PAPER_SOURCE_ATTRIBUTE_SET.has(name))) {
    return null;
  }
  if (names.size !== attributes.length) return null;
  const source: WorkDocumentPaperSource = {};
  for (const attribute of attributes) {
    const name = xmlAttributeLocalName(attribute);
    if (name !== 'first' && name !== 'other') return null;
    const parsed = parseBoundedDocxInteger(attribute.value.trim(), {
      minimum: 0,
      maximum: MAX_WORD_PAPER_SOURCE_CODE,
    });
    if (parsed === null) return null;
    source[name] = parsed;
  }
  const normalized = normalizeDocumentPaperSource(source);
  if (
    !normalized ||
    (normalized.first === undefined && normalized.other === undefined)
  ) {
    return null;
  }
  return normalized;
}

function importedPageMargins(element: Element): WorkDocumentPageMargins | null {
  const values: Partial<Record<WorkDocumentPageMarginKey, number>> = {};
  const seen = new Set<string>();
  for (const item of Array.from(element.attributes)) {
    const namespace =
      item.namespaceURI || xmlAttributeNamespace(element, item) || '';
    if (RELATIONSHIP_NAMESPACES.has(namespace)) return null;
    if (namespace && namespace !== element.namespaceURI) return null;
    const name = xmlAttributeLocalName(item);
    if (!(PAGE_MARGIN_KEYS as readonly string[]).includes(name)) return null;
    if (seen.has(name)) return null;
    seen.add(name);
    const signed = name === 'top' || name === 'bottom';
    const value = parseDocxTwipsMeasure(item.value.trim(), {
      minimum: signed ? MIN_SIGNED_WORD_TWIPS : 0,
      maximum: signed ? MAX_SIGNED_WORD_TWIPS : MAX_UNSIGNED_WORD_TWIPS,
      signed,
      strict: false,
    });
    if (value === null) return null;
    values[name as WorkDocumentPageMarginKey] = value;
  }
  if (seen.size !== PAGE_MARGIN_KEYS.length) return null;
  return normalizeDocumentPageMargins(values);
}

function onOffValue(element: Element): boolean {
  const matches = Array.from(element.attributes).filter(
    (candidate) =>
      xmlAttributeLocalName(candidate) === 'val' &&
      xmlAttributeNamespace(element, candidate) === element.namespaceURI,
  );
  if (matches.length > 1) return true;
  const value = matches[0]?.value ?? null;
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
