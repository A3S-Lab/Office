import {
  cloneXmlElement,
  declareInheritedNamespaces,
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { attribute, directChild, parseXml } from './work-ooxml-package';

export const DOCUMENT_TABLE_FLOAT_OMML_ATTRIBUTE =
  'data-office-table-float-omml';
export const MAX_DOCUMENT_TABLE_FLOAT_OMML_LENGTH = 65_536;

const DOCX_RELATIONSHIP_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  'http://purl.oclc.org/ooxml/officeDocument/relationships',
  'http://schemas.openxmlformats.org/package/2006/relationships',
]);
const DOCX_WORDPROCESSING_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  'http://purl.oclc.org/ooxml/wordprocessingml/main',
]);

const WORD_NAMESPACE =
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main';
const MAX_BOUNDED_SIGNED_TWIPS = 31_680;
const MAX_BOUNDED_UNSIGNED_TWIPS = 31_680;

const REVIEWABLE_TBLP_PR_ATTRIBUTES = new Set([
  'horzAnchor',
  'vertAnchor',
  'tblpX',
  'tblpY',
  'tblpXSpec',
  'tblpYSpec',
  'leftFromText',
  'rightFromText',
  'topFromText',
  'bottomFromText',
]);

const HORZ_ANCHORS = new Set(['text', 'margin', 'page'] as const);
const VERT_ANCHORS = new Set(['text', 'margin', 'page'] as const);
const X_SPECS = new Set([
  'left',
  'center',
  'right',
  'inside',
  'outside',
] as const);
const Y_SPECS = new Set([
  'inline',
  'top',
  'center',
  'bottom',
  'inside',
  'outside',
] as const);

export type DocumentTableFloatHorzAnchor = 'text' | 'margin' | 'page';
export type DocumentTableFloatVertAnchor = 'text' | 'margin' | 'page';
export type DocumentTableFloatXSpec =
  | 'left'
  | 'center'
  | 'right'
  | 'inside'
  | 'outside';
export type DocumentTableFloatYSpec =
  | 'inline'
  | 'top'
  | 'center'
  | 'bottom'
  | 'inside'
  | 'outside';

/**
 * Relationship-free, attribute-only floating-table position admitted as a
 * reviewable `w:tblPrChange` prior (`w:tblpPr`).
 */
export interface DocumentTableFloatPosition {
  horzAnchor: DocumentTableFloatHorzAnchor;
  vertAnchor: DocumentTableFloatVertAnchor;
  tblpX?: number;
  tblpY?: number;
  tblpXSpec?: DocumentTableFloatXSpec;
  tblpYSpec?: DocumentTableFloatYSpec;
  leftFromText?: number;
  rightFromText?: number;
  topFromText?: number;
  bottomFromText?: number;
}

export function encodeDocumentTableFloatOmml(omml: string): string {
  if (!omml || omml.length > MAX_DOCUMENT_TABLE_FLOAT_OMML_LENGTH) return '';
  const bytes = new TextEncoder().encode(omml);
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
}

export function decodeDocumentTableFloatOmml(encoded: string): string | null {
  if (!encoded) return null;
  try {
    const binary = atob(encoded);
    const bytes = Uint8Array.from(binary, (character) =>
      character.charCodeAt(0),
    );
    const omml = new TextDecoder().decode(bytes).trim();
    if (!omml || omml.length > MAX_DOCUMENT_TABLE_FLOAT_OMML_LENGTH) {
      return null;
    }
    return omml;
  } catch {
    return null;
  }
}

export function documentTableFloatOmmlFromElement(
  element: Element,
): string | null {
  if (!(element instanceof HTMLElement)) return null;
  return decodeDocumentTableFloatOmml(
    element.dataset.officeTableFloatOmml ?? '',
  );
}

export function applyDocumentTableFloatOmmlToElement(
  element: HTMLElement,
  omml: string | null | undefined,
): void {
  const encoded = omml ? encodeDocumentTableFloatOmml(omml) : '';
  if (encoded) element.dataset.officeTableFloatOmml = encoded;
  else delete element.dataset.officeTableFloatOmml;
}

export function docxTableFloatHasRelationshipBindings(
  element: Element,
): boolean {
  for (const attribute of Array.from(element.attributes)) {
    const namespace =
      attribute.namespaceURI || xmlAttributeNamespace(element, attribute) || '';
    if (DOCX_RELATIONSHIP_NAMESPACES.has(namespace)) return true;
  }
  return false;
}

export function serializePreservableDocxTableFloat(
  tableProperties: Element | null | undefined,
): string | null {
  const float = tableProperties ? directChild(tableProperties, 'tblpPr') : null;
  if (!float) return null;
  const namespace = float.namespaceURI ?? '';
  if (!DOCX_WORDPROCESSING_NAMESPACES.has(namespace)) return null;
  if (docxTableFloatHasRelationshipBindings(float)) return null;
  if (float.children.length > 0) return null;
  const prefix = float.prefix || 'w';
  const shell = parseXml(
    `<${prefix}:${float.localName} xmlns:${prefix}="${namespace}"/>`,
    'opaque table float shell',
  );
  const clone = cloneXmlElement(shell, float);
  declareInheritedNamespaces(clone, float);
  shell.replaceChild(clone, shell.documentElement);
  const omml = new XMLSerializer().serializeToString(clone).trim();
  if (!omml || omml.length > MAX_DOCUMENT_TABLE_FLOAT_OMML_LENGTH) return null;
  return omml;
}

export function importDocxTableFloatElement(
  document: Document,
  omml: string,
): Element | null {
  try {
    const parsed = parseXml(omml, 'preserved table float');
    const root = parsed.documentElement;
    if (
      root.localName !== 'tblpPr' ||
      !DOCX_WORDPROCESSING_NAMESPACES.has(root.namespaceURI ?? '')
    ) {
      return null;
    }
    if (docxTableFloatHasRelationshipBindings(root)) return null;
    const imported = cloneXmlElement(document, root);
    declareInheritedNamespaces(imported, root);
    return imported;
  } catch {
    return null;
  }
}

export function normalizeDocumentTableFloatPosition(
  value: unknown,
): DocumentTableFloatPosition | null {
  if (!value || typeof value !== 'object' || Array.isArray(value)) return null;
  const record = value as Record<string, unknown>;
  const keys = Object.keys(record);
  if (
    !keys.length ||
    keys.some((key) => !REVIEWABLE_TBLP_PR_ATTRIBUTES.has(key))
  ) {
    return null;
  }
  const horzAnchor = normalizeHorzAnchor(record.horzAnchor);
  const vertAnchor = normalizeVertAnchor(record.vertAnchor);
  if (!horzAnchor || !vertAnchor) return null;

  const float: DocumentTableFloatPosition = { horzAnchor, vertAnchor };

  if ('tblpX' in record && record.tblpX !== undefined) {
    const tblpX = normalizeSignedTwips(record.tblpX);
    if (tblpX === null) return null;
    float.tblpX = tblpX;
  }
  if ('tblpY' in record && record.tblpY !== undefined) {
    const tblpY = normalizeSignedTwips(record.tblpY);
    if (tblpY === null) return null;
    float.tblpY = tblpY;
  }
  if ('tblpXSpec' in record && record.tblpXSpec !== undefined) {
    const tblpXSpec = normalizeXSpec(record.tblpXSpec);
    if (!tblpXSpec) return null;
    float.tblpXSpec = tblpXSpec;
  }
  if ('tblpYSpec' in record && record.tblpYSpec !== undefined) {
    const tblpYSpec = normalizeYSpec(record.tblpYSpec);
    if (!tblpYSpec) return null;
    float.tblpYSpec = tblpYSpec;
  }
  if (
    float.tblpX === undefined &&
    float.tblpXSpec === undefined &&
    float.tblpY === undefined &&
    float.tblpYSpec === undefined
  ) {
    return null;
  }
  if (float.tblpX === undefined && float.tblpXSpec === undefined) return null;
  if (float.tblpY === undefined && float.tblpYSpec === undefined) return null;

  for (const side of [
    'leftFromText',
    'rightFromText',
    'topFromText',
    'bottomFromText',
  ] as const) {
    if (!(side in record) || record[side] === undefined) continue;
    const distance = normalizeUnsignedTwips(record[side]);
    if (distance === null) return null;
    float[side] = distance;
  }
  return orderedDocumentTableFloatPosition(float);
}

export function orderedDocumentTableFloatPosition(
  float: DocumentTableFloatPosition,
): DocumentTableFloatPosition {
  const ordered: DocumentTableFloatPosition = {
    horzAnchor: float.horzAnchor,
    vertAnchor: float.vertAnchor,
  };
  if (float.tblpX !== undefined) ordered.tblpX = float.tblpX;
  if (float.tblpY !== undefined) ordered.tblpY = float.tblpY;
  if (float.tblpXSpec) ordered.tblpXSpec = float.tblpXSpec;
  if (float.tblpYSpec) ordered.tblpYSpec = float.tblpYSpec;
  if (float.leftFromText !== undefined)
    ordered.leftFromText = float.leftFromText;
  if (float.rightFromText !== undefined)
    ordered.rightFromText = float.rightFromText;
  if (float.topFromText !== undefined) ordered.topFromText = float.topFromText;
  if (float.bottomFromText !== undefined)
    ordered.bottomFromText = float.bottomFromText;
  return ordered;
}

export function parseReviewableDocxTableFloatElement(
  element: Element,
): DocumentTableFloatPosition | null {
  if (
    element.localName !== 'tblpPr' ||
    !DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '')
  ) {
    return null;
  }
  if (element.children.length > 0) return null;
  if (docxTableFloatHasRelationshipBindings(element)) return null;
  const namespace = element.namespaceURI;
  if (!namespace) return null;
  const attributes = Array.from(element.attributes).filter(
    (candidate) => xmlAttributeNamespace(element, candidate) === namespace,
  );
  if (!attributes.length) return null;
  if (
    attributes.some(
      (candidate) =>
        !REVIEWABLE_TBLP_PR_ATTRIBUTES.has(xmlAttributeLocalName(candidate)),
    )
  ) {
    return null;
  }
  return normalizeDocumentTableFloatPosition({
    horzAnchor: attribute(element, 'horzAnchor'),
    vertAnchor: attribute(element, 'vertAnchor'),
    ...(attribute(element, 'tblpX') !== null
      ? { tblpX: attribute(element, 'tblpX') }
      : {}),
    ...(attribute(element, 'tblpY') !== null
      ? { tblpY: attribute(element, 'tblpY') }
      : {}),
    ...(attribute(element, 'tblpXSpec') !== null
      ? { tblpXSpec: attribute(element, 'tblpXSpec') }
      : {}),
    ...(attribute(element, 'tblpYSpec') !== null
      ? { tblpYSpec: attribute(element, 'tblpYSpec') }
      : {}),
    ...(attribute(element, 'leftFromText') !== null
      ? { leftFromText: attribute(element, 'leftFromText') }
      : {}),
    ...(attribute(element, 'rightFromText') !== null
      ? { rightFromText: attribute(element, 'rightFromText') }
      : {}),
    ...(attribute(element, 'topFromText') !== null
      ? { topFromText: attribute(element, 'topFromText') }
      : {}),
    ...(attribute(element, 'bottomFromText') !== null
      ? { bottomFromText: attribute(element, 'bottomFromText') }
      : {}),
  });
}

export function parseReviewableDocumentTableFloatFromOmml(
  omml: string | null | undefined,
): DocumentTableFloatPosition | null {
  if (typeof omml !== 'string' || !omml.trim()) return null;
  try {
    const parsed = parseXml(omml, 'reviewable table float');
    return parseReviewableDocxTableFloatElement(parsed.documentElement);
  } catch {
    return null;
  }
}

export function serializeDocumentTableFloatOmmlFromPosition(
  float: DocumentTableFloatPosition,
): string | null {
  const normalized = normalizeDocumentTableFloatPosition(float);
  if (!normalized) return null;
  const document = parseXml(
    `<w:tblpPr xmlns:w="${WORD_NAMESPACE}"/>`,
    'reviewable table float shell',
  );
  const element = document.documentElement;
  applyDocumentTableFloatPositionToElement(element, normalized);
  const omml = new XMLSerializer().serializeToString(element).trim();
  if (!omml || omml.length > MAX_DOCUMENT_TABLE_FLOAT_OMML_LENGTH) return null;
  return omml;
}

export function applyDocumentTableFloatPositionToElement(
  element: Element,
  float: DocumentTableFloatPosition,
): void {
  const namespace = element.namespaceURI || WORD_NAMESPACE;
  const prefix = element.prefix || 'w';
  const set = (localName: string, value: string) => {
    element.setAttributeNS(namespace, `${prefix}:${localName}`, value);
  };
  set('horzAnchor', float.horzAnchor);
  set('vertAnchor', float.vertAnchor);
  if (float.tblpX !== undefined) set('tblpX', String(float.tblpX));
  if (float.tblpY !== undefined) set('tblpY', String(float.tblpY));
  if (float.tblpXSpec) set('tblpXSpec', float.tblpXSpec);
  if (float.tblpYSpec) set('tblpYSpec', float.tblpYSpec);
  if (float.leftFromText !== undefined)
    set('leftFromText', String(float.leftFromText));
  if (float.rightFromText !== undefined)
    set('rightFromText', String(float.rightFromText));
  if (float.topFromText !== undefined)
    set('topFromText', String(float.topFromText));
  if (float.bottomFromText !== undefined)
    set('bottomFromText', String(float.bottomFromText));
}

function normalizeHorzAnchor(
  value: unknown,
): DocumentTableFloatHorzAnchor | null {
  return typeof value === 'string' &&
    (HORZ_ANCHORS as Set<string>).has(value.trim())
    ? (value.trim() as DocumentTableFloatHorzAnchor)
    : null;
}

function normalizeVertAnchor(
  value: unknown,
): DocumentTableFloatVertAnchor | null {
  return typeof value === 'string' &&
    (VERT_ANCHORS as Set<string>).has(value.trim())
    ? (value.trim() as DocumentTableFloatVertAnchor)
    : null;
}

function normalizeXSpec(value: unknown): DocumentTableFloatXSpec | null {
  return typeof value === 'string' && (X_SPECS as Set<string>).has(value.trim())
    ? (value.trim() as DocumentTableFloatXSpec)
    : null;
}

function normalizeYSpec(value: unknown): DocumentTableFloatYSpec | null {
  return typeof value === 'string' && (Y_SPECS as Set<string>).has(value.trim())
    ? (value.trim() as DocumentTableFloatYSpec)
    : null;
}

function normalizeSignedTwips(value: unknown): number | null {
  const numeric =
    typeof value === 'number'
      ? value
      : typeof value === 'string' && value.trim()
        ? Number(value.trim())
        : Number.NaN;
  if (
    !Number.isInteger(numeric) ||
    numeric < -MAX_BOUNDED_SIGNED_TWIPS ||
    numeric > MAX_BOUNDED_SIGNED_TWIPS
  ) {
    return null;
  }
  return numeric;
}

function normalizeUnsignedTwips(value: unknown): number | null {
  const numeric =
    typeof value === 'number'
      ? value
      : typeof value === 'string' && value.trim()
        ? Number(value.trim())
        : Number.NaN;
  if (
    !Number.isInteger(numeric) ||
    numeric < 0 ||
    numeric > MAX_BOUNDED_UNSIGNED_TWIPS
  ) {
    return null;
  }
  return numeric;
}
