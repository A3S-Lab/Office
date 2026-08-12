import type JSZip from 'jszip';
import { normalizeDocumentHref } from './work-document-links';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import {
  XMLNS_NAMESPACE,
  xmlAttributeLocalName,
  xmlAttributeNamespace,
  xmlDeclaredPrefix,
  xmlNamespaceUri,
} from './work-docx-settings-xml';
import { directChildren, parseXml } from './work-ooxml-package';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';

const PACKAGE_RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/relationships';
const TRANSITIONAL_RELATIONSHIP_NAMESPACE =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const STRICT_RELATIONSHIP_NAMESPACE =
  'http://purl.oclc.org/ooxml/officeDocument/relationships';
const RELATIONSHIP_NAMESPACES = new Set([
  TRANSITIONAL_RELATIONSHIP_NAMESPACE,
  STRICT_RELATIONSHIP_NAMESPACE,
]);
const HYPERLINK_RELATIONSHIP_TYPES = new Set([
  `${TRANSITIONAL_RELATIONSHIP_NAMESPACE}/hyperlink`,
  `${STRICT_RELATIONSHIP_NAMESPACE}/hyperlink`,
]);
const GENERATED_HYPERLINK_RELATIONSHIP_TYPE = `${TRANSITIONAL_RELATIONSHIP_NAMESPACE}/hyperlink`;
const RELATIONSHIP_ID_PATTERN = /^[A-Za-z_][A-Za-z0-9_.-]{0,254}$/;
const MAX_RELATIONSHIPS = 65_536;

export interface DocxHyperlinkDestination {
  kind: 'external' | 'internal';
  relationshipId?: string;
  target: string;
}

interface RelationshipRecord {
  id: string;
  target: string;
  targetMode: string;
  type: string;
}

interface RelationshipDocument {
  byId: Map<string, RelationshipRecord>;
  document: Document;
  usedIds: Set<string>;
}

export interface DocxHyperlinkRelationshipState {
  archive: JSZip;
  dirty: boolean;
  generated: RelationshipDocument | null;
  generatedMalformed: boolean;
  path: string;
  source: RelationshipDocument | null;
}

export interface DocxExternalHyperlinkRequest {
  sourceId: string;
  target: string;
}

export async function loadDocxHyperlinkRelationshipState(
  generatedArchive: JSZip,
  sourceArchive: JSZip,
  ownerPart: string,
): Promise<DocxHyperlinkRelationshipState> {
  const path = relationshipPartPath(ownerPart);
  const [generated, source] = await Promise.all([
    loadRelationships(generatedArchive, path, 'generated'),
    loadRelationships(sourceArchive, path, 'source'),
  ]);
  return {
    archive: generatedArchive,
    dirty: false,
    generated: generated.document,
    generatedMalformed: generated.malformed,
    path,
    source: source.document,
  };
}

export function readDocxHyperlinkDestination(
  hyperlink: Element,
  state: DocxHyperlinkRelationshipState,
  role: 'generated' | 'source',
): DocxHyperlinkDestination | null {
  if (!DOCX_WORDPROCESSING_NAMESPACES.has(hyperlink.namespaceURI ?? '')) {
    return null;
  }
  const anchor = uniqueNamespacedAttribute(
    hyperlink,
    'anchor',
    DOCX_WORDPROCESSING_NAMESPACES,
  );
  const relationshipId = uniqueNamespacedAttribute(
    hyperlink,
    'id',
    RELATIONSHIP_NAMESPACES,
  );
  if (anchor === undefined || relationshipId === undefined) return null;
  const normalizedAnchor = anchor === null ? null : normalizeAnchor(anchor);
  if (anchor !== null && !normalizedAnchor) return null;
  if (normalizedAnchor && relationshipId) return null;
  if (normalizedAnchor) return { kind: 'internal', target: normalizedAnchor };
  if (!relationshipId || !RELATIONSHIP_ID_PATTERN.test(relationshipId)) {
    return null;
  }
  const relationships = role === 'generated' ? state.generated : state.source;
  const relationship = relationships?.byId.get(relationshipId);
  if (!relationship || !isExternalHyperlinkRelationship(relationship)) {
    return null;
  }
  const target = normalizeDocumentHref(relationship.target);
  if (!target || target.startsWith('#')) return null;
  return {
    kind: 'external',
    relationshipId,
    target,
  };
}

export function ensureDocxExternalHyperlinkRelationships(
  state: DocxHyperlinkRelationshipState,
  requests: readonly DocxExternalHyperlinkRequest[],
): string[] | null {
  if (!requests.length) return [];
  if (state.generatedMalformed) return null;
  const relationships = state.generated ?? createRelationshipsDocument();
  const usedIds = new Set(relationships.usedIds);
  const idByTarget = new Map(
    Array.from(relationships.byId.values()).flatMap((item) => {
      const target = isExternalHyperlinkRelationship(item)
        ? normalizeDocumentHref(item.target)
        : null;
      return target && !target.startsWith('#') ? [[target, item.id]] : [];
    }),
  );
  const pending: Array<{ id: string; target: string }> = [];
  const result: string[] = [];
  for (const request of requests) {
    if (!RELATIONSHIP_ID_PATTERN.test(request.sourceId)) return null;
    const target = normalizeDocumentHref(request.target);
    if (!target || target.startsWith('#')) return null;
    const existingId = idByTarget.get(target);
    if (existingId) {
      result.push(existingId);
      continue;
    }
    if (usedIds.size >= MAX_RELATIONSHIPS) return null;
    const id = usedIds.has(request.sourceId)
      ? nextRelationshipId(usedIds)
      : request.sourceId;
    usedIds.add(id);
    idByTarget.set(target, id);
    pending.push({ id, target });
    result.push(id);
  }
  if (!pending.length) return result;
  state.generated = relationships;
  for (const item of pending) {
    appendExternalHyperlinkRelationship(relationships, item.id, item.target);
  }
  state.dirty = true;
  return result;
}

export function setDocxHyperlinkDestination(
  hyperlink: Element,
  root: Element,
  destination: DocxHyperlinkDestination,
  relationshipId?: string,
): boolean {
  removeNamespacedAttribute(
    hyperlink,
    'anchor',
    DOCX_WORDPROCESSING_NAMESPACES,
  );
  removeNamespacedAttribute(hyperlink, 'id', RELATIONSHIP_NAMESPACES);
  if (destination.kind === 'internal') {
    const namespace = hyperlink.namespaceURI;
    if (!namespace || !DOCX_WORDPROCESSING_NAMESPACES.has(namespace)) {
      return false;
    }
    setNamespacedAttribute(
      hyperlink,
      root,
      namespace,
      'w',
      'anchor',
      destination.target,
    );
    return true;
  }
  if (!relationshipId || !RELATIONSHIP_ID_PATTERN.test(relationshipId)) {
    return false;
  }
  setNamespacedAttribute(
    hyperlink,
    root,
    TRANSITIONAL_RELATIONSHIP_NAMESPACE,
    'r',
    'id',
    relationshipId,
  );
  return true;
}

export function flushDocxHyperlinkRelationships(
  state: DocxHyperlinkRelationshipState,
): void {
  if (!state.dirty || !state.generated) return;
  state.archive.file(state.path, serializeUtf8Xml(state.generated.document));
}

function normalizeAnchor(value: string): string | null {
  const anchor = value.trim();
  return anchor.length <= 255 && normalizeDocumentHref(`#${anchor}`)
    ? anchor
    : null;
}

function isExternalHyperlinkRelationship(
  relationship: RelationshipRecord,
): boolean {
  return (
    HYPERLINK_RELATIONSHIP_TYPES.has(relationship.type) &&
    relationship.targetMode.toLowerCase() === 'external'
  );
}

async function loadRelationships(
  archive: JSZip,
  path: string,
  role: string,
): Promise<{ document: RelationshipDocument | null; malformed: boolean }> {
  const entry = archive.file(path);
  if (!entry) return { document: null, malformed: false };
  try {
    const document = parseXml(
      decodeXmlBytes(await entry.async('uint8array'), `${role} DOCX ${path}`),
      `${role} DOCX ${path}`,
    );
    const root = document.documentElement;
    if (
      root.localName !== 'Relationships' ||
      root.namespaceURI !== PACKAGE_RELATIONSHIP_NAMESPACE
    ) {
      return { document: null, malformed: true };
    }
    const byId = new Map<string, RelationshipRecord>();
    const usedIds = new Set<string>();
    const elements = directChildren(root);
    if (elements.length > MAX_RELATIONSHIPS) {
      return { document: null, malformed: true };
    }
    for (const element of elements) {
      if (
        element.localName !== 'Relationship' ||
        element.namespaceURI !== PACKAGE_RELATIONSHIP_NAMESPACE
      ) {
        return { document: null, malformed: true };
      }
      const record = relationshipRecord(element);
      if (!record || usedIds.has(record.id)) {
        return { document: null, malformed: true };
      }
      byId.set(record.id, record);
      usedIds.add(record.id);
    }
    return { document: { byId, document, usedIds }, malformed: false };
  } catch {
    return { document: null, malformed: true };
  }
}

function relationshipRecord(element: Element): RelationshipRecord | null {
  const id = element.getAttribute('Id')?.trim() ?? '';
  const target = element.getAttribute('Target')?.trim() ?? '';
  const type = element.getAttribute('Type')?.trim() ?? '';
  const targetMode = element.getAttribute('TargetMode')?.trim() ?? '';
  if (!RELATIONSHIP_ID_PATTERN.test(id) || !target || !type) return null;
  return { id, target, targetMode, type };
}

function createRelationshipsDocument(): RelationshipDocument {
  const document = parseXml(
    `<Relationships xmlns="${PACKAGE_RELATIONSHIP_NAMESPACE}"/>`,
    'generated DOCX hyperlink relationships',
  );
  return { byId: new Map(), document, usedIds: new Set() };
}

function appendExternalHyperlinkRelationship(
  relationships: RelationshipDocument,
  id: string,
  target: string,
): void {
  const element = relationships.document.createElementNS(
    PACKAGE_RELATIONSHIP_NAMESPACE,
    'Relationship',
  );
  element.setAttribute('Id', id);
  element.setAttribute('Type', GENERATED_HYPERLINK_RELATIONSHIP_TYPE);
  element.setAttribute('Target', target);
  element.setAttribute('TargetMode', 'External');
  relationships.document.documentElement.append(element);
  relationships.byId.set(id, {
    id,
    target,
    targetMode: 'External',
    type: GENERATED_HYPERLINK_RELATIONSHIP_TYPE,
  });
  relationships.usedIds.add(id);
}

function uniqueNamespacedAttribute(
  element: Element,
  localName: string,
  namespaces: ReadonlySet<string>,
): string | null | undefined {
  const matches = Array.from(element.attributes).filter(
    (item) =>
      xmlAttributeLocalName(item) === localName &&
      namespaces.has(xmlAttributeNamespace(element, item) ?? ''),
  );
  return matches.length <= 1 ? (matches[0]?.value.trim() ?? null) : undefined;
}

function removeNamespacedAttribute(
  element: Element,
  localName: string,
  namespaces: ReadonlySet<string>,
): void {
  for (const item of Array.from(element.attributes)) {
    if (
      xmlAttributeLocalName(item) === localName &&
      namespaces.has(xmlAttributeNamespace(element, item) ?? '')
    ) {
      element.removeAttributeNode(item);
    }
  }
}

function setNamespacedAttribute(
  element: Element,
  root: Element,
  namespace: string,
  preferredPrefix: string,
  localName: string,
  value: string,
): void {
  let prefix =
    xmlDeclaredPrefix(element, namespace) ?? xmlDeclaredPrefix(root, namespace);
  if (!prefix) {
    prefix = availablePrefix(root, preferredPrefix);
    root.setAttributeNS(XMLNS_NAMESPACE, `xmlns:${prefix}`, namespace);
  }
  element.setAttributeNS(namespace, `${prefix}:${localName}`, value);
}

function availablePrefix(root: Element, preferred: string): string {
  if (!xmlNamespaceUri(root, preferred)) return preferred;
  let index = 1;
  while (xmlNamespaceUri(root, `${preferred}${index}`)) index += 1;
  return `${preferred}${index}`;
}

function nextRelationshipId(used: ReadonlySet<string>): string {
  let index = 1;
  while (used.has(`rId${index}`)) index += 1;
  return `rId${index}`;
}

function relationshipPartPath(ownerPart: string): string {
  const separator = ownerPart.lastIndexOf('/');
  const directory = separator < 0 ? '' : ownerPart.slice(0, separator + 1);
  const file = ownerPart.slice(separator + 1);
  return `${directory}_rels/${file}.rels`;
}
