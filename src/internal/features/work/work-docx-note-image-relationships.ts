import JSZip from 'jszip';
import {
  descendants,
  directChildren,
  parseXml,
  resolvePartTarget,
} from './work-ooxml-package';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';

const DRAWINGML_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/drawingml/2006/main',
  'http://purl.oclc.org/ooxml/drawingml/main',
]);
const RELATIONSHIP_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  'http://purl.oclc.org/ooxml/officeDocument/relationships',
]);
const PACKAGE_RELATIONSHIPS_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/relationships';
const IMAGE_RELATIONSHIP_TYPES = new Set([
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  'http://purl.oclc.org/ooxml/officeDocument/relationships/image',
]);
const NOTE_PARTS = ['word/footnotes.xml', 'word/endnotes.xml'] as const;
const MEDIA_PLACEHOLDER_PATTERN =
  /^rId\{([A-Za-z0-9][A-Za-z0-9._-]{0,255})\}$/u;
const MAX_NOTE_IMAGES = 4_096;

interface RelationshipsPart {
  document: Document;
  path: string;
  root: Element;
}

export async function patchDocxNoteImageRelationships(
  buffer: ArrayBuffer,
): Promise<ArrayBuffer> {
  const archive = await JSZip.loadAsync(buffer);
  let changed = false;
  let imageCount = 0;
  for (const partPath of NOTE_PARTS) {
    const entry = archive.file(partPath);
    if (!entry) continue;
    const document = parseXml(
      decodeXmlBytes(
        await entry.async('uint8array'),
        `generated DOCX ${partPath}`,
      ),
      `generated DOCX ${partPath}`,
    );
    const blips = descendants(document, 'blip').filter((element) =>
      DRAWINGML_NAMESPACES.has(element.namespaceURI ?? ''),
    );
    imageCount += blips.length;
    if (imageCount > MAX_NOTE_IMAGES) {
      throw new Error('Generated DOCX exceeds the note-image limit.');
    }
    if (!blips.length) continue;
    const relationships = await loadRelationshipsPart(archive, partPath);
    const usedIds = new Set<string>();
    const byId = new Map<string, Element>();
    const byTarget = new Map<string, string>();
    for (const relationship of directChildren(
      relationships.root,
      'Relationship',
    )) {
      const id = relationship.getAttribute('Id');
      if (!id) continue;
      if (byId.has(id)) {
        throw new Error(
          'Generated DOCX note relationships contain duplicate IDs.',
        );
      }
      usedIds.add(id);
      byId.set(id, relationship);
      const resolvedTarget = validImageTarget(archive, partPath, relationship);
      if (resolvedTarget && !byTarget.has(resolvedTarget)) {
        byTarget.set(resolvedTarget, id);
      }
    }
    let partChanged = false;
    let relationshipsChanged = false;
    for (const blip of blips) {
      const embed = relationshipAttribute(blip, 'embed');
      const match = embed ? MEDIA_PLACEHOLDER_PATTERN.exec(embed.value) : null;
      if (!embed || !match) continue;
      const target = `media/${match[1]}`;
      const resolvedTarget = resolvePartTarget(partPath, target);
      if (!archive.file(resolvedTarget)) {
        throw new Error(
          'Generated DOCX note image references a missing media payload.',
        );
      }
      let id = byTarget.get(resolvedTarget);
      if (!id) {
        id = nextRelationshipId(usedIds);
        usedIds.add(id);
        const relationship = relationships.document.createElementNS(
          PACKAGE_RELATIONSHIPS_NAMESPACE,
          'Relationship',
        );
        relationship.setAttribute('Id', id);
        relationship.setAttribute(
          'Type',
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
        );
        relationship.setAttribute('Target', target);
        relationships.root.append(relationship);
        byId.set(id, relationship);
        byTarget.set(resolvedTarget, id);
        relationshipsChanged = true;
      }
      embed.value = id;
      partChanged = true;
    }
    validateImageRelationships(archive, partPath, blips, byId);
    if (partChanged) {
      archive.file(partPath, serializeUtf8Xml(document));
      changed = true;
    }
    if (relationshipsChanged) {
      archive.file(
        relationships.path,
        serializeUtf8Xml(relationships.document),
      );
      changed = true;
    }
  }
  return changed ? archive.generateAsync({ type: 'arraybuffer' }) : buffer;
}

async function loadRelationshipsPart(
  archive: JSZip,
  sourcePart: string,
): Promise<RelationshipsPart> {
  const path = relationshipsPartPath(sourcePart);
  const entry = archive.file(path);
  const document = entry
    ? parseXml(
        decodeXmlBytes(
          await entry.async('uint8array'),
          `generated DOCX ${path}`,
        ),
        `generated DOCX ${path}`,
      )
    : parseXml(
        `<Relationships xmlns="${PACKAGE_RELATIONSHIPS_NAMESPACE}"/>`,
        `generated DOCX ${path}`,
      );
  const root = document.documentElement;
  if (
    root.localName !== 'Relationships' ||
    root.namespaceURI !== PACKAGE_RELATIONSHIPS_NAMESPACE
  ) {
    throw new Error('Generated DOCX note relationships part is malformed.');
  }
  return { document, path, root };
}

function validateImageRelationships(
  archive: JSZip,
  sourcePart: string,
  blips: readonly Element[],
  byId: ReadonlyMap<string, Element>,
): void {
  for (const blip of blips) {
    const id = relationshipAttribute(blip, 'embed')?.value;
    const relationship = id ? byId.get(id) : null;
    if (!relationship || !validImageTarget(archive, sourcePart, relationship)) {
      throw new Error('Generated DOCX note image relationship is invalid.');
    }
  }
}

function validImageTarget(
  archive: JSZip,
  sourcePart: string,
  relationship: Element,
): string | null {
  const type = relationship.getAttribute('Type');
  const target = relationship.getAttribute('Target');
  const mode = relationship.getAttribute('TargetMode');
  if (
    !type ||
    !IMAGE_RELATIONSHIP_TYPES.has(type) ||
    !target ||
    mode !== null
  ) {
    return null;
  }
  const resolved = resolvePartTarget(sourcePart, target);
  return resolved.startsWith('word/media/') && archive.file(resolved)
    ? resolved
    : null;
}

function relationshipAttribute(
  element: Element,
  localName: string,
): Attr | null {
  return (
    Array.from(element.attributes).find(
      (item) =>
        xmlAttributeLocalName(item) === localName &&
        RELATIONSHIP_NAMESPACES.has(xmlAttributeNamespace(element, item) ?? ''),
    ) ?? null
  );
}

function relationshipsPartPath(sourcePart: string): string {
  const separator = sourcePart.lastIndexOf('/');
  const directory = sourcePart.slice(0, separator + 1);
  const fileName = sourcePart.slice(separator + 1);
  return `${directory}_rels/${fileName}.rels`;
}

function nextRelationshipId(used: ReadonlySet<string>): string {
  let index = 1;
  while (used.has(`rId${index}`)) index += 1;
  return `rId${index}`;
}
