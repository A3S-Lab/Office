import type JSZip from 'jszip';
import { normalizeDocumentImageIdentity } from './work-document-image-identity';
import {
  DOCX_WORDPROCESSING_NAMESPACES,
  mergeDocxIgnorableExtensionsAtPairs,
  type DocxExtensionDocumentRole,
  type DocxIgnorableExtensionPair,
} from './work-docx-ignorable-extension-preservation';
import { directChildren, parseXml } from './work-ooxml-package';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';

const WORDPROCESSING_DRAWING_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  'http://purl.oclc.org/ooxml/drawingml/wordprocessingDrawing',
]);
const DRAWINGML_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/drawingml/2006/main',
  'http://purl.oclc.org/ooxml/drawingml/main',
]);
const PICTURE_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/drawingml/2006/picture',
  'http://purl.oclc.org/ooxml/drawingml/picture',
]);
const WORDPROCESSING_DRAWING_2010_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing';
const MARKUP_COMPATIBILITY_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';
const MAX_DRAWING_PARTS = 512;
const MAX_DRAWING_IDENTITIES = 4_096;

interface DrawingPart {
  document: Document;
  path: string;
}

interface DrawingRecord {
  element: Element;
  part: DrawingPart;
}

interface DrawingMergeBatch {
  generated: DrawingPart;
  pairs: DocxIgnorableExtensionPair[];
  source: DrawingPart;
}

export async function preserveDocxDrawingExtensions(
  generated: JSZip,
  source: JSZip,
  generatedPaths: readonly string[],
  sourcePaths: readonly string[],
): Promise<void> {
  if (
    generatedPaths.length > MAX_DRAWING_PARTS ||
    sourcePaths.length > MAX_DRAWING_PARTS
  ) {
    throw new Error('Registered source DOCX exceeds the drawing-part limit.');
  }
  const [generatedParts, sourceParts] = await Promise.all([
    loadDrawingParts(generated, generatedPaths, 'generated'),
    loadDrawingParts(source, sourcePaths, 'source'),
  ]);
  const generatedIndex = indexDrawings(generatedParts, false);
  const sourceIndex = indexDrawings(sourceParts, true);
  const batches = drawingMergeBatches(generatedIndex, sourceIndex);
  const changedParts = new Set<DrawingPart>();
  for (const batch of batches) {
    mergeDocxIgnorableExtensionsAtPairs(
      batch.generated.document,
      batch.source.document,
      batch.pairs,
      {
        semanticKey: drawingSemanticKey,
        isAdditionalSemanticNamespace: isKnownOoxmlDrawingNamespace,
        allowExtensionNamespace: (namespace) =>
          !isKnownOoxmlDrawingNamespace(namespace),
      },
    );
    changedParts.add(batch.generated);
  }
  for (const part of changedParts) {
    generated.file(part.path, serializeUtf8Xml(part.document));
  }
}

async function loadDrawingParts(
  archive: JSZip,
  paths: readonly string[],
  role: 'generated' | 'source',
): Promise<DrawingPart[]> {
  const parts = await Promise.all(
    paths.map(async (path) => {
      const entry = archive.file(path);
      if (!entry) return null;
      try {
        const document = parseXml(
          decodeXmlBytes(
            await entry.async('uint8array'),
            `${role} DOCX ${path}`,
          ),
          `${role} DOCX ${path}`,
        );
        return isDrawingPartRoot(document.documentElement, path)
          ? { document, path }
          : null;
      } catch {
        return null;
      }
    }),
  );
  return parts.filter((part): part is DrawingPart => Boolean(part));
}

function isDrawingPartRoot(root: Element, path: string): boolean {
  const expected = /^word\/header\d*\.xml$/i.test(path)
    ? 'hdr'
    : /^word\/footer\d*\.xml$/i.test(path)
      ? 'ftr'
      : 'document';
  return (
    root.localName === expected &&
    DOCX_WORDPROCESSING_NAMESPACES.has(root.namespaceURI ?? '')
  );
}

function indexDrawings(
  parts: readonly DrawingPart[],
  source: boolean,
): Map<string, DrawingRecord[]> {
  const result = new Map<string, DrawingRecord[]>();
  let count = 0;
  for (const part of parts) {
    for (const element of Array.from(
      part.document.querySelectorAll('*'),
    ).filter(isWordprocessingDrawing)) {
      count += 1;
      if (count > MAX_DRAWING_IDENTITIES) {
        throw new Error(
          `${source ? 'Registered source' : 'Generated'} DOCX exceeds the drawing-identity limit.`,
        );
      }
      const identity = drawingIdentity(element);
      if (!identity) continue;
      const records = result.get(identity) ?? [];
      records.push({ element, part });
      result.set(identity, records);
    }
  }
  return result;
}

function isWordprocessingDrawing(element: Element): boolean {
  return (
    (element.localName === 'anchor' || element.localName === 'inline') &&
    WORDPROCESSING_DRAWING_NAMESPACES.has(element.namespaceURI ?? '')
  );
}

function drawingIdentity(drawing: Element): string | null {
  const properties = directChildren(drawing, 'docPr').find((item) =>
    WORDPROCESSING_DRAWING_NAMESPACES.has(item.namespaceURI ?? ''),
  );
  if (!properties) return null;
  const identity = normalizeDocumentImageIdentity({
    docPropertiesId: properties.getAttribute('id'),
    anchorId: drawing2010Attribute(drawing, 'anchorId'),
    editId: drawing2010Attribute(drawing, 'editId'),
  });
  return identity
    ? `${identity.anchorId}\u0000${identity.docPropertiesId}`
    : null;
}

function drawing2010Attribute(element: Element, name: string): string | null {
  for (const item of Array.from(element.attributes)) {
    if (
      xmlAttributeLocalName(item) === name &&
      xmlAttributeNamespace(element, item) ===
        WORDPROCESSING_DRAWING_2010_NAMESPACE
    ) {
      return item.value;
    }
  }
  return null;
}

function drawingMergeBatches(
  generated: ReadonlyMap<string, readonly DrawingRecord[]>,
  source: ReadonlyMap<string, readonly DrawingRecord[]>,
): DrawingMergeBatch[] {
  const batches = new Map<string, DrawingMergeBatch>();
  for (const [identity, sourceRecords] of source) {
    const generatedRecords = generated.get(identity) ?? [];
    if (sourceRecords.length !== 1 || generatedRecords.length !== 1) continue;
    const sourceRecord = sourceRecords[0];
    const generatedRecord = generatedRecords[0];
    const key = `${generatedRecord.part.path}\u0000${sourceRecord.part.path}`;
    const batch = batches.get(key) ?? {
      generated: generatedRecord.part,
      source: sourceRecord.part,
      pairs: [],
    };
    batch.pairs.push({
      generated: generatedRecord.element,
      source: sourceRecord.element,
    });
    batches.set(key, batch);
  }
  return [...batches.values()];
}

function drawingSemanticKey(
  element: Element,
  _role: DocxExtensionDocumentRole,
): string {
  const family = drawingNamespaceFamily(element.namespaceURI);
  if (family === 'a' && element.localName === 'ext') {
    return `ooxml:a:ext:${element.getAttribute('uri') ?? ''}`;
  }
  return family
    ? `ooxml:${family}:${element.localName}`
    : `{${element.namespaceURI ?? ''}}${element.localName}`;
}

function drawingNamespaceFamily(namespace: string | null): string | null {
  if (!namespace) return null;
  if (DOCX_WORDPROCESSING_NAMESPACES.has(namespace)) return 'w';
  if (WORDPROCESSING_DRAWING_NAMESPACES.has(namespace)) return 'wp';
  if (DRAWINGML_NAMESPACES.has(namespace)) return 'a';
  if (PICTURE_NAMESPACES.has(namespace)) return 'pic';
  return null;
}

function isKnownOoxmlDrawingNamespace(namespace: string): boolean {
  if (namespace === MARKUP_COMPATIBILITY_NAMESPACE) return false;
  return (
    Boolean(drawingNamespaceFamily(namespace)) ||
    namespace.startsWith('http://schemas.microsoft.com/office/') ||
    namespace.startsWith('http://schemas.openxmlformats.org/') ||
    namespace.startsWith('http://purl.oclc.org/ooxml/') ||
    namespace.startsWith('urn:schemas-microsoft-com:') ||
    namespace.startsWith('urn:microsoft-com:office:')
  );
}
