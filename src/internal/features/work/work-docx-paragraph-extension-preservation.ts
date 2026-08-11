import type JSZip from 'jszip';
import { normalizeDocumentParagraphIdentity } from './work-document-paragraph-identity';
import {
  DOCX_WORDPROCESSING_NAMESPACES,
  mergeDocxIgnorableExtensionsAtPairs,
  type DocxExtensionDocumentRole,
  type DocxIgnorableExtensionPair,
} from './work-docx-ignorable-extension-preservation';
import { descendants, parseXml } from './work-ooxml-package';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';

const WORD_2010_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2010/wordml';
const MARKUP_COMPATIBILITY_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';
const MAX_PARAGRAPH_PARTS = 512;
const MAX_PARAGRAPH_IDENTITIES = 65_536;

type ParagraphPartFamily = 'document' | 'footer' | 'header';

interface ParagraphPart {
  document: Document;
  family: ParagraphPartFamily;
  path: string;
}

interface ParagraphRecord {
  element: Element;
  identityKey: string;
  paragraphIdKey: string;
  part: ParagraphPart;
}

interface ParagraphIndex {
  byIdentity: Map<string, ParagraphRecord[]>;
  paragraphIdCounts: Map<string, number>;
}

interface ParagraphMergeBatch {
  generated: ParagraphPart;
  pairs: DocxIgnorableExtensionPair[];
  source: ParagraphPart;
}

export async function preserveDocxParagraphExtensions(
  generated: JSZip,
  source: JSZip,
  generatedPaths: readonly string[],
  sourcePaths: readonly string[],
): Promise<void> {
  if (
    generatedPaths.length > MAX_PARAGRAPH_PARTS ||
    sourcePaths.length > MAX_PARAGRAPH_PARTS
  ) {
    throw new Error('Registered source DOCX exceeds the paragraph-part limit.');
  }
  const [generatedParts, sourceParts] = await Promise.all([
    loadParagraphParts(generated, generatedPaths, 'generated'),
    loadParagraphParts(source, sourcePaths, 'source'),
  ]);
  const generatedIndex = indexParagraphs(generatedParts, false);
  const sourceIndex = indexParagraphs(sourceParts, true);
  const batches = paragraphMergeBatches(generatedIndex, sourceIndex);
  const changedParts = new Set<ParagraphPart>();
  for (const batch of batches) {
    mergeDocxIgnorableExtensionsAtPairs(
      batch.generated.document,
      batch.source.document,
      batch.pairs,
      {
        semanticKey: paragraphSemanticKey,
        isAdditionalSemanticNamespace: isKnownOoxmlNamespace,
        allowExtensionNamespace: (namespace) =>
          !isKnownOoxmlNamespace(namespace),
        allowMatchedElementMerge: allowParagraphScopeMerge,
      },
    );
    changedParts.add(batch.generated);
  }
  for (const part of changedParts) {
    generated.file(part.path, serializeUtf8Xml(part.document));
  }
}

async function loadParagraphParts(
  archive: JSZip,
  paths: readonly string[],
  role: 'generated' | 'source',
): Promise<ParagraphPart[]> {
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
        const family = paragraphPartFamily(document.documentElement, path);
        return family ? { document, family, path } : null;
      } catch {
        return null;
      }
    }),
  );
  return parts.filter((part): part is ParagraphPart => Boolean(part));
}

function paragraphPartFamily(
  root: Element,
  path: string,
): ParagraphPartFamily | null {
  const family: ParagraphPartFamily = /^word\/header\d*\.xml$/i.test(path)
    ? 'header'
    : /^word\/footer\d*\.xml$/i.test(path)
      ? 'footer'
      : 'document';
  const expected =
    family === 'header' ? 'hdr' : family === 'footer' ? 'ftr' : 'document';
  return root.localName === expected &&
    DOCX_WORDPROCESSING_NAMESPACES.has(root.namespaceURI ?? '')
    ? family
    : null;
}

function indexParagraphs(
  parts: readonly ParagraphPart[],
  source: boolean,
): ParagraphIndex {
  const byIdentity = new Map<string, ParagraphRecord[]>();
  const paragraphIdCounts = new Map<string, number>();
  let count = 0;
  for (const part of parts) {
    for (const element of descendants(part.document, 'p')) {
      if (!DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '')) {
        continue;
      }
      count += 1;
      if (count > MAX_PARAGRAPH_IDENTITIES) {
        throw new Error(
          `${source ? 'Registered source' : 'Generated'} DOCX exceeds the paragraph-identity limit.`,
        );
      }
      const identity = normalizeDocumentParagraphIdentity({
        paragraphId: word2010Attribute(element, 'paraId'),
        textId: word2010Attribute(element, 'textId'),
      });
      if (!identity) continue;
      const paragraphIdKey = `${part.family}\u0000${identity.paragraphId}`;
      const identityKey = `${paragraphIdKey}\u0000${identity.textId}`;
      paragraphIdCounts.set(
        paragraphIdKey,
        (paragraphIdCounts.get(paragraphIdKey) ?? 0) + 1,
      );
      const records = byIdentity.get(identityKey) ?? [];
      records.push({ element, identityKey, paragraphIdKey, part });
      byIdentity.set(identityKey, records);
    }
  }
  return { byIdentity, paragraphIdCounts };
}

function paragraphMergeBatches(
  generated: ParagraphIndex,
  source: ParagraphIndex,
): ParagraphMergeBatch[] {
  const batches = new Map<string, ParagraphMergeBatch>();
  for (const [identity, sourceRecords] of source.byIdentity) {
    const generatedRecords = generated.byIdentity.get(identity) ?? [];
    if (sourceRecords.length !== 1 || generatedRecords.length !== 1) continue;
    const sourceRecord = sourceRecords[0];
    const generatedRecord = generatedRecords[0];
    if (
      source.paragraphIdCounts.get(sourceRecord.paragraphIdKey) !== 1 ||
      generated.paragraphIdCounts.get(generatedRecord.paragraphIdKey) !== 1
    ) {
      continue;
    }
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

function word2010Attribute(element: Element, name: string): string | null {
  return (
    Array.from(element.attributes).find(
      (item) =>
        xmlAttributeLocalName(item) === name &&
        xmlAttributeNamespace(element, item) === WORD_2010_NAMESPACE,
    )?.value ?? null
  );
}

function allowParagraphScopeMerge(
  generated: Element,
  _source: Element,
  depth: number,
): boolean {
  if (depth === 0) return true;
  if (depth > 1) return true;
  return (
    generated.localName === 'pPr' &&
    DOCX_WORDPROCESSING_NAMESPACES.has(generated.namespaceURI ?? '')
  );
}

function paragraphSemanticKey(
  element: Element,
  _role: DocxExtensionDocumentRole,
): string {
  return DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '')
    ? `word:${element.localName}`
    : `{${element.namespaceURI ?? ''}}${element.localName}`;
}

function isKnownOoxmlNamespace(namespace: string): boolean {
  if (namespace === MARKUP_COMPATIBILITY_NAMESPACE) return false;
  return (
    DOCX_WORDPROCESSING_NAMESPACES.has(namespace) ||
    namespace.startsWith('http://schemas.microsoft.com/office/') ||
    namespace.startsWith('http://schemas.openxmlformats.org/') ||
    namespace.startsWith('http://purl.oclc.org/ooxml/') ||
    namespace.startsWith('urn:schemas-microsoft-com:') ||
    namespace.startsWith('urn:microsoft-com:office:')
  );
}
