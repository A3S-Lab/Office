import type JSZip from 'jszip';
import {
  normalizeDocumentParagraphId,
  normalizeDocumentParagraphIdentity,
} from './work-document-paragraph-identity';
import {
  DOCX_WORDPROCESSING_NAMESPACES,
  mergeDocxIgnorableExtensionsAtPairs,
  type DocxExtensionDocumentRole,
  type DocxIgnorableExtensionPair,
} from './work-docx-ignorable-extension-preservation';
import { descendants, directChildren, parseXml } from './work-ooxml-package';
import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';

const WORD_2010_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2010/wordml';
const MARKUP_COMPATIBILITY_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';
const MAX_TABLE_PARTS = 512;
const MAX_TABLE_SCOPES = 65_536;

type TablePartFamily = 'document' | 'footer' | 'header';
type TableScopeKind = 'cell' | 'row' | 'table';

interface TablePart {
  document: Document;
  family: TablePartFamily;
  path: string;
}

interface WordIdentity {
  paragraphId: string;
  textId: string;
}

interface TableScopeRecord {
  componentKeys: string[];
  element: Element;
  identityKey: string;
  kind: TableScopeKind;
  part: TablePart;
}

interface TableScopeIndex {
  byIdentity: Map<string, TableScopeRecord[]>;
  componentCounts: Map<string, number>;
}

interface TableMergeBatch {
  generated: TablePart;
  pairs: DocxIgnorableExtensionPair[];
  source: TablePart;
}

export async function preserveDocxTableExtensions(
  generated: JSZip,
  source: JSZip,
  generatedPaths: readonly string[],
  sourcePaths: readonly string[],
): Promise<void> {
  if (
    generatedPaths.length > MAX_TABLE_PARTS ||
    sourcePaths.length > MAX_TABLE_PARTS
  ) {
    throw new Error('Registered source DOCX exceeds the table-part limit.');
  }
  const [generatedParts, sourceParts] = await Promise.all([
    loadTableParts(generated, generatedPaths, 'generated'),
    loadTableParts(source, sourcePaths, 'source'),
  ]);
  const generatedIndex = indexTableScopes(generatedParts, false);
  const sourceIndex = indexTableScopes(sourceParts, true);
  const batches = tableMergeBatches(generatedIndex, sourceIndex);
  const changedParts = new Set<TablePart>();
  for (const batch of batches) {
    mergeDocxIgnorableExtensionsAtPairs(
      batch.generated.document,
      batch.source.document,
      batch.pairs,
      {
        semanticKey: tableSemanticKey,
        isAdditionalSemanticNamespace: isKnownOoxmlNamespace,
        allowExtensionNamespace: (namespace) =>
          !isKnownOoxmlNamespace(namespace),
        allowMatchedElementMerge: allowTableScopeMerge,
      },
    );
    changedParts.add(batch.generated);
  }
  for (const part of changedParts) {
    generated.file(part.path, serializeUtf8Xml(part.document));
  }
}

async function loadTableParts(
  archive: JSZip,
  paths: readonly string[],
  role: 'generated' | 'source',
): Promise<TablePart[]> {
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
        const family = tablePartFamily(document.documentElement, path);
        return family ? { document, family, path } : null;
      } catch {
        return null;
      }
    }),
  );
  return parts.filter((part): part is TablePart => Boolean(part));
}

function tablePartFamily(root: Element, path: string): TablePartFamily | null {
  const family: TablePartFamily = /^word\/header\d*\.xml$/i.test(path)
    ? 'header'
    : /^word\/footer\d*\.xml$/i.test(path)
      ? 'footer'
      : 'document';
  const expected =
    family === 'header' ? 'hdr' : family === 'footer' ? 'ftr' : 'document';
  return root.localName === expected && isWordElement(root) ? family : null;
}

function indexTableScopes(
  parts: readonly TablePart[],
  source: boolean,
): TableScopeIndex {
  const byIdentity = new Map<string, TableScopeRecord[]>();
  const componentCounts = new Map<string, number>();
  const identities = new Map<Element, WordIdentity>();
  let scopeCount = 0;
  for (const part of parts) {
    for (const localName of ['p', 'tr'] as const) {
      for (const element of wordDescendants(part.document, localName)) {
        const paragraphId = normalizeDocumentParagraphId(
          word2010Attribute(element, 'paraId'),
        );
        if (paragraphId) {
          const key = componentKey(part.family, paragraphId);
          componentCounts.set(key, (componentCounts.get(key) ?? 0) + 1);
        }
        const identity = wordIdentity(element);
        if (!identity) continue;
        identities.set(element, identity);
      }
    }
  }
  for (const part of parts) {
    for (const table of wordDescendants(part.document, 'tbl')) {
      scopeCount = registerTableScope(
        table,
        part,
        identities,
        byIdentity,
        scopeCount,
        source,
      );
    }
    for (const row of wordDescendants(part.document, 'tr')) {
      scopeCount = registerRowScope(
        row,
        part,
        identities,
        byIdentity,
        scopeCount,
        source,
      );
    }
    for (const cell of wordDescendants(part.document, 'tc')) {
      scopeCount = registerCellScope(
        cell,
        part,
        identities,
        byIdentity,
        scopeCount,
        source,
      );
    }
  }
  return { byIdentity, componentCounts };
}

function registerTableScope(
  table: Element,
  part: TablePart,
  identities: ReadonlyMap<Element, WordIdentity>,
  records: Map<string, TableScopeRecord[]>,
  count: number,
  source: boolean,
): number {
  const rows = ownedDescendants(table, 'tr', 'tbl');
  const rowIdentities = rows.map((row) => identities.get(row));
  if (!rows.length || rowIdentities.some((identity) => !identity)) return count;
  const paragraphIds = rowIdentities.map(
    (identity) => (identity as WordIdentity).paragraphId,
  );
  return registerScope(
    {
      componentKeys: paragraphIds.map((id) => componentKey(part.family, id)),
      element: table,
      identityKey: scopeKey('table', part.family, paragraphIds),
      kind: 'table',
      part,
    },
    records,
    count,
    source,
  );
}

function registerRowScope(
  row: Element,
  part: TablePart,
  identities: ReadonlyMap<Element, WordIdentity>,
  records: Map<string, TableScopeRecord[]>,
  count: number,
  source: boolean,
): number {
  const identity = identities.get(row);
  if (!identity) return count;
  return registerScope(
    {
      componentKeys: [componentKey(part.family, identity.paragraphId)],
      element: row,
      identityKey: scopeKey('row', part.family, [
        identity.paragraphId,
        identity.textId,
      ]),
      kind: 'row',
      part,
    },
    records,
    count,
    source,
  );
}

function registerCellScope(
  cell: Element,
  part: TablePart,
  identities: ReadonlyMap<Element, WordIdentity>,
  records: Map<string, TableScopeRecord[]>,
  count: number,
  source: boolean,
): number {
  const row = closestWordAncestor(cell.parentElement, 'tr');
  const rowIdentity = row ? identities.get(row) : null;
  const paragraphs = ownedDescendants(cell, 'p', 'tc');
  const paragraphIdentities = paragraphs.map((item) => identities.get(item));
  const stableParagraphIdentities = paragraphIdentities.filter(
    (identity): identity is WordIdentity => Boolean(identity),
  );
  if (
    !rowIdentity ||
    !stableParagraphIdentities.length ||
    paragraphs.some(
      (paragraph, index) =>
        !paragraphIdentities[index] &&
        !isGeneratedTrailingCellParagraph(cell, paragraph),
    )
  ) {
    return count;
  }
  const paragraphIds = stableParagraphIdentities.map(
    (identity) => identity.paragraphId,
  );
  return registerScope(
    {
      componentKeys: [
        componentKey(part.family, rowIdentity.paragraphId),
        ...paragraphIds.map((id) => componentKey(part.family, id)),
      ],
      element: cell,
      identityKey: scopeKey('cell', part.family, [
        rowIdentity.paragraphId,
        ...paragraphIds,
      ]),
      kind: 'cell',
      part,
    },
    records,
    count,
    source,
  );
}

function isGeneratedTrailingCellParagraph(
  cell: Element,
  paragraph: Element,
): boolean {
  if (
    paragraph.textContent?.trim() ||
    directChildren(paragraph).some((child) => child.localName !== 'pPr')
  ) {
    return false;
  }
  const blocks = directChildren(cell).filter(isWordElement);
  const index = blocks.indexOf(paragraph);
  return (
    index === blocks.length - 1 &&
    blocks.slice(0, index).some((child) => child.localName === 'tbl')
  );
}

function registerScope(
  record: TableScopeRecord,
  records: Map<string, TableScopeRecord[]>,
  count: number,
  source: boolean,
): number {
  const nextCount = count + 1;
  if (nextCount > MAX_TABLE_SCOPES) {
    throw new Error(
      `${source ? 'Registered source' : 'Generated'} DOCX exceeds the table-scope limit.`,
    );
  }
  const matches = records.get(record.identityKey) ?? [];
  matches.push(record);
  records.set(record.identityKey, matches);
  return nextCount;
}

function tableMergeBatches(
  generated: TableScopeIndex,
  source: TableScopeIndex,
): TableMergeBatch[] {
  const batches = new Map<string, TableMergeBatch>();
  for (const [identity, sourceRecords] of source.byIdentity) {
    const generatedRecords = generated.byIdentity.get(identity) ?? [];
    if (sourceRecords.length !== 1 || generatedRecords.length !== 1) continue;
    const sourceRecord = sourceRecords[0];
    const generatedRecord = generatedRecords[0];
    if (
      sourceRecord.kind !== generatedRecord.kind ||
      !hasUniqueComponents(sourceRecord, source) ||
      !hasUniqueComponents(generatedRecord, generated)
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

function hasUniqueComponents(
  record: TableScopeRecord,
  index: TableScopeIndex,
): boolean {
  return record.componentKeys.every(
    (key) => index.componentCounts.get(key) === 1,
  );
}

function wordIdentity(element: Element): WordIdentity | null {
  return normalizeDocumentParagraphIdentity({
    paragraphId: word2010Attribute(element, 'paraId'),
    textId: word2010Attribute(element, 'textId'),
  });
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

function wordDescendants(document: Document, localName: string): Element[] {
  return descendants(document, localName).filter(isWordElement);
}

function ownedDescendants(
  owner: Element,
  localName: string,
  ownerLocalName: string,
): Element[] {
  return descendants(owner, localName).filter(
    (element) =>
      isWordElement(element) &&
      closestWordAncestor(element.parentElement, ownerLocalName) === owner,
  );
}

function closestWordAncestor(
  element: Element | null,
  localName: string,
): Element | null {
  let current = element;
  while (current) {
    if (current.localName === localName && isWordElement(current)) {
      return current;
    }
    current = current.parentElement;
  }
  return null;
}

function isWordElement(element: Element): boolean {
  return DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '');
}

function componentKey(family: TablePartFamily, paragraphId: string): string {
  return `${family}\u0000${paragraphId}`;
}

function scopeKey(
  kind: TableScopeKind,
  family: TablePartFamily,
  components: readonly string[],
): string {
  return `${kind}\u0000${family}\u0000${components.join('\u0001')}`;
}

function allowTableScopeMerge(
  generated: Element,
  _source: Element,
  depth: number,
): boolean {
  if (depth === 0 || depth > 1) return true;
  return (
    ['tblPr', 'trPr', 'tcPr'].includes(generated.localName) &&
    isWordElement(generated)
  );
}

function tableSemanticKey(
  element: Element,
  _role: DocxExtensionDocumentRole,
): string {
  return isWordElement(element)
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
