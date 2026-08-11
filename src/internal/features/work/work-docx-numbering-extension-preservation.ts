import type JSZip from 'jszip';
import {
  DOCX_WORDPROCESSING_NAMESPACES,
  mergeDocxIgnorableExtensions,
  type DocxExtensionDocumentRole,
} from './work-docx-ignorable-extension-preservation';
import { directChild, directChildren, parseXml } from './work-ooxml-package';
import {
  assertXmlRoot,
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';

const NUMBERING_PATH = 'word/numbering.xml';
const MAX_NUMBERING_RECORDS = 4_096;

export interface DocxSourceNumberingIdentity {
  numId: number;
  abstractNumId?: number;
}

interface NumberingIndex {
  abstracts: ReadonlyMap<string, Element>;
  nums: ReadonlyMap<string, Element>;
}

interface NumberingIdentityMap {
  abstracts: ReadonlyMap<string, string>;
  nums: ReadonlyMap<string, string>;
}

export async function preserveDocxNumberingExtensions(
  generated: JSZip,
  source: JSZip,
  sourceIdentities: readonly (DocxSourceNumberingIdentity | null)[],
  generatedPath = NUMBERING_PATH,
  sourcePath = NUMBERING_PATH,
): Promise<void> {
  const generatedEntry = generated.file(generatedPath);
  const sourceEntry = source.file(sourcePath);
  if (!generatedEntry || !sourceEntry) return;
  const [generatedBytes, sourceBytes] = await Promise.all([
    generatedEntry.async('uint8array'),
    sourceEntry.async('uint8array'),
  ]);
  const generatedDocument = parseXml(
    decodeXmlBytes(generatedBytes, `generated DOCX ${generatedPath}`),
    `generated DOCX ${generatedPath}`,
  );
  const sourceDocument = parseXml(
    decodeXmlBytes(sourceBytes, `source DOCX ${sourcePath}`),
    `source DOCX ${sourcePath}`,
  );
  assertXmlRoot(
    generatedDocument.documentElement,
    'numbering',
    DOCX_WORDPROCESSING_NAMESPACES,
    'Generated DOCX word/numbering.xml is not WordprocessingML numbering.',
  );
  assertXmlRoot(
    sourceDocument.documentElement,
    'numbering',
    DOCX_WORDPROCESSING_NAMESPACES,
    'Registered source DOCX word/numbering.xml is not WordprocessingML numbering.',
  );
  const sourceIndex = indexNumbering(sourceDocument, true);
  const generatedIndex = indexNumbering(generatedDocument, false);
  const identityMap = numberingIdentityMap(
    generatedDocument,
    sourceIndex,
    generatedIndex,
    sourceIdentities,
  );
  mergeDocxIgnorableExtensions(generatedDocument, sourceDocument, {
    semanticKey: (element, role) =>
      numberingSemanticKey(element, role, identityMap),
  });
  generated.file(generatedPath, serializeUtf8Xml(generatedDocument));
}

function indexNumbering(document: Document, validate: boolean): NumberingIndex {
  const abstracts = indexNumberingChildren(
    document.documentElement,
    'abstractNum',
    'abstractNumId',
    validate,
  );
  const nums = indexNumberingChildren(
    document.documentElement,
    'num',
    'numId',
    validate,
  );
  if (validate) {
    for (const abstract of abstracts.values()) {
      assertUniqueLevelIdentities(abstract, 'lvl');
    }
    for (const num of nums.values()) {
      assertUniqueLevelIdentities(num, 'lvlOverride');
    }
  }
  return { abstracts, nums };
}

function indexNumberingChildren(
  root: Element,
  localName: string,
  attributeName: string,
  validate: boolean,
): Map<string, Element> {
  const children = directChildren(root, localName).filter((item) =>
    DOCX_WORDPROCESSING_NAMESPACES.has(item.namespaceURI ?? ''),
  );
  if (children.length > MAX_NUMBERING_RECORDS) {
    throw new Error(
      'Registered source DOCX exceeds the numbering-record limit.',
    );
  }
  const result = new Map<string, Element>();
  for (const child of children) {
    const id = decimalWordAttribute(child, attributeName);
    if (!id) continue;
    if (validate && result.has(id)) {
      throw new Error(
        'Registered source DOCX numbering contains duplicate numbering identities.',
      );
    }
    result.set(id, child);
  }
  return result;
}

function assertUniqueLevelIdentities(parent: Element, localName: string): void {
  const levels = new Set<string>();
  for (const level of directChildren(parent, localName).filter((item) =>
    DOCX_WORDPROCESSING_NAMESPACES.has(item.namespaceURI ?? ''),
  )) {
    const id = decimalWordAttribute(level, 'ilvl');
    if (!id) continue;
    if (levels.has(id)) {
      throw new Error(
        'Registered source DOCX numbering contains duplicate level identities.',
      );
    }
    levels.add(id);
  }
}

function numberingIdentityMap(
  generated: Document,
  sourceIndex: NumberingIndex,
  generatedIndex: NumberingIndex,
  sourceIdentities: readonly (DocxSourceNumberingIdentity | null)[],
): NumberingIdentityMap {
  const abstractCandidates = new Map<string, Set<string>>();
  const numCandidates = new Map<string, Set<string>>();
  const generatedAbstracts = directChildren(
    generated.documentElement,
    'abstractNum',
  ).filter((item) =>
    DOCX_WORDPROCESSING_NAMESPACES.has(item.namespaceURI ?? ''),
  );
  if (generatedAbstracts.length < sourceIdentities.length) {
    return { abstracts: new Map(), nums: new Map() };
  }
  const generatedConfigurations = generatedAbstracts.slice(
    generatedAbstracts.length - sourceIdentities.length,
  );
  const generatedNums = Array.from(generatedIndex.nums.values());
  for (const [index, sourceIdentity] of sourceIdentities.entries()) {
    if (!sourceIdentity) continue;
    const sourceNumId = String(sourceIdentity.numId);
    const sourceNum = sourceIndex.nums.get(sourceNumId);
    const sourceAbstractId =
      sourceIdentity.abstractNumId === undefined
        ? decimalWordAttribute(
            sourceNum ? directChild(sourceNum, 'abstractNumId') : undefined,
            'val',
          )
        : String(sourceIdentity.abstractNumId);
    const generatedAbstract = generatedConfigurations[index];
    const generatedAbstractId = generatedAbstract
      ? decimalWordAttribute(generatedAbstract, 'abstractNumId')
      : null;
    if (
      !sourceNum ||
      !sourceAbstractId ||
      !sourceIndex.abstracts.has(sourceAbstractId) ||
      !generatedAbstractId
    ) {
      continue;
    }
    addIdentityCandidate(
      abstractCandidates,
      sourceAbstractId,
      generatedAbstractId,
    );
    const matchingNums = generatedNums.filter(
      (num) =>
        decimalWordAttribute(directChild(num, 'abstractNumId'), 'val') ===
        generatedAbstractId,
    );
    if (matchingNums.length === 1) {
      const generatedNumId = decimalWordAttribute(matchingNums[0], 'numId');
      if (generatedNumId) {
        addIdentityCandidate(numCandidates, sourceNumId, generatedNumId);
      }
    }
  }
  return {
    abstracts: uniqueIdentityMappings(abstractCandidates),
    nums: uniqueIdentityMappings(numCandidates),
  };
}

function addIdentityCandidate(
  candidates: Map<string, Set<string>>,
  sourceId: string,
  generatedId: string,
): void {
  const generatedIds = candidates.get(sourceId);
  if (generatedIds) generatedIds.add(generatedId);
  else candidates.set(sourceId, new Set([generatedId]));
}

function uniqueIdentityMappings(
  candidates: ReadonlyMap<string, ReadonlySet<string>>,
): Map<string, string> {
  const result = new Map<string, string>();
  for (const [sourceId, generatedIds] of candidates) {
    if (generatedIds.size !== 1) continue;
    const generatedId = generatedIds.values().next().value;
    if (generatedId) result.set(sourceId, generatedId);
  }
  return result;
}

function numberingSemanticKey(
  element: Element,
  role: DocxExtensionDocumentRole,
  identityMap: NumberingIdentityMap,
): string {
  if (!DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '')) {
    return `{${element.namespaceURI ?? ''}}${element.localName}`;
  }
  if (element.localName === 'abstractNum') {
    return mappedIdentityKey(
      'abstractNum',
      decimalWordAttribute(element, 'abstractNumId'),
      role,
      identityMap.abstracts,
    );
  }
  if (element.localName === 'num') {
    return mappedIdentityKey(
      'num',
      decimalWordAttribute(element, 'numId'),
      role,
      identityMap.nums,
    );
  }
  if (element.localName === 'lvl' || element.localName === 'lvlOverride') {
    return `word:${element.localName}:${decimalWordAttribute(element, 'ilvl') ?? 'unidentified'}`;
  }
  return `word:${element.localName}`;
}

function mappedIdentityKey(
  kind: string,
  id: string | null,
  role: DocxExtensionDocumentRole,
  identities: ReadonlyMap<string, string>,
): string {
  if (!id) return `word:${kind}:unidentified`;
  if (role === 'generated') return `word:${kind}:${id}`;
  const generatedId = identities.get(id);
  return generatedId
    ? `word:${kind}:${generatedId}`
    : `word:${kind}:source-unmapped:${id}`;
}

function decimalWordAttribute(
  element: Element | null | undefined,
  name: string,
): string | null {
  if (!element) return null;
  const value = wordAttribute(element, name);
  if (!value || !/^\d{1,10}$/.test(value)) return null;
  const number = Number(value);
  return Number.isSafeInteger(number) && number <= 2_147_483_647
    ? String(number)
    : null;
}

function wordAttribute(element: Element, name: string): string | null {
  for (const item of Array.from(element.attributes)) {
    if (
      DOCX_WORDPROCESSING_NAMESPACES.has(
        xmlAttributeNamespace(element, item) ?? '',
      ) &&
      xmlAttributeLocalName(item) === name
    ) {
      return item.value.trim() || null;
    }
  }
  return null;
}
