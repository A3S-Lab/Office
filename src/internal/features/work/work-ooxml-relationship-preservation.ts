import type JSZip from 'jszip';
import {
  attribute,
  directChildren,
  parseXml,
  resolvePartTarget,
} from './work-ooxml-package';
import {
  isExcludedDocxPart,
  isExcludedRelationshipType,
  isRegeneratedDocxNoteCommentPart,
  isSafeOpcPartPath,
} from './work-ooxml-package-security';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';

export interface PreservedRelationshipReference {
  id: string;
  target: string;
  targetMode?: string;
  type: string;
}

export type PreservedRelationshipReferences = Map<
  string,
  Map<string, PreservedRelationshipReference>
>;

export async function preserveDocxRelationships(
  generated: JSZip,
  source: JSZip,
  sourcePaths: readonly string[],
  finalPartPaths: ReadonlySet<string>,
  generatedPartPaths: ReadonlySet<string>,
  generatedByLower: Map<string, string>,
): Promise<PreservedRelationshipReferences> {
  const result: PreservedRelationshipReferences = new Map();
  for (const sourcePath of sourcePaths.filter(isRelationshipsPart)) {
    if (!isSafeOpcPartPath(sourcePath) || isExcludedDocxPart(sourcePath)) {
      continue;
    }
    const ownerPart = relationshipOwnerPart(sourcePath);
    if (ownerPart && !hasPath(finalPartPaths, ownerPart)) continue;
    if (isRegeneratedDocxNoteCommentPart(ownerPart)) continue;
    if (ownerPart.toLowerCase() === 'word/settings.xml') continue;
    const sourceEntry = source.file(sourcePath);
    if (!sourceEntry) continue;
    const sourceDocument = parseXml(
      decodeXmlBytes(
        await sourceEntry.async('uint8array'),
        `source DOCX ${sourcePath}`,
      ),
      `source DOCX ${sourcePath}`,
    );
    assertUniqueRelationshipIds(sourceDocument, sourcePath);
    const generatedPath = generatedByLower.get(sourcePath.toLowerCase());
    if (!generatedPath) {
      removeUnsafeRelationships(
        sourceDocument,
        ownerPart,
        finalPartPaths,
        Boolean(ownerPart) && !hasPath(generatedPartPaths, ownerPart),
      );
      const references = referencesFromDocument(sourceDocument);
      if (references.size) result.set(ownerPart.toLowerCase(), references);
      generated.file(sourcePath, serializeUtf8Xml(sourceDocument));
      generatedByLower.set(sourcePath.toLowerCase(), sourcePath);
      continue;
    }

    const generatedEntry = generated.file(generatedPath);
    if (!generatedEntry) continue;
    const generatedDocument = parseXml(
      decodeXmlBytes(
        await generatedEntry.async('uint8array'),
        `generated DOCX ${generatedPath}`,
      ),
      `generated DOCX ${generatedPath}`,
    );
    const references = mergeRelationships(
      generatedDocument,
      sourceDocument,
      ownerPart,
      finalPartPaths,
    );
    if (references.size) result.set(ownerPart.toLowerCase(), references);
    generated.file(generatedPath, serializeUtf8Xml(generatedDocument));
  }
  return result;
}

export function isRelationshipsPart(path: string): boolean {
  return (
    /^_rels\/\.rels$/i.test(path) || /(^|\/)_rels\/[^/]+\.rels$/i.test(path)
  );
}

function removeUnsafeRelationships(
  document: Document,
  ownerPart: string,
  finalPartPaths: ReadonlySet<string>,
  allowExternal: boolean,
): void {
  for (const relationship of directChildren(
    document.documentElement,
    'Relationship',
  )) {
    if (
      !shouldPreserveRelationship(
        relationship,
        ownerPart,
        finalPartPaths,
        allowExternal,
      )
    ) {
      relationship.remove();
    }
  }
}

function mergeRelationships(
  generated: Document,
  source: Document,
  ownerPart: string,
  finalPartPaths: ReadonlySet<string>,
): Map<string, PreservedRelationshipReference> {
  const root = generated.documentElement;
  const generatedItems = directChildren(root, 'Relationship');
  const usedIds = new Set(
    generatedItems.map((item) => attribute(item, 'Id') ?? '').filter(Boolean),
  );
  const bySignature = new Map(
    generatedItems.map((item) => [relationshipSignature(item), item] as const),
  );
  const references = new Map<string, PreservedRelationshipReference>();
  for (const sourceItem of directChildren(
    source.documentElement,
    'Relationship',
  )) {
    if (
      !shouldPreserveRelationship(sourceItem, ownerPart, finalPartPaths, false)
    ) {
      continue;
    }
    const sourceId = attribute(sourceItem, 'Id')?.trim() ?? '';
    const signature = relationshipSignature(sourceItem);
    let finalItem = bySignature.get(signature);
    if (!finalItem) {
      finalItem = generated.importNode(sourceItem, true) as Element;
      const id = usedIds.has(sourceId) ? nextRelationshipId(usedIds) : sourceId;
      finalItem.setAttribute('Id', id);
      usedIds.add(id);
      bySignature.set(signature, finalItem);
      root.append(finalItem);
    }
    references.set(sourceId, relationshipReference(finalItem));
  }
  return references;
}

function shouldPreserveRelationship(
  relationship: Element,
  ownerPart: string,
  finalPartPaths: ReadonlySet<string>,
  allowExternal: boolean,
): boolean {
  const type = attribute(relationship, 'Type')?.trim() ?? '';
  const target = attribute(relationship, 'Target')?.trim() ?? '';
  const targetMode = (attribute(relationship, 'TargetMode') ?? '')
    .trim()
    .toLowerCase();
  if (!type || !target || isExcludedRelationshipType(type)) return false;
  if (targetMode && targetMode !== 'internal' && targetMode !== 'external')
    return false;
  if (targetMode === 'external') return allowExternal;
  const resolved = resolvePartTarget(ownerPart, target);
  return (
    isSafeOpcPartPath(resolved) &&
    !isExcludedDocxPart(resolved) &&
    hasPath(finalPartPaths, resolved)
  );
}

function assertUniqueRelationshipIds(
  document: Document,
  sourcePath: string,
): void {
  const used = new Set<string>();
  for (const item of directChildren(document.documentElement, 'Relationship')) {
    const id = attribute(item, 'Id')?.trim() ?? '';
    if (!id) {
      throw new Error(
        `Registered source ${sourcePath} has a missing relationship ID.`,
      );
    }
    if (used.has(id)) {
      throw new Error(
        `Registered source ${sourcePath} has duplicate relationship IDs.`,
      );
    }
    used.add(id);
  }
}

function referencesFromDocument(
  document: Document,
): Map<string, PreservedRelationshipReference> {
  return new Map(
    directChildren(document.documentElement, 'Relationship').map((item) => {
      const reference = relationshipReference(item);
      return [reference.id, reference] as const;
    }),
  );
}

function relationshipReference(
  relationship: Element,
): PreservedRelationshipReference {
  return {
    id: attribute(relationship, 'Id')?.trim() ?? '',
    target: attribute(relationship, 'Target')?.trim() ?? '',
    targetMode: attribute(relationship, 'TargetMode')?.trim() || undefined,
    type: attribute(relationship, 'Type')?.trim() ?? '',
  };
}

function relationshipOwnerPart(path: string): string {
  if (/^_rels\/\.rels$/i.test(path)) return '';
  const match = /^(.*\/)?_rels\/([^/]+)\.rels$/i.exec(path);
  if (!match) return '';
  return `${match[1] ?? ''}${match[2]}`;
}

function relationshipSignature(relationship: Element): string {
  return [
    attribute(relationship, 'Type') ?? '',
    attribute(relationship, 'Target') ?? '',
    (attribute(relationship, 'TargetMode') ?? '').toLowerCase(),
  ].join('\u0000');
}

function nextRelationshipId(used: ReadonlySet<string>): string {
  let index = 1;
  while (used.has(`rId${index}`)) index += 1;
  return `rId${index}`;
}

function hasPath(paths: ReadonlySet<string>, expected: string): boolean {
  return paths.has(expected.toLowerCase());
}
