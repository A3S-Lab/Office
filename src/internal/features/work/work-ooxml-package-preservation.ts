import JSZip from 'jszip';
import {
  attribute,
  directChildren,
  parseXml,
  resolvePartTarget,
} from './work-ooxml-package';

const CONTENT_TYPES_PATH = '[Content_Types].xml';
const CONTENT_TYPES_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/content-types';

interface OoxmlContentTypes {
  document: Document;
  defaults: Map<string, string>;
  overrides: Map<string, string>;
}

/**
 * Copies safe source-only DOCX parts into a newly generated package.
 *
 * Generated parts remain authoritative. Digital signatures are invalid after
 * an edit, and macro, ActiveX, and custom-ribbon parts are intentionally not
 * propagated into a macro-free DOCX export.
 */
export async function preserveDocxSourcePackage(
  generatedBuffer: ArrayBuffer,
  sourceBuffer: ArrayBuffer,
): Promise<ArrayBuffer> {
  const [generated, source] = await Promise.all([
    JSZip.loadAsync(generatedBuffer),
    JSZip.loadAsync(sourceBuffer),
  ]);
  const generatedTypes = await readContentTypes(generated, 'generated DOCX');
  const sourceTypes = await readContentTypes(source, 'source DOCX');
  assertWordprocessingDocument(generatedTypes, 'Generated OOXML package');
  assertWordprocessingDocument(sourceTypes, 'Registered source OOXML package');
  const generatedPaths = packagePaths(generated);
  const sourcePaths = packagePaths(source);
  const generatedByLower = pathLookup(generatedPaths);
  const ambiguousSourcePaths = ambiguousPaths(sourcePaths);
  validateSourcePackagePaths(sourcePaths, sourceTypes, ambiguousSourcePaths);
  const preservedPaths = new Set<string>();

  for (const path of sourcePaths) {
    const lowerPath = path.toLowerCase();
    const sourceType = contentTypeForPath(sourceTypes, path);
    if (
      path === CONTENT_TYPES_PATH ||
      isRelationshipsPart(path) ||
      isExcludedDocxPart(path) ||
      isExcludedDocxContentType(sourceType) ||
      generatedByLower.has(lowerPath) ||
      !sourceType
    ) {
      continue;
    }
    const entry = source.file(path);
    if (!entry) continue;
    generated.file(path, await entry.async('uint8array'));
    preservedPaths.add(path);
    generatedByLower.set(lowerPath, path);
  }

  const finalPartPaths = new Set(generatedByLower.keys());
  await preserveRelationships(
    generated,
    source,
    sourcePaths,
    finalPartPaths,
    generatedByLower,
  );
  preserveContentTypes(generatedTypes, sourceTypes, preservedPaths, generated);

  return generated.generateAsync({
    type: 'arraybuffer',
    compression: 'DEFLATE',
  });
}

async function preserveRelationships(
  generated: JSZip,
  source: JSZip,
  sourcePaths: readonly string[],
  finalPartPaths: ReadonlySet<string>,
  generatedByLower: Map<string, string>,
): Promise<void> {
  for (const sourcePath of sourcePaths.filter(isRelationshipsPart)) {
    if (!isSafeOpcPartPath(sourcePath) || isExcludedDocxPart(sourcePath)) {
      continue;
    }
    const ownerPart = relationshipOwnerPart(sourcePath);
    if (ownerPart && !hasPath(finalPartPaths, ownerPart)) continue;
    const sourceEntry = source.file(sourcePath);
    if (!sourceEntry) continue;
    const sourceDocument = parseXml(
      await sourceEntry.async('text'),
      `source DOCX ${sourcePath}`,
    );
    const generatedPath = generatedByLower.get(sourcePath.toLowerCase());
    if (!generatedPath) {
      removeUnsafeRelationships(sourceDocument, ownerPart, finalPartPaths);
      generated.file(
        sourcePath,
        new XMLSerializer().serializeToString(sourceDocument),
      );
      generatedByLower.set(sourcePath.toLowerCase(), sourcePath);
      continue;
    }

    const generatedEntry = generated.file(generatedPath);
    if (!generatedEntry) continue;
    const generatedDocument = parseXml(
      await generatedEntry.async('text'),
      `generated DOCX ${generatedPath}`,
    );
    mergeRelationships(
      generatedDocument,
      sourceDocument,
      ownerPart,
      finalPartPaths,
    );
    generated.file(
      generatedPath,
      new XMLSerializer().serializeToString(generatedDocument),
    );
  }
}

function removeUnsafeRelationships(
  document: Document,
  ownerPart: string,
  finalPartPaths: ReadonlySet<string>,
): void {
  for (const relationship of directChildren(
    document.documentElement,
    'Relationship',
  )) {
    if (
      !shouldPreserveRelationship(relationship, ownerPart, finalPartPaths, true)
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
): void {
  const root = generated.documentElement;
  const generatedItems = directChildren(root, 'Relationship');
  const usedIds = new Set(
    generatedItems.map((item) => attribute(item, 'Id') ?? '').filter(Boolean),
  );
  const existing = new Set(generatedItems.map(relationshipSignature));
  for (const sourceItem of directChildren(
    source.documentElement,
    'Relationship',
  )) {
    if (
      !shouldPreserveRelationship(sourceItem, ownerPart, finalPartPaths, false)
    ) {
      continue;
    }
    const signature = relationshipSignature(sourceItem);
    if (existing.has(signature)) continue;
    const imported = generated.importNode(sourceItem, true) as Element;
    const sourceId = attribute(imported, 'Id')?.trim() ?? '';
    const id =
      sourceId && !usedIds.has(sourceId)
        ? sourceId
        : nextRelationshipId(usedIds);
    imported.setAttribute('Id', id);
    usedIds.add(id);
    existing.add(signature);
    root.append(imported);
  }
}

function shouldPreserveRelationship(
  relationship: Element,
  ownerPart: string,
  finalPartPaths: ReadonlySet<string>,
  allowExternal: boolean,
): boolean {
  const type = attribute(relationship, 'Type')?.trim() ?? '';
  const target = attribute(relationship, 'Target')?.trim() ?? '';
  if (!type || !target || isExcludedRelationshipType(type)) return false;
  if (
    (attribute(relationship, 'TargetMode') ?? '').toLowerCase() === 'external'
  ) {
    return allowExternal;
  }
  const resolved = resolvePartTarget(ownerPart, target);
  return (
    isSafeOpcPartPath(resolved) &&
    !isExcludedDocxPart(resolved) &&
    hasPath(finalPartPaths, resolved)
  );
}

function preserveContentTypes(
  generated: OoxmlContentTypes,
  source: OoxmlContentTypes,
  preservedPaths: ReadonlySet<string>,
  archive: JSZip,
): void {
  const root = generated.document.documentElement;
  for (const path of preservedPaths) {
    const sourceType = contentTypeForPath(source, path);
    if (!sourceType || contentTypeForPath(generated, path) === sourceType) {
      continue;
    }
    const partName = `/${path}`;
    const existing = directChildren(root, 'Override').find(
      (item) =>
        normalizePartName(attribute(item, 'PartName') ?? '').toLowerCase() ===
        path.toLowerCase(),
    );
    if (existing) {
      existing.setAttribute('ContentType', sourceType);
    } else {
      const override = generated.document.createElementNS(
        CONTENT_TYPES_NAMESPACE,
        'Override',
      );
      override.setAttribute('PartName', partName);
      override.setAttribute('ContentType', sourceType);
      root.append(override);
    }
    generated.overrides.set(path.toLowerCase(), sourceType);
  }
  archive.file(
    CONTENT_TYPES_PATH,
    new XMLSerializer().serializeToString(generated.document),
  );
}

async function readContentTypes(
  archive: JSZip,
  label: string,
): Promise<OoxmlContentTypes> {
  const entry = archive.file(CONTENT_TYPES_PATH);
  if (!entry) throw new Error(`${label} content types part is missing.`);
  const document = parseXml(
    await entry.async('text'),
    `${label} ${CONTENT_TYPES_PATH}`,
  );
  return {
    document,
    defaults: new Map(
      directChildren(document.documentElement, 'Default').flatMap((item) => {
        const extension = attribute(item, 'Extension')?.toLowerCase();
        const contentType = attribute(item, 'ContentType');
        return extension && contentType ? [[extension, contentType]] : [];
      }),
    ),
    overrides: new Map(
      directChildren(document.documentElement, 'Override').flatMap((item) => {
        const partName = normalizePartName(
          attribute(item, 'PartName') ?? '',
        ).toLowerCase();
        const contentType = attribute(item, 'ContentType');
        return partName && contentType ? [[partName, contentType]] : [];
      }),
    ),
  };
}

function contentTypeForPath(
  types: OoxmlContentTypes,
  path: string,
): string | undefined {
  const normalized = normalizePartName(path).toLowerCase();
  const override = types.overrides.get(normalized);
  if (override) return override;
  const fileName = normalized.slice(normalized.lastIndexOf('/') + 1);
  const separator = fileName.lastIndexOf('.');
  return separator >= 0
    ? types.defaults.get(fileName.slice(separator + 1))
    : undefined;
}

function assertWordprocessingDocument(
  types: OoxmlContentTypes,
  label: string,
): void {
  const contentType = contentTypeForPath(types, 'word/document.xml')
    ?.trim()
    .toLowerCase();
  if (
    !contentType ||
    (!contentType.includes('wordprocessingml.document.main+xml') &&
      !contentType.includes('ms-word.document.macroenabled.main+xml'))
  ) {
    throw new Error(`${label} is not a WordprocessingML document.`);
  }
}

function packagePaths(archive: JSZip): string[] {
  return Object.entries(archive.files)
    .filter(([, entry]) => !entry.dir)
    .map(([path]) => path);
}

function pathLookup(paths: readonly string[]): Map<string, string> {
  return new Map(paths.map((path) => [path.toLowerCase(), path] as const));
}

function ambiguousPaths(paths: readonly string[]): Set<string> {
  const seen = new Set<string>();
  const ambiguous = new Set<string>();
  for (const path of paths) {
    const key = path.toLowerCase();
    if (seen.has(key)) ambiguous.add(key);
    seen.add(key);
  }
  return ambiguous;
}

function hasPath(paths: ReadonlySet<string>, expected: string): boolean {
  return paths.has(expected.toLowerCase());
}

function validateSourcePackagePaths(
  paths: readonly string[],
  types: OoxmlContentTypes,
  ambiguous: ReadonlySet<string>,
): void {
  if (ambiguous.size) {
    throw new Error(
      'Registered source OOXML package contains case-ambiguous part names.',
    );
  }
  for (const path of paths) {
    if (!isSafeOpcPartPath(path)) {
      throw new Error(
        `Registered source OOXML package contains an unsafe part name: ${path}`,
      );
    }
    if (
      path !== CONTENT_TYPES_PATH &&
      !isExcludedDocxPart(path) &&
      !contentTypeForPath(types, path)
    ) {
      throw new Error(
        `Registered source OOXML package part has no content type: ${path}`,
      );
    }
  }
}

function isSafeOpcPartPath(path: string): boolean {
  if (
    !path ||
    path.startsWith('/') ||
    path.includes('\\') ||
    /[\u0000-\u001f\u007f]/.test(path)
  ) {
    return false;
  }
  const segments = path.split('/');
  return segments.every(
    (segment) => Boolean(segment) && segment !== '.' && segment !== '..',
  );
}

function isRelationshipsPart(path: string): boolean {
  return (
    /^_rels\/\.rels$/i.test(path) || /(^|\/)_rels\/[^/]+\.rels$/i.test(path)
  );
}

function relationshipOwnerPart(path: string): string {
  if (/^_rels\/\.rels$/i.test(path)) return '';
  const match = /^(.*\/)?_rels\/([^/]+)\.rels$/i.exec(path);
  if (!match) return '';
  return `${match[1] ?? ''}${match[2]}`;
}

function normalizePartName(value: string): string {
  return value.replace(/^\/+/, '');
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

function isExcludedDocxPart(path: string): boolean {
  const normalized = path.toLowerCase();
  return (
    normalized.startsWith('_xmlsignatures/') ||
    normalized.startsWith('customui/') ||
    normalized.startsWith('word/activex/') ||
    /(^|\/)vbaproject(?:signature)?\.bin(?:\.rels)?$/.test(normalized) ||
    normalized === 'word/vbadata.xml'
  );
}

function isExcludedRelationshipType(type: string): boolean {
  const normalized = type.toLowerCase();
  return [
    'digital-signature',
    'vbaproject',
    'activex',
    'ui/extensibility',
  ].some((marker) => normalized.includes(marker));
}

function isExcludedDocxContentType(type: string | undefined): boolean {
  const normalized = type?.toLowerCase() ?? '';
  return [
    'digital-signature',
    'vbaproject',
    'activex',
    'customui',
    'custom-ui',
    'ribbon',
  ].some((marker) => normalized.includes(marker));
}
