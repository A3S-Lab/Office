import JSZip from 'jszip';
import { preserveDocxFontTable } from './work-docx-font-table-preservation';
import {
  preserveDocxNumberingExtensions,
  type DocxSourceNumberingIdentity,
} from './work-docx-numbering-extension-preservation';
import { preserveDocxSettingsExtensions } from './work-docx-settings-preservation';
import { preserveDocxStyleExtensions } from './work-docx-style-extension-preservation';
import { attribute, directChildren, parseXml } from './work-ooxml-package';
import {
  isExcludedDocxContentType,
  isExcludedDocxPart,
  isSafeOpcPartPath,
} from './work-ooxml-package-security';
import {
  isRelationshipsPart,
  preserveDocxRelationships,
} from './work-ooxml-relationship-preservation';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';

const CONTENT_TYPES_PATH = '[Content_Types].xml';
const CONTENT_TYPES_NAMESPACE =
  'http://schemas.openxmlformats.org/package/2006/content-types';

interface OoxmlContentTypes {
  document: Document;
  defaults: Map<string, string>;
  overrides: Map<string, string>;
}

export interface DocxSourcePackagePreservationOptions {
  numberingIdentities?: readonly (DocxSourceNumberingIdentity | null)[];
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
  options: DocxSourcePackagePreservationOptions = {},
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
  const generatedPartPaths = new Set(generatedByLower.keys());
  const sourceByLower = pathLookup(sourcePaths);
  const ambiguousSourcePaths = ambiguousPaths(sourcePaths);
  validateSourcePackagePaths(sourcePaths, sourceTypes, ambiguousSourcePaths);
  const generatedSettingsPath = generatedByLower.get('word/settings.xml');
  const sourceSettingsPath = sourceByLower.get('word/settings.xml');
  if (
    generatedSettingsPath &&
    sourceSettingsPath &&
    isDocxSettingsContentType(
      contentTypeForPath(generatedTypes, generatedSettingsPath),
    ) &&
    isDocxSettingsContentType(
      contentTypeForPath(sourceTypes, sourceSettingsPath),
    )
  ) {
    await preserveDocxSettingsExtensions(
      generated,
      source,
      generatedSettingsPath,
      sourceSettingsPath,
    );
  }
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
  const relationshipReferences = await preserveDocxRelationships(
    generated,
    source,
    sourcePaths,
    finalPartPaths,
    generatedPartPaths,
    generatedByLower,
  );
  const generatedFontTablePath = generatedByLower.get('word/fonttable.xml');
  const sourceFontTablePath = sourceByLower.get('word/fonttable.xml');
  if (
    generatedFontTablePath &&
    sourceFontTablePath &&
    isDocxFontTableContentType(
      contentTypeForPath(generatedTypes, generatedFontTablePath),
    ) &&
    isDocxFontTableContentType(
      contentTypeForPath(sourceTypes, sourceFontTablePath),
    )
  ) {
    const fontPartPaths = new Set(
      Array.from(preservedPaths)
        .filter((path) =>
          isObfuscatedFontContentType(contentTypeForPath(sourceTypes, path)),
        )
        .map((path) => path.toLowerCase()),
    );
    await preserveDocxFontTable(
      generated,
      source,
      relationshipReferences.get('word/fonttable.xml') ?? new Map(),
      fontPartPaths,
      generatedFontTablePath,
      sourceFontTablePath,
    );
  }
  const generatedStylesPath = generatedByLower.get('word/styles.xml');
  const sourceStylesPath = sourceByLower.get('word/styles.xml');
  if (
    generatedStylesPath &&
    sourceStylesPath &&
    isDocxStylesContentType(
      contentTypeForPath(generatedTypes, generatedStylesPath),
    ) &&
    isDocxStylesContentType(contentTypeForPath(sourceTypes, sourceStylesPath))
  ) {
    await preserveDocxStyleExtensions(
      generated,
      source,
      generatedStylesPath,
      sourceStylesPath,
    );
  }
  const generatedNumberingPath = generatedByLower.get('word/numbering.xml');
  const sourceNumberingPath = sourceByLower.get('word/numbering.xml');
  if (
    generatedNumberingPath &&
    sourceNumberingPath &&
    isDocxNumberingContentType(
      contentTypeForPath(generatedTypes, generatedNumberingPath),
    ) &&
    isDocxNumberingContentType(
      contentTypeForPath(sourceTypes, sourceNumberingPath),
    )
  ) {
    await preserveDocxNumberingExtensions(
      generated,
      source,
      options.numberingIdentities ?? [],
      generatedNumberingPath,
      sourceNumberingPath,
    );
  }
  preserveContentTypes(generatedTypes, sourceTypes, preservedPaths, generated);

  return generated.generateAsync({
    type: 'arraybuffer',
    compression: 'DEFLATE',
  });
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
  archive.file(CONTENT_TYPES_PATH, serializeUtf8Xml(generated.document));
}

async function readContentTypes(
  archive: JSZip,
  label: string,
): Promise<OoxmlContentTypes> {
  const entry = archive.file(CONTENT_TYPES_PATH);
  if (!entry) throw new Error(`${label} content types part is missing.`);
  const document = parseXml(
    decodeXmlBytes(
      await entry.async('uint8array'),
      `${label} ${CONTENT_TYPES_PATH}`,
    ),
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

function normalizePartName(value: string): string {
  return value.replace(/^\/+/, '');
}

function isDocxSettingsContentType(type: string | undefined): boolean {
  return (
    type?.trim().toLowerCase() ===
    'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'
  );
}

function isDocxFontTableContentType(type: string | undefined): boolean {
  return (
    type?.trim().toLowerCase() ===
    'application/vnd.openxmlformats-officedocument.wordprocessingml.fonttable+xml'
  );
}

function isDocxStylesContentType(type: string | undefined): boolean {
  return (
    type?.trim().toLowerCase() ===
    'application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml'
  );
}

function isDocxNumberingContentType(type: string | undefined): boolean {
  return (
    type?.trim().toLowerCase() ===
    'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml'
  );
}

function isObfuscatedFontContentType(type: string | undefined): boolean {
  return (
    type?.trim().toLowerCase() ===
    'application/vnd.openxmlformats-officedocument.obfuscatedfont'
  );
}
