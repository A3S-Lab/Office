import type JSZip from 'jszip';
import {
  DOCX_WORDPROCESSING_NAMESPACES,
  mergeDocxIgnorableExtensions,
} from './work-docx-ignorable-extension-preservation';
import { assertXmlRoot } from './work-docx-settings-xml';
import {
  DOCX_COLOR_SCHEME_MAPPING_ATTRIBUTES,
  parseDocxColorSchemeMappingElement,
} from './work-docx-theme';
import {
  directChildren,
  parseXml,
  xmlNamespacePrefix,
} from './work-ooxml-package';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';

const SETTINGS_PATH = 'word/settings.xml';
const SOURCE_SETTING_DENYLIST = new Set([
  'attachedTemplate',
  'documentProtection',
  'mailMerge',
  'trackRevisions',
  'updateFields',
  'writeProtection',
]);

/**
 * Merges passive extension markup from source settings into generated settings.
 *
 * WordprocessingML settings and relationship-bound markup remain generated-only.
 * The source contribution is limited to ignorable extension attributes/elements
 * and structurally valid, non-conflicting mc:AlternateContent blocks, plus a
 * strictly validated color-scheme mapping required by preserved theme colors.
 */
export async function preserveDocxSettingsExtensions(
  generated: JSZip,
  source: JSZip,
  generatedPath = SETTINGS_PATH,
  sourcePath = SETTINGS_PATH,
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
    'settings',
    DOCX_WORDPROCESSING_NAMESPACES,
    'Generated DOCX word/settings.xml is not WordprocessingML settings.',
  );
  assertXmlRoot(
    sourceDocument.documentElement,
    'settings',
    DOCX_WORDPROCESSING_NAMESPACES,
    'Registered source DOCX word/settings.xml is not WordprocessingML settings.',
  );

  mergeDocxIgnorableExtensions(generatedDocument, sourceDocument, {
    allowSemanticElement: (element, generatedSemanticNames) =>
      !SOURCE_SETTING_DENYLIST.has(element.localName) &&
      !generatedSemanticNames.has(element.localName),
  });
  preserveColorSchemeMapping(generatedDocument, sourceDocument);
  generated.file(generatedPath, serializeUtf8Xml(generatedDocument));
}

function preserveColorSchemeMapping(
  generated: Document,
  source: Document,
): void {
  const sourceCandidates = directChildren(
    source.documentElement,
    'clrSchemeMapping',
  ).filter(
    (element) => element.namespaceURI === source.documentElement.namespaceURI,
  );
  if (sourceCandidates.length !== 1) return;
  const mapping = parseDocxColorSchemeMappingElement(sourceCandidates[0]);
  if (!mapping) return;

  const generatedCandidates = directChildren(
    generated.documentElement,
    'clrSchemeMapping',
  ).filter(
    (element) =>
      element.namespaceURI === generated.documentElement.namespaceURI,
  );
  if (generatedCandidates.length > 1) return;

  const namespace = generated.documentElement.namespaceURI;
  if (!namespace || !DOCX_WORDPROCESSING_NAMESPACES.has(namespace)) return;
  const prefix =
    xmlNamespacePrefix(generated.documentElement, namespace) ?? 'w';
  const preserved = generated.createElementNS(
    namespace,
    `${prefix}:clrSchemeMapping`,
  );
  for (const [
    semantic,
    attributeName,
  ] of DOCX_COLOR_SCHEME_MAPPING_ATTRIBUTES) {
    const value = mapping.get(semantic);
    if (value)
      preserved.setAttributeNS(namespace, `${prefix}:${attributeName}`, value);
  }

  const existing = generatedCandidates[0];
  if (existing) {
    existing.replaceWith(preserved);
    return;
  }
  const sourceChildren = directChildren(source.documentElement);
  const sourceIndex = sourceChildren.indexOf(sourceCandidates[0]);
  const generatedChildren = directChildren(generated.documentElement);
  const anchor = sourceChildren
    .slice(sourceIndex + 1)
    .reduce<Element | null>((match, sourceChild) => {
      if (match) return match;
      const candidates = generatedChildren.filter(
        (generatedChild) =>
          generatedChild.localName === sourceChild.localName &&
          DOCX_WORDPROCESSING_NAMESPACES.has(generatedChild.namespaceURI ?? ''),
      );
      return candidates.length === 1 ? candidates[0] : null;
    }, null);
  generated.documentElement.insertBefore(preserved, anchor);
}
