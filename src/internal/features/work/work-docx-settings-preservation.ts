import type JSZip from 'jszip';
import {
  DOCX_WORDPROCESSING_NAMESPACES,
  mergeDocxIgnorableExtensions,
} from './work-docx-ignorable-extension-preservation';
import { assertXmlRoot } from './work-docx-settings-xml';
import { parseXml } from './work-ooxml-package';
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
 * and structurally valid, non-conflicting mc:AlternateContent blocks.
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
  generated.file(generatedPath, serializeUtf8Xml(generatedDocument));
}
