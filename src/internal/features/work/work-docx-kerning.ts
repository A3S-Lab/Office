import JSZip from 'jszip';
import {
  DOCUMENT_KERNING_THRESHOLD_MAX_HALF_POINTS,
  DOCUMENT_KERNING_THRESHOLD_MIN_HALF_POINTS,
  normalizeDocumentKerningThresholdHalfPoints,
} from './work-document-kerning';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import {
  XMLNS_NAMESPACE,
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { parseBoundedDocxInteger } from './work-docx-twips';
import { descendants, directChildren, parseXml } from './work-ooxml-package';
import { decodeXmlBytes, serializeUtf8Xml } from './work-ooxml-xml';

export const DOCX_EXPLICIT_ZERO_KERNING_THRESHOLD_SENTINEL =
  DOCUMENT_KERNING_THRESHOLD_MAX_HALF_POINTS + 1;

const KERNING_PART_PATTERN =
  /^word\/(?:document|header\d*|footer\d*|footnotes|endnotes)\.xml$/i;

export type DocxKerningThresholdInspection =
  | { status: 'absent' }
  | { status: 'invalid' }
  | { status: 'valid'; value: number };

export function docxKerningThresholdValue(value: unknown): number | undefined {
  const threshold = normalizeDocumentKerningThresholdHalfPoints(value);
  if (threshold === null) return undefined;
  return threshold === 0
    ? DOCX_EXPLICIT_ZERO_KERNING_THRESHOLD_SENTINEL
    : threshold;
}

export function docxKerningThresholdHalfPointsFromProperties(
  properties: Element | null | undefined,
): number | undefined {
  const inspection = inspectDocxKerningThresholdHalfPoints(properties);
  return inspection.status === 'valid' ? inspection.value : undefined;
}

export function resolveDocxKerningThresholdHalfPoints(
  propertySources: readonly Element[],
): number | undefined {
  let threshold: number | undefined;
  for (const properties of propertySources) {
    const candidate = docxKerningThresholdHalfPointsFromProperties(properties);
    if (candidate !== undefined) threshold = candidate;
  }
  return threshold;
}

export function inspectDocxKerningThresholdHalfPoints(
  properties: Element | null | undefined,
): DocxKerningThresholdInspection {
  if (!properties) return { status: 'absent' };
  const localMatches = directChildren(properties, 'kern');
  if (!localMatches.length) return { status: 'absent' };
  const nativeMatches = localMatches.filter((element) =>
    DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? ''),
  );
  if (localMatches.length !== 1 || nativeMatches.length !== 1) {
    return { status: 'invalid' };
  }
  const kerning = nativeMatches[0];
  if (
    directChildren(kerning).length ||
    Array.from(kerning.childNodes).some(
      (node) =>
        (node.nodeType === Node.TEXT_NODE ||
          node.nodeType === Node.CDATA_SECTION_NODE) &&
        Boolean(node.textContent?.trim()),
    )
  ) {
    return { status: 'invalid' };
  }
  const attributes = Array.from(kerning.attributes).filter(
    (attribute) =>
      xmlAttributeNamespace(kerning, attribute) !== XMLNS_NAMESPACE &&
      attribute.name !== 'xmlns' &&
      !attribute.name.startsWith('xmlns:'),
  );
  if (
    attributes.length !== 1 ||
    xmlAttributeLocalName(attributes[0]) !== 'val' ||
    !DOCX_WORDPROCESSING_NAMESPACES.has(
      xmlAttributeNamespace(kerning, attributes[0]) ?? '',
    )
  ) {
    return { status: 'invalid' };
  }
  const value = parseBoundedDocxInteger(attributes[0].value.trim(), {
    minimum: DOCUMENT_KERNING_THRESHOLD_MIN_HALF_POINTS,
    maximum: DOCUMENT_KERNING_THRESHOLD_MAX_HALF_POINTS,
    signed: false,
  });
  return value === null ? { status: 'invalid' } : { status: 'valid', value };
}

export async function patchDocxExplicitZeroKerningThresholds(
  buffer: ArrayBuffer,
): Promise<ArrayBuffer> {
  const archive = await JSZip.loadAsync(buffer);
  let replacements = 0;
  for (const entry of Object.values(archive.files)) {
    if (entry.dir || !KERNING_PART_PATTERN.test(entry.name)) continue;
    const document = parseXml(
      decodeXmlBytes(
        await entry.async('uint8array'),
        `generated DOCX ${entry.name}`,
      ),
      `generated DOCX ${entry.name}`,
    );
    let changed = false;
    for (const kerning of descendants(document, 'kern')) {
      if (
        !DOCX_WORDPROCESSING_NAMESPACES.has(kerning.namespaceURI ?? '') ||
        wordValue(kerning) !==
          String(DOCX_EXPLICIT_ZERO_KERNING_THRESHOLD_SENTINEL)
      ) {
        continue;
      }
      setWordValue(kerning, '0');
      replacements += 1;
      changed = true;
    }
    if (changed) archive.file(entry.name, serializeUtf8Xml(document));
  }
  if (!replacements) {
    throw new Error('Generated DOCX explicit-zero kerning was not emitted.');
  }
  return archive.generateAsync({ type: 'arraybuffer' });
}

function wordValue(element: Element): string | null {
  const values = Array.from(element.attributes).filter(
    (attribute) =>
      xmlAttributeLocalName(attribute) === 'val' &&
      DOCX_WORDPROCESSING_NAMESPACES.has(
        xmlAttributeNamespace(element, attribute) ?? '',
      ),
  );
  return values.length === 1 ? values[0].value : null;
}

function setWordValue(element: Element, value: string): void {
  const existing = Array.from(element.attributes).find(
    (attribute) =>
      xmlAttributeLocalName(attribute) === 'val' &&
      DOCX_WORDPROCESSING_NAMESPACES.has(
        xmlAttributeNamespace(element, attribute) ?? '',
      ),
  );
  const namespace = element.namespaceURI ?? '';
  const prefix = element.prefix || 'w';
  if (existing) {
    existing.value = value;
  } else {
    element.setAttributeNS(namespace, `${prefix}:val`, value);
  }
}
