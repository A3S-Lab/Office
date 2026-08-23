import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import {
  XMLNS_NAMESPACE,
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { directChildren } from './work-ooxml-package';

export type DocxHiddenTextInspection =
  | { status: 'absent' }
  | { status: 'invalid' }
  | { status: 'valid'; value: boolean };

const ON_VALUES = new Set(['1', 'on', 'true']);
const OFF_VALUES = new Set(['0', 'off', 'false']);

export function inspectDocxHiddenText(
  properties: Element | null | undefined,
): DocxHiddenTextInspection {
  if (!properties) return { status: 'absent' };
  const localMatches = directChildren(properties, 'vanish');
  if (!localMatches.length) return { status: 'absent' };
  const nativeMatches = localMatches.filter((element) =>
    DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? ''),
  );
  if (localMatches.length !== 1 || nativeMatches.length !== 1) {
    return { status: 'invalid' };
  }
  const vanish = nativeMatches[0];
  if (
    directChildren(vanish).length ||
    Array.from(vanish.childNodes).some(
      (node) =>
        (node.nodeType === Node.TEXT_NODE ||
          node.nodeType === Node.CDATA_SECTION_NODE) &&
        Boolean(node.textContent?.trim()),
    )
  ) {
    return { status: 'invalid' };
  }
  const attributes = Array.from(vanish.attributes).filter(
    (attribute) =>
      xmlAttributeNamespace(vanish, attribute) !== XMLNS_NAMESPACE &&
      attribute.name !== 'xmlns' &&
      !attribute.name.startsWith('xmlns:'),
  );
  if (!attributes.length) return { status: 'valid', value: true };
  if (
    attributes.length !== 1 ||
    xmlAttributeLocalName(attributes[0]) !== 'val' ||
    !DOCX_WORDPROCESSING_NAMESPACES.has(
      xmlAttributeNamespace(vanish, attributes[0]) ?? '',
    )
  ) {
    return { status: 'invalid' };
  }
  const value = attributes[0].value;
  if (ON_VALUES.has(value)) return { status: 'valid', value: true };
  if (OFF_VALUES.has(value)) return { status: 'valid', value: false };
  return { status: 'invalid' };
}

export function docxHiddenTextFromProperties(
  properties: Element | null | undefined,
): boolean | undefined {
  const inspection = inspectDocxHiddenText(properties);
  return inspection.status === 'valid' ? inspection.value : undefined;
}

export function resolveDocxHiddenText(
  propertySources: readonly Element[],
): boolean | undefined {
  let hiddenText: boolean | undefined;
  for (const properties of propertySources) {
    const candidate = docxHiddenTextFromProperties(properties);
    if (candidate !== undefined) hiddenText = candidate;
  }
  return hiddenText;
}
