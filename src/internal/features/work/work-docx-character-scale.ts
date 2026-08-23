import {
  DOCUMENT_CHARACTER_SCALE_DEFAULT_PERCENT,
  DOCUMENT_CHARACTER_SCALE_MAX_PERCENT,
  DOCUMENT_CHARACTER_SCALE_MIN_PERCENT,
  normalizeDocumentCharacterScalePercent,
} from './work-document-character-scale';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import {
  XMLNS_NAMESPACE,
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { parseBoundedDocxInteger } from './work-docx-twips';
import { directChildren } from './work-ooxml-package';

export function docxCharacterScaleValue(value: unknown): number | undefined {
  return normalizeDocumentCharacterScalePercent(value) ?? undefined;
}

export function docxCharacterScalePercentFromProperties(
  properties: Element | null | undefined,
): number | undefined {
  if (!properties) return undefined;
  const localMatches = directChildren(properties, 'w');
  const scales = localMatches.filter((element) =>
    DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? ''),
  );
  if (localMatches.length !== 1 || scales.length !== 1) return undefined;
  const scale = scales[0];
  if (
    directChildren(scale).length ||
    Array.from(scale.childNodes).some(
      (node) =>
        (node.nodeType === Node.TEXT_NODE ||
          node.nodeType === Node.CDATA_SECTION_NODE) &&
        Boolean(node.textContent?.trim()),
    )
  ) {
    return undefined;
  }
  const attributes = Array.from(scale.attributes).filter(
    (attribute) =>
      xmlAttributeNamespace(scale, attribute) !== XMLNS_NAMESPACE &&
      attribute.name !== 'xmlns' &&
      !attribute.name.startsWith('xmlns:'),
  );
  if (!attributes.length) return DOCUMENT_CHARACTER_SCALE_DEFAULT_PERCENT;
  if (
    attributes.length !== 1 ||
    xmlAttributeLocalName(attributes[0]) !== 'val' ||
    !DOCX_WORDPROCESSING_NAMESPACES.has(
      xmlAttributeNamespace(scale, attributes[0]) ?? '',
    )
  ) {
    return undefined;
  }
  return (
    parseBoundedDocxInteger(attributes[0].value.trim(), {
      minimum: DOCUMENT_CHARACTER_SCALE_MIN_PERCENT,
      maximum: DOCUMENT_CHARACTER_SCALE_MAX_PERCENT,
      signed: false,
    }) ?? undefined
  );
}
