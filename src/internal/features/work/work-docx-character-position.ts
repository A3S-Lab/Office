import {
  DOCUMENT_CHARACTER_POSITION_MAX_HALF_POINTS,
  DOCUMENT_CHARACTER_POSITION_MIN_HALF_POINTS,
  normalizeDocumentCharacterPositionHalfPoints,
} from './work-document-character-position';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import {
  parseDocxHalfPointsMeasure,
  STRICT_WORDPROCESSING_NAMESPACE,
} from './work-docx-twips';
import {
  XMLNS_NAMESPACE,
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { directChildren } from './work-ooxml-package';

export function docxCharacterPositionValue(value: unknown): string | undefined {
  const position = normalizeDocumentCharacterPositionHalfPoints(value);
  return position === null ? undefined : String(position);
}

export function docxCharacterPositionHalfPointsFromProperties(
  properties: Element | null | undefined,
): number | undefined {
  if (!properties) return undefined;
  const localMatches = directChildren(properties, 'position');
  const positions = localMatches.filter((element) =>
    DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? ''),
  );
  if (localMatches.length !== 1 || positions.length !== 1) return undefined;
  const position = positions[0];
  if (
    directChildren(position).length ||
    Array.from(position.childNodes).some(
      (node) =>
        (node.nodeType === Node.TEXT_NODE ||
          node.nodeType === Node.CDATA_SECTION_NODE) &&
        Boolean(node.textContent?.trim()),
    )
  ) {
    return undefined;
  }
  const attributes = Array.from(position.attributes).filter(
    (attribute) =>
      xmlAttributeNamespace(position, attribute) !== XMLNS_NAMESPACE &&
      attribute.name !== 'xmlns' &&
      !attribute.name.startsWith('xmlns:'),
  );
  if (
    attributes.length !== 1 ||
    xmlAttributeLocalName(attributes[0]) !== 'val' ||
    !DOCX_WORDPROCESSING_NAMESPACES.has(
      xmlAttributeNamespace(position, attributes[0]) ?? '',
    )
  ) {
    return undefined;
  }
  return (
    parseDocxHalfPointsMeasure(attributes[0].value.trim(), {
      minimum: DOCUMENT_CHARACTER_POSITION_MIN_HALF_POINTS,
      maximum: DOCUMENT_CHARACTER_POSITION_MAX_HALF_POINTS,
      signed: true,
      strict: position.namespaceURI === STRICT_WORDPROCESSING_NAMESPACE,
    }) ?? undefined
  );
}
