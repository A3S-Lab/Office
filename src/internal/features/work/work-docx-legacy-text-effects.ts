import {
  DOCUMENT_LEGACY_TEXT_EFFECT_NAMES,
  documentLegacyTextEffectsConflict,
  type WorkDocumentLegacyTextEffects,
} from './work-document-legacy-text-effects';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import {
  XMLNS_NAMESPACE,
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { directChildren } from './work-ooxml-package';

export type DocxLegacyTextEffectsInspection =
  | { status: 'absent' }
  | { status: 'invalid' }
  | { status: 'valid'; value: WorkDocumentLegacyTextEffects };

const ON_VALUES = new Set(['1', 'on', 'true']);
const OFF_VALUES = new Set(['0', 'off', 'false']);

export function inspectDocxLegacyTextEffects(
  properties: Element | null | undefined,
): DocxLegacyTextEffectsInspection {
  if (!properties) return { status: 'absent' };
  const effects: WorkDocumentLegacyTextEffects = {};
  let present = false;
  for (const name of DOCUMENT_LEGACY_TEXT_EFFECT_NAMES) {
    const localMatches = directChildren(properties, name);
    if (!localMatches.length) continue;
    present = true;
    const nativeMatches = localMatches.filter((element) =>
      DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? ''),
    );
    if (localMatches.length !== 1 || nativeMatches.length !== 1) {
      return { status: 'invalid' };
    }
    const value = inspectOnOffLeaf(nativeMatches[0]);
    if (value === null) return { status: 'invalid' };
    effects[name] = value;
  }
  if (!present) return { status: 'absent' };
  return documentLegacyTextEffectsConflict(effects)
    ? { status: 'invalid' }
    : { status: 'valid', value: effects };
}

export function docxLegacyTextEffectsFromProperties(
  properties: Element | null | undefined,
): WorkDocumentLegacyTextEffects | undefined {
  const inspection = inspectDocxLegacyTextEffects(properties);
  return inspection.status === 'valid' ? inspection.value : undefined;
}

export function resolveDocxLegacyTextEffects(
  propertySources: readonly Element[],
): WorkDocumentLegacyTextEffects | undefined {
  const effects: WorkDocumentLegacyTextEffects = {};
  let present = false;
  for (const properties of propertySources) {
    const inspection = inspectDocxLegacyTextEffects(properties);
    if (inspection.status === 'invalid') return undefined;
    if (inspection.status !== 'valid') continue;
    Object.assign(effects, inspection.value);
    present = true;
    if (documentLegacyTextEffectsConflict(effects)) return undefined;
  }
  return present ? effects : undefined;
}

function inspectOnOffLeaf(element: Element): boolean | null {
  if (
    directChildren(element).length ||
    Array.from(element.childNodes).some(
      (node) =>
        (node.nodeType === Node.TEXT_NODE ||
          node.nodeType === Node.CDATA_SECTION_NODE) &&
        Boolean(node.textContent?.trim()),
    )
  ) {
    return null;
  }
  const attributes = Array.from(element.attributes).filter(
    (attribute) =>
      xmlAttributeNamespace(element, attribute) !== XMLNS_NAMESPACE &&
      attribute.name !== 'xmlns' &&
      !attribute.name.startsWith('xmlns:'),
  );
  if (!attributes.length) return true;
  if (
    attributes.length !== 1 ||
    xmlAttributeLocalName(attributes[0]) !== 'val' ||
    !DOCX_WORDPROCESSING_NAMESPACES.has(
      xmlAttributeNamespace(element, attributes[0]) ?? '',
    )
  ) {
    return null;
  }
  const value = attributes[0].value;
  if (ON_VALUES.has(value)) return true;
  if (OFF_VALUES.has(value)) return false;
  return null;
}
