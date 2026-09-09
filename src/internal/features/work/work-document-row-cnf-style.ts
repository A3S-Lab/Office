import {
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { attribute } from './work-ooxml-package';

const CNF_STYLE_PATTERN = /^[01]{12}$/;

/**
 * ECMA ST_Cnf: exactly 12 binary digits describing table conditional formatting.
 */
export function normalizeDocumentRowCnfStyle(value: unknown): string | null {
  if (typeof value !== 'string') return null;
  const normalized = value.trim();
  if (!CNF_STYLE_PATTERN.test(normalized)) return null;
  return normalized;
}

export function parseDocxRowCnfStyleElement(element: Element): string | null {
  if (element.localName !== 'cnfStyle' || element.children.length > 0) {
    return null;
  }
  const namespace = element.namespaceURI;
  if (!namespace) return null;
  const attributes = Array.from(element.attributes).filter(
    (candidate) => xmlAttributeNamespace(element, candidate) === namespace,
  );
  if (attributes.length !== 1) return null;
  if (xmlAttributeLocalName(attributes[0]!) !== 'val') return null;
  return normalizeDocumentRowCnfStyle(attribute(element, 'val'));
}
