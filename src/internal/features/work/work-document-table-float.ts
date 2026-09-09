import {
  cloneXmlElement,
  declareInheritedNamespaces,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { directChild, parseXml } from './work-ooxml-package';

export const DOCUMENT_TABLE_FLOAT_OMML_ATTRIBUTE =
  'data-office-table-float-omml';
export const MAX_DOCUMENT_TABLE_FLOAT_OMML_LENGTH = 65_536;

const DOCX_RELATIONSHIP_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  'http://purl.oclc.org/ooxml/officeDocument/relationships',
  'http://schemas.openxmlformats.org/package/2006/relationships',
]);
const DOCX_WORDPROCESSING_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  'http://purl.oclc.org/ooxml/wordprocessingml/main',
]);

export function encodeDocumentTableFloatOmml(omml: string): string {
  if (!omml || omml.length > MAX_DOCUMENT_TABLE_FLOAT_OMML_LENGTH) return '';
  const bytes = new TextEncoder().encode(omml);
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
}

export function decodeDocumentTableFloatOmml(encoded: string): string | null {
  if (!encoded) return null;
  try {
    const binary = atob(encoded);
    const bytes = Uint8Array.from(binary, (character) =>
      character.charCodeAt(0),
    );
    const omml = new TextDecoder().decode(bytes).trim();
    if (!omml || omml.length > MAX_DOCUMENT_TABLE_FLOAT_OMML_LENGTH) {
      return null;
    }
    return omml;
  } catch {
    return null;
  }
}

export function documentTableFloatOmmlFromElement(
  element: Element,
): string | null {
  if (!(element instanceof HTMLElement)) return null;
  return decodeDocumentTableFloatOmml(
    element.dataset.officeTableFloatOmml ?? '',
  );
}

export function applyDocumentTableFloatOmmlToElement(
  element: HTMLElement,
  omml: string | null | undefined,
): void {
  const encoded = omml ? encodeDocumentTableFloatOmml(omml) : '';
  if (encoded) element.dataset.officeTableFloatOmml = encoded;
  else delete element.dataset.officeTableFloatOmml;
}

export function docxTableFloatHasRelationshipBindings(
  element: Element,
): boolean {
  for (const attribute of Array.from(element.attributes)) {
    const namespace =
      attribute.namespaceURI || xmlAttributeNamespace(element, attribute) || '';
    if (DOCX_RELATIONSHIP_NAMESPACES.has(namespace)) return true;
  }
  return false;
}

export function serializePreservableDocxTableFloat(
  tableProperties: Element | null | undefined,
): string | null {
  const float = tableProperties ? directChild(tableProperties, 'tblpPr') : null;
  if (!float) return null;
  const namespace = float.namespaceURI ?? '';
  if (!DOCX_WORDPROCESSING_NAMESPACES.has(namespace)) return null;
  if (docxTableFloatHasRelationshipBindings(float)) return null;
  if (float.children.length > 0) return null;
  const prefix = float.prefix || 'w';
  const shell = parseXml(
    `<${prefix}:${float.localName} xmlns:${prefix}="${namespace}"/>`,
    'opaque table float shell',
  );
  const clone = cloneXmlElement(shell, float);
  declareInheritedNamespaces(clone, float);
  shell.replaceChild(clone, shell.documentElement);
  const omml = new XMLSerializer().serializeToString(clone).trim();
  if (!omml || omml.length > MAX_DOCUMENT_TABLE_FLOAT_OMML_LENGTH) return null;
  return omml;
}

export function importDocxTableFloatElement(
  document: Document,
  omml: string,
): Element | null {
  try {
    const parsed = parseXml(omml, 'preserved table float');
    const root = parsed.documentElement;
    if (
      root.localName !== 'tblpPr' ||
      !DOCX_WORDPROCESSING_NAMESPACES.has(root.namespaceURI ?? '')
    ) {
      return null;
    }
    if (docxTableFloatHasRelationshipBindings(root)) return null;
    const imported = cloneXmlElement(document, root);
    declareInheritedNamespaces(imported, root);
    return imported;
  } catch {
    return null;
  }
}
