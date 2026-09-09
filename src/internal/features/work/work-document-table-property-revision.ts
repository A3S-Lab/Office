import {
  cloneXmlElement,
  declareInheritedNamespaces,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { isSupportedDocxTableFormattingChange } from './work-docx-table-format-change-import';
import { isSupportedDocxRowFormattingChange } from './work-docx-row-format-change-import';
import { isSupportedDocxCellFormattingChange } from './work-docx-cell-format-change-import';
import { isSupportedDocxSectionFormattingChange } from './work-docx-section-format-change-import';
import { directChild, parseXml } from './work-ooxml-package';

export type DocxTableScopedPropertyRevisionKind =
  | 'tblPrChange'
  | 'trPrChange'
  | 'tcPrChange'
  | 'sectPrChange';

export const DOCUMENT_TABLE_PROPERTY_REVISION_OMML_ATTRIBUTE =
  'data-office-table-property-revision-omml';
export const DOCUMENT_ROW_PROPERTY_REVISION_OMML_ATTRIBUTE =
  'data-office-row-property-revision-omml';
export const DOCUMENT_CELL_PROPERTY_REVISION_OMML_ATTRIBUTE =
  'data-office-cell-property-revision-omml';
export const MAX_DOCUMENT_TABLE_PROPERTY_REVISION_OMML_LENGTH = 65_536;

const DOCX_RELATIONSHIP_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  'http://purl.oclc.org/ooxml/officeDocument/relationships',
  'http://schemas.openxmlformats.org/package/2006/relationships',
]);
const DOCX_WORDPROCESSING_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  'http://purl.oclc.org/ooxml/wordprocessingml/main',
]);

/**
 * Encodes a relationship-free table/row/cell property-revision snapshot so
 * untouched structures can round-trip without inventing review UI.
 */
export function encodeDocumentTablePropertyRevisionOmml(omml: string): string {
  if (!omml || omml.length > MAX_DOCUMENT_TABLE_PROPERTY_REVISION_OMML_LENGTH) {
    return '';
  }
  const bytes = new TextEncoder().encode(omml);
  let binary = '';
  for (const byte of bytes) binary += String.fromCharCode(byte);
  return btoa(binary);
}

export function decodeDocumentTablePropertyRevisionOmml(
  encoded: string,
): string | null {
  if (!encoded) return null;
  try {
    const binary = atob(encoded);
    const bytes = Uint8Array.from(binary, (character) =>
      character.charCodeAt(0),
    );
    const omml = new TextDecoder().decode(bytes).trim();
    if (
      !omml ||
      omml.length > MAX_DOCUMENT_TABLE_PROPERTY_REVISION_OMML_LENGTH
    ) {
      return null;
    }
    return omml;
  } catch {
    return null;
  }
}

export function documentTablePropertyRevisionOmmlFromElement(
  element: Element,
): string | null {
  if (!(element instanceof HTMLElement)) return null;
  return decodeDocumentTablePropertyRevisionOmml(
    element.dataset.officeTablePropertyRevisionOmml ?? '',
  );
}

export function documentRowPropertyRevisionOmmlFromElement(
  element: Element,
): string | null {
  if (!(element instanceof HTMLElement)) return null;
  return decodeDocumentTablePropertyRevisionOmml(
    element.dataset.officeRowPropertyRevisionOmml ?? '',
  );
}

export function documentCellPropertyRevisionOmmlFromElement(
  element: Element,
): string | null {
  if (!(element instanceof HTMLElement)) return null;
  return decodeDocumentTablePropertyRevisionOmml(
    element.dataset.officeCellPropertyRevisionOmml ?? '',
  );
}

export function applyDocumentTablePropertyRevisionOmmlToElement(
  element: HTMLElement,
  omml: string | null | undefined,
): void {
  const encoded = omml ? encodeDocumentTablePropertyRevisionOmml(omml) : '';
  if (encoded) element.dataset.officeTablePropertyRevisionOmml = encoded;
  else delete element.dataset.officeTablePropertyRevisionOmml;
}

export function applyDocumentRowPropertyRevisionOmmlToElement(
  element: HTMLElement,
  omml: string | null | undefined,
): void {
  const encoded = omml ? encodeDocumentTablePropertyRevisionOmml(omml) : '';
  if (encoded) element.dataset.officeRowPropertyRevisionOmml = encoded;
  else delete element.dataset.officeRowPropertyRevisionOmml;
}

export function applyDocumentCellPropertyRevisionOmmlToElement(
  element: HTMLElement,
  omml: string | null | undefined,
): void {
  const encoded = omml ? encodeDocumentTablePropertyRevisionOmml(omml) : '';
  if (encoded) element.dataset.officeCellPropertyRevisionOmml = encoded;
  else delete element.dataset.officeCellPropertyRevisionOmml;
}

export function docxElementSubtreeHasRelationshipBindings(
  element: Element,
): boolean {
  const stack: Element[] = [element];
  while (stack.length) {
    const current = stack.pop();
    if (!current) continue;
    for (const attribute of Array.from(current.attributes)) {
      const namespace =
        attribute.namespaceURI ||
        xmlAttributeNamespace(current, attribute) ||
        '';
      if (DOCX_RELATIONSHIP_NAMESPACES.has(namespace)) return true;
    }
    for (const child of Array.from(current.children)) {
      stack.push(child);
    }
  }
  return false;
}

export function serializePreservableDocxScopedPropertyRevision(
  properties: Element | null | undefined,
  kind: DocxTableScopedPropertyRevisionKind,
): string | null {
  const change = properties ? directChild(properties, kind) : null;
  if (!change) return null;
  const namespace = change.namespaceURI ?? '';
  if (!DOCX_WORDPROCESSING_NAMESPACES.has(namespace)) return null;
  if (docxElementSubtreeHasRelationshipBindings(change)) return null;
  const prefix = change.prefix || 'w';
  const shell = parseXml(
    `<${prefix}:${change.localName} xmlns:${prefix}="${namespace}"/>`,
    'opaque scoped property revision shell',
  );
  const clone = cloneXmlElement(shell, change);
  declareInheritedNamespaces(clone, change);
  shell.replaceChild(clone, shell.documentElement);
  const omml = new XMLSerializer().serializeToString(clone).trim();
  if (!omml || omml.length > MAX_DOCUMENT_TABLE_PROPERTY_REVISION_OMML_LENGTH) {
    return null;
  }
  return omml;
}

export function serializePreservableDocxTablePropertyRevision(
  tableProperties: Element | null | undefined,
): string | null {
  // Reviewable relationship-free w:tblPrChange (jc/tblW/tblInd) is promoted to
  // a table-formatting change; keep the opaque OMML path for broader shapes.
  const change = tableProperties
    ? directChild(tableProperties, 'tblPrChange')
    : null;
  if (change && isSupportedDocxTableFormattingChange(change)) {
    return null;
  }
  return serializePreservableDocxScopedPropertyRevision(
    tableProperties,
    'tblPrChange',
  );
}

export function serializePreservableDocxRowPropertyRevision(
  rowProperties: Element | null | undefined,
): string | null {
  const change = rowProperties
    ? directChild(rowProperties, 'trPrChange')
    : null;
  if (change && isSupportedDocxRowFormattingChange(change)) {
    return null;
  }
  return serializePreservableDocxScopedPropertyRevision(
    rowProperties,
    'trPrChange',
  );
}

export function serializePreservableDocxCellPropertyRevision(
  cellProperties: Element | null | undefined,
): string | null {
  const change = cellProperties
    ? directChild(cellProperties, 'tcPrChange')
    : null;
  if (change && isSupportedDocxCellFormattingChange(change)) {
    return null;
  }
  return serializePreservableDocxScopedPropertyRevision(
    cellProperties,
    'tcPrChange',
  );
}

export function serializePreservableDocxSectionPropertyRevision(
  sectionProperties: Element | null | undefined,
): string | null {
  const change = sectionProperties
    ? directChild(sectionProperties, 'sectPrChange')
    : null;
  if (change && isSupportedDocxSectionFormattingChange(change)) {
    return null;
  }
  return serializePreservableDocxScopedPropertyRevision(
    sectionProperties,
    'sectPrChange',
  );
}

export function importDocxScopedPropertyRevisionElement(
  document: Document,
  omml: string,
  kind: DocxTableScopedPropertyRevisionKind,
): Element | null {
  try {
    const parsed = parseXml(omml, 'preserved scoped property revision');
    const root = parsed.documentElement;
    if (
      root.localName !== kind ||
      !DOCX_WORDPROCESSING_NAMESPACES.has(root.namespaceURI ?? '')
    ) {
      return null;
    }
    if (docxElementSubtreeHasRelationshipBindings(root)) return null;
    const imported = cloneXmlElement(document, root);
    declareInheritedNamespaces(imported, root);
    return imported;
  } catch {
    return null;
  }
}

export function importDocxTablePropertyRevisionElement(
  document: Document,
  omml: string,
): Element | null {
  return importDocxScopedPropertyRevisionElement(document, omml, 'tblPrChange');
}
