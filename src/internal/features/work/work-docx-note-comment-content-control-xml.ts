import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import { isKnownOoxmlNamespace } from './work-docx-note-comment-hyperlink-content';
import { xmlNamespacePrefix } from './work-ooxml-package';
import {
  XMLNS_NAMESPACE,
  xmlAttributeLocalName,
  xmlAttributeNamespace,
  xmlNamespaceUri,
} from './work-docx-settings-xml';

export const CONTENT_CONTROL_MC_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';
export const CONTENT_CONTROL_RELATIONSHIP_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  'http://purl.oclc.org/ooxml/officeDocument/relationships',
]);
export const CONTENT_CONTROL_WORD_2012_NAMESPACE =
  'http://schemas.microsoft.com/office/word/2012/wordml';

export function isContentControlSemanticNamespace(namespace: string): boolean {
  return (
    namespace === CONTENT_CONTROL_MC_NAMESPACE ||
    CONTENT_CONTROL_RELATIONSHIP_NAMESPACES.has(namespace) ||
    isKnownOoxmlNamespace(namespace)
  );
}

export function isDocxWordElement(element: Element): boolean {
  return DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '');
}

export function wordDirectChildren(
  element: Element,
  localName: string,
): Element[] {
  return Array.from(element.children).filter(
    (child) => child.localName === localName && isDocxWordElement(child),
  );
}

export function hasOnlyPassiveContentControlAttributes(
  element: Element,
): boolean {
  return Array.from(element.attributes).every((item) => {
    const namespace = xmlAttributeNamespace(element, item);
    return Boolean(
      namespace === XMLNS_NAMESPACE ||
        namespace === CONTENT_CONTROL_MC_NAMESPACE ||
        (namespace &&
          !CONTENT_CONTROL_RELATIONSHIP_NAMESPACES.has(namespace) &&
          !isKnownOoxmlNamespace(namespace)),
    );
  });
}

export function hasUnsupportedContentControlSemanticChild(
  element: Element,
): boolean {
  return Array.from(element.children).some(
    (child) =>
      !isDocxWordElement(child) &&
      isContentControlSemanticNamespace(child.namespaceURI ?? ''),
  );
}

export function hasContentControlRelationshipReference(root: Element): boolean {
  return [root, ...Array.from(root.querySelectorAll('*'))].some((element) =>
    Array.from(element.attributes).some((item) =>
      CONTENT_CONTROL_RELATIONSHIP_NAMESPACES.has(
        xmlAttributeNamespace(element, item) ?? '',
      ),
    ),
  );
}

export function createWordElement(
  document: Document,
  context: Element,
  localName: string,
): Element {
  const namespace =
    context.namespaceURI ?? [...DOCX_WORDPROCESSING_NAMESPACES][0];
  const prefix = xmlNamespacePrefix(context, namespace) ?? 'w';
  return document.createElementNS(namespace, `${prefix}:${localName}`);
}

export function createNamespacedElement(
  document: Document,
  context: Element,
  namespace: string,
  preferredPrefix: string,
  localName: string,
): Element {
  const prefix = availableNamespacePrefix(
    context.ownerDocument.documentElement,
    namespace,
    preferredPrefix,
  );
  return document.createElementNS(namespace, `${prefix}:${localName}`);
}

export function setWordContentControlAttribute(
  element: Element,
  context: Element,
  localName: string,
  value: string,
): void {
  const namespace =
    context.namespaceURI ?? [...DOCX_WORDPROCESSING_NAMESPACES][0];
  const prefix = xmlNamespacePrefix(context, namespace) ?? 'w';
  element.setAttributeNS(namespace, `${prefix}:${localName}`, value);
}

export function setNamespacedContentControlAttribute(
  element: Element,
  context: Element,
  namespace: string,
  preferredPrefix: string,
  localName: string,
  value: string,
): void {
  const prefix = availableNamespacePrefix(
    context.ownerDocument.documentElement,
    namespace,
    preferredPrefix,
  );
  element.setAttributeNS(namespace, `${prefix}:${localName}`, value);
}

export function ensureIgnorableContentControlNamespace(
  root: Element,
  namespace: string,
  preferredPrefix: string,
): void {
  const prefix = availableNamespacePrefix(root, namespace, preferredPrefix);
  if (!xmlNamespacePrefix(root, namespace)) {
    root.setAttributeNS(XMLNS_NAMESPACE, `xmlns:${prefix}`, namespace);
  }
  const mcPrefix = availableNamespacePrefix(
    root,
    CONTENT_CONTROL_MC_NAMESPACE,
    'mc',
  );
  if (!xmlNamespacePrefix(root, CONTENT_CONTROL_MC_NAMESPACE)) {
    root.setAttributeNS(
      XMLNS_NAMESPACE,
      `xmlns:${mcPrefix}`,
      CONTENT_CONTROL_MC_NAMESPACE,
    );
  }
  const existing = Array.from(root.attributes).find(
    (item) =>
      xmlAttributeLocalName(item) === 'Ignorable' &&
      xmlAttributeNamespace(root, item) === CONTENT_CONTROL_MC_NAMESPACE,
  );
  const prefixes = new Set(
    (existing?.value ?? '').trim().split(/\s+/u).filter(Boolean),
  );
  prefixes.add(prefix);
  const value = Array.from(prefixes).join(' ');
  if (existing) existing.value = value;
  else {
    root.setAttributeNS(
      CONTENT_CONTROL_MC_NAMESPACE,
      `${mcPrefix}:Ignorable`,
      value,
    );
  }
}

function availableNamespacePrefix(
  root: Element,
  namespace: string,
  preferred: string,
): string {
  const existing = xmlNamespacePrefix(root, namespace);
  if (existing) return existing;
  let prefix = preferred;
  let index = 1;
  while (xmlNamespaceUri(root, prefix)) {
    prefix = `${preferred}${index}`;
    index += 1;
  }
  return prefix;
}
