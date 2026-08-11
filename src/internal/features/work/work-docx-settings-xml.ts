import { attribute, xmlNamespacePrefix } from './work-ooxml-package';

export const XMLNS_NAMESPACE = 'http://www.w3.org/2000/xmlns/';
export const XML_NAMESPACE = 'http://www.w3.org/XML/1998/namespace';
const MARKUP_COMPATIBILITY_NAMESPACE =
  'http://schemas.openxmlformats.org/markup-compatibility/2006';

export function declareInheritedNamespaces(
  importedRoot: Element,
  sourceRoot: Element,
): void {
  const mappings = new Map<string, string>();
  for (const element of [
    sourceRoot,
    ...Array.from(sourceRoot.querySelectorAll('*')),
  ]) {
    if (element.prefix && element.namespaceURI && element.prefix !== 'xml') {
      if (!mappings.has(element.prefix)) {
        mappings.set(element.prefix, element.namespaceURI);
      }
    }
    for (const item of Array.from(element.attributes)) {
      const prefix = xmlAttributePrefix(item);
      const namespace = xmlAttributeNamespace(element, item);
      if (
        prefix &&
        prefix !== 'xmlns' &&
        prefix !== 'xml' &&
        namespace &&
        !mappings.has(prefix)
      ) {
        mappings.set(prefix, namespace);
      }
    }
    if (
      element.namespaceURI === MARKUP_COMPATIBILITY_NAMESPACE &&
      element.localName === 'Choice'
    ) {
      for (const prefix of (attribute(element, 'Requires') ?? '')
        .trim()
        .split(/\s+/)
        .filter(Boolean)) {
        const namespace = xmlNamespaceUri(element, prefix);
        if (namespace && !mappings.has(prefix)) {
          mappings.set(prefix, namespace);
        }
      }
    }
  }
  for (const [prefix, namespace] of mappings) {
    importedRoot.setAttributeNS(XMLNS_NAMESPACE, `xmlns:${prefix}`, namespace);
  }
}

export function collectRetainedNamespaces(
  root: Element,
  retained: Set<string>,
): void {
  for (const element of [root, ...Array.from(root.querySelectorAll('*'))]) {
    if (element.namespaceURI) retained.add(element.namespaceURI);
    for (const item of Array.from(element.attributes)) {
      const namespace = xmlAttributeNamespace(element, item);
      if (namespace && namespace !== XMLNS_NAMESPACE) {
        retained.add(namespace);
      }
    }
  }
}

/** Clones XML without relying on DOM importNode prefix preservation. */
export function cloneXmlElement(document: Document, source: Element): Element {
  const namespace = source.namespaceURI;
  const prefix = xmlNamespacePrefix(source, namespace);
  const qualifiedName = prefix
    ? `${prefix}:${source.localName}`
    : source.localName;
  const clone = document.createElementNS(namespace, qualifiedName);
  for (const item of Array.from(source.attributes)) {
    const namespaceUri = xmlAttributeNamespace(source, item);
    if (namespaceUri) {
      clone.setAttributeNS(namespaceUri, item.name, item.value);
    } else {
      clone.setAttribute(item.name, item.value);
    }
  }
  for (const child of Array.from(source.childNodes)) {
    if (child instanceof Element) {
      clone.append(cloneXmlElement(document, child));
    } else if (child.nodeType === Node.TEXT_NODE) {
      clone.append(document.createTextNode(child.textContent ?? ''));
    } else if (child.nodeType === Node.CDATA_SECTION_NODE) {
      clone.append(document.createCDATASection(child.textContent ?? ''));
    } else if (child.nodeType === Node.COMMENT_NODE) {
      clone.append(document.createComment(child.textContent ?? ''));
    }
  }
  return clone;
}

export function xmlNamespaceUri(
  element: Element,
  prefix: string,
): string | null {
  if (prefix === 'xml') return XML_NAMESPACE;
  let current: Element | null = element;
  while (current) {
    if (current.prefix === prefix && current.namespaceURI) {
      return current.namespaceURI;
    }
    const declaration = Array.from(current.attributes).find(
      (item) =>
        isNamespaceDeclaration(item) &&
        (prefix ? item.name === `xmlns:${prefix}` : item.name === 'xmlns'),
    );
    if (declaration) return declaration.value;
    current = current.parentElement;
  }
  return null;
}

export function xmlDeclaredPrefix(
  root: Element,
  namespace: string,
): string | null {
  for (const item of Array.from(root.attributes)) {
    if (!isNamespaceDeclaration(item) || item.value !== namespace) continue;
    return item.name === 'xmlns' ? null : item.name.slice('xmlns:'.length);
  }
  return null;
}

export function xmlAttributeNamespace(
  element: Element,
  item: Attr,
): string | null {
  if (item.namespaceURI) return item.namespaceURI;
  const prefix = xmlAttributePrefix(item);
  return prefix ? xmlNamespaceUri(element, prefix) : null;
}

export function xmlAttributePrefix(item: Attr): string | null {
  if (item.prefix) return item.prefix;
  const separator = item.name.indexOf(':');
  return separator > 0 ? item.name.slice(0, separator) : null;
}

export function xmlAttributeLocalName(item: Attr): string {
  const separator = item.localName.indexOf(':');
  return separator >= 0 ? item.localName.slice(separator + 1) : item.localName;
}

export function hasXmlAttribute(
  element: Element,
  namespace: string,
  localName: string,
): boolean {
  return Array.from(element.attributes).some(
    (item) =>
      xmlAttributeNamespace(element, item) === namespace &&
      xmlAttributeLocalName(item) === localName,
  );
}

export function hasUnexpectedXmlAttributes(
  element: Element,
  allowedUnqualified: ReadonlySet<string>,
): boolean {
  return Array.from(element.attributes).some((item) => {
    if (isNamespaceDeclaration(item)) return false;
    return (
      xmlAttributeNamespace(element, item) !== null ||
      !allowedUnqualified.has(item.name)
    );
  });
}

export function hasNonWhitespaceXmlText(element: Element): boolean {
  return Array.from(element.childNodes).some(
    (child) =>
      (child.nodeType === Node.TEXT_NODE ||
        child.nodeType === Node.CDATA_SECTION_NODE) &&
      Boolean(child.textContent?.trim()),
  );
}

export function isStructurallyValidAlternateContent(element: Element): boolean {
  if (
    hasNonWhitespaceXmlText(element) ||
    hasUnexpectedXmlAttributes(element, new Set())
  ) {
    return false;
  }
  const children = Array.from(element.children);
  let choices = 0;
  let fallbacks = 0;
  let sawFallback = false;
  for (const child of children) {
    if (
      child.namespaceURI !== MARKUP_COMPATIBILITY_NAMESPACE ||
      (child.localName !== 'Choice' && child.localName !== 'Fallback')
    ) {
      return false;
    }
    const allowedAttributes =
      child.localName === 'Choice' ? new Set(['Requires']) : new Set<string>();
    if (
      hasNonWhitespaceXmlText(child) ||
      hasUnexpectedXmlAttributes(child, allowedAttributes)
    ) {
      return false;
    }
    if (child.localName === 'Fallback') {
      fallbacks += 1;
      sawFallback = true;
      if (fallbacks > 1) return false;
      continue;
    }
    if (sawFallback) return false;
    choices += 1;
    const required = (attribute(child, 'Requires') ?? '')
      .trim()
      .split(/\s+/)
      .filter(Boolean);
    if (
      !required.length ||
      required.some(
        (prefix) =>
          !/^[A-Za-z_][A-Za-z0-9_.-]*$/.test(prefix) ||
          !xmlNamespaceUri(child, prefix),
      )
    ) {
      return false;
    }
  }
  return choices > 0;
}

export function assertXmlRoot(
  root: Element,
  localName: string,
  namespaces: ReadonlySet<string>,
  message: string,
): void {
  if (
    root.localName !== localName ||
    !namespaces.has(root.namespaceURI ?? '')
  ) {
    throw new Error(message);
  }
}

export function decodeXmlBytes(bytes: Uint8Array, label: string): string {
  let encoding: 'utf-8' | 'utf-16le' | 'utf-16be' = 'utf-8';
  let offset = 0;
  if (bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    offset = 3;
  } else if (bytes[0] === 0xff && bytes[1] === 0xfe) {
    encoding = 'utf-16le';
    offset = 2;
  } else if (bytes[0] === 0xfe && bytes[1] === 0xff) {
    encoding = 'utf-16be';
    offset = 2;
  }
  try {
    return new TextDecoder(encoding, { fatal: true }).decode(
      bytes.subarray(offset),
    );
  } catch {
    throw new Error(`${label} uses an invalid ${encoding} XML encoding.`);
  }
}

export function serializeUtf8Xml(document: Document): string {
  const serialized = new XMLSerializer().serializeToString(document);
  const body = serialized.replace(/^\s*<\?xml[^?]*\?>\s*/i, '');
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${body}`;
}

function isNamespaceDeclaration(attributeNode: Attr): boolean {
  return (
    attributeNode.namespaceURI === XMLNS_NAMESPACE ||
    attributeNode.name === 'xmlns' ||
    attributeNode.name.startsWith('xmlns:')
  );
}
