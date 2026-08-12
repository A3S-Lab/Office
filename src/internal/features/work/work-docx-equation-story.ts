import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import { XML_NAMESPACE } from './work-docx-settings-xml';
import { xmlNamespacePrefix } from './work-ooxml-package';

export type DocxEquationPlacement = 'inline' | 'block';

export function isDocxEquationLikeRoot(element: Element): boolean {
  return element.localName === 'oMath' || element.localName === 'oMathPara';
}

export function closestDocxEquationLikeRoot(
  element: Element | null,
): Element | null {
  let current = element;
  while (current) {
    if (isDocxEquationLikeRoot(current)) return current;
    current = current.parentElement;
  }
  return null;
}

export function docxEquationPlacement(
  element: Element,
): DocxEquationPlacement | null {
  const parent = element.parentElement;
  if (!parent) return null;
  if (element.localName === 'oMath') {
    return parent.localName === 'p' &&
      DOCX_WORDPROCESSING_NAMESPACES.has(parent.namespaceURI ?? '')
      ? 'inline'
      : null;
  }
  if (element.localName !== 'oMathPara') return null;
  if (
    ['body', 'tc', 'hdr', 'ftr', 'footnote', 'endnote', 'sdtContent'].includes(
      parent.localName,
    ) &&
    DOCX_WORDPROCESSING_NAMESPACES.has(parent.namespaceURI ?? '')
  ) {
    return 'block';
  }
  return null;
}

export function docxEquationWordReplacement(
  document: Document,
  source: Element,
  text: string,
  placement: DocxEquationPlacement,
): Element {
  return wordFallback(document, source, text, placement);
}

export function docxEquationWordFallbackForContext(
  document: Document,
  source: Element,
  text: string,
  placement: DocxEquationPlacement | null,
): Element | null {
  if (placement) return wordFallback(document, source, text, placement);
  const parent = source.parentElement;
  if (
    !parent ||
    !DOCX_WORDPROCESSING_NAMESPACES.has(parent.namespaceURI ?? '')
  ) {
    return null;
  }
  const namespace = parent.namespaceURI ?? wordNamespace(source);
  const prefix = wordPrefix(source, namespace);
  if (parent.localName === 'r') {
    return wordText(document, namespace, prefix, text);
  }
  if (
    [
      'del',
      'hyperlink',
      'ins',
      'moveFrom',
      'moveTo',
      'p',
      'sdtContent',
    ].includes(parent.localName)
  ) {
    return wordRun(document, namespace, prefix, text);
  }
  if (
    ['body', 'endnote', 'footnote', 'ftr', 'hdr', 'tc'].includes(
      parent.localName,
    )
  ) {
    const paragraph = document.createElementNS(namespace, `${prefix}:p`);
    paragraph.append(wordRun(document, namespace, prefix, text));
    return paragraph;
  }
  return null;
}

export function replaceDocxEquationTextMarker(
  text: Text,
  marker: string,
  replacement: HTMLElement,
): void {
  const offset = text.data.indexOf(marker);
  if (offset < 0) return;
  const fragment = text.ownerDocument.createDocumentFragment();
  const before = text.data.slice(0, offset);
  const after = text.data.slice(offset + marker.length);
  if (before) fragment.append(text.ownerDocument.createTextNode(before));
  fragment.append(replacement);
  if (after) fragment.append(text.ownerDocument.createTextNode(after));
  text.replaceWith(fragment);
}

export function docxEquationTextNodes(root: ParentNode): Text[] {
  const walker = root.ownerDocument?.createTreeWalker(
    root,
    NodeFilter.SHOW_TEXT,
  );
  const nodes: Text[] = [];
  if (!walker) return nodes;
  while (walker.nextNode()) nodes.push(walker.currentNode as Text);
  return nodes;
}

export function escapeDocxEquationHtml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;');
}

function wordFallback(
  document: Document,
  source: Element,
  text: string,
  placement: DocxEquationPlacement,
): Element {
  const namespace = wordNamespace(source);
  const prefix = wordPrefix(source, namespace);
  const run = wordRun(document, namespace, prefix, text);
  if (placement === 'inline') return run;
  const paragraph = document.createElementNS(namespace, `${prefix}:p`);
  paragraph.append(run);
  return paragraph;
}

function wordRun(
  document: Document,
  namespace: string,
  prefix: string,
  text: string,
): Element {
  const run = document.createElementNS(namespace, `${prefix}:r`);
  run.append(wordText(document, namespace, prefix, text));
  return run;
}

function wordText(
  document: Document,
  namespace: string,
  prefix: string,
  text: string,
): Element {
  const value = document.createElementNS(namespace, `${prefix}:t`);
  value.setAttributeNS(XML_NAMESPACE, 'xml:space', 'preserve');
  value.textContent = text;
  return value;
}

function wordNamespace(element: Element): string {
  let current: Element | null = element.parentElement;
  while (current) {
    if (DOCX_WORDPROCESSING_NAMESPACES.has(current.namespaceURI ?? '')) {
      return current.namespaceURI ?? [...DOCX_WORDPROCESSING_NAMESPACES][0];
    }
    current = current.parentElement;
  }
  return [...DOCX_WORDPROCESSING_NAMESPACES][0];
}

function wordPrefix(element: Element, namespace: string): string {
  return xmlNamespacePrefix(element, namespace) || 'w';
}
