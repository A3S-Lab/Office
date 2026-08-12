import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import { xmlNamespacePrefix } from './work-ooxml-package';
import {
  XML_NAMESPACE,
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';

const RELATIONSHIP_NAMESPACES = new Set([
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  'http://purl.oclc.org/ooxml/officeDocument/relationships',
]);
const MAX_PROPERTY_ELEMENTS = 256;
const MAX_PROPERTY_DEPTH = 32;

const COMMENT_RUN_PROPERTIES = new Set([
  'b',
  'bCs',
  'bdr',
  'caps',
  'color',
  'cs',
  'dstrike',
  'eastAsianLayout',
  'effect',
  'em',
  'emboss',
  'fitText',
  'highlight',
  'i',
  'iCs',
  'imprint',
  'kern',
  'lang',
  'noProof',
  'outline',
  'position',
  'rtl',
  'rFonts',
  'shadow',
  'shd',
  'smallCaps',
  'snapToGrid',
  'spacing',
  'specVanish',
  'strike',
  'sz',
  'szCs',
  'u',
  'vanish',
  'vertAlign',
  'w',
  'webHidden',
]);

const NOTE_UNMODELED_RUN_PROPERTIES = new Set([
  'bdr',
  'caps',
  'eastAsianLayout',
  'effect',
  'em',
  'emboss',
  'fitText',
  'imprint',
  'kern',
  'lang',
  'noProof',
  'outline',
  'position',
  'shadow',
  'smallCaps',
  'spacing',
  'specVanish',
  'vanish',
  'w',
  'webHidden',
]);

export function preserveDocxNoteCommentRunProperties(
  generatedDocument: Document,
  generatedRun: Element,
  sourceRun: Element,
  kind: 'comment' | 'note',
): void {
  const sourceProperties = wordDirectChildren(sourceRun, 'rPr')[0];
  if (!sourceProperties) return;
  const allowed =
    kind === 'comment' ? COMMENT_RUN_PROPERTIES : NOTE_UNMODELED_RUN_PROPERTIES;
  const sourceGroups = eligiblePropertyGroups(sourceProperties, allowed);
  if (!sourceGroups.size) return;

  let generatedProperties = wordDirectChildren(generatedRun, 'rPr')[0];
  for (const [localName, sourceMatches] of sourceGroups) {
    if (sourceMatches.length !== 1) continue;
    if (
      generatedProperties &&
      wordDirectChildren(generatedProperties, localName).length > 0
    ) {
      continue;
    }
    if (hasRelationshipReference(sourceMatches[0])) continue;
    generatedProperties ??= createRunProperties(
      generatedDocument,
      generatedRun,
    );
    const clone = cloneWordProperty(
      generatedDocument,
      sourceMatches[0],
      generatedRun,
    );
    if (clone) generatedProperties.append(clone);
  }
}

export function cloneDocxNoteCommentRunProperties(
  document: Document,
  sourceProperties: Element,
  targetContext: Element,
  kind: 'comment' | 'note',
): Element | null {
  const allowed =
    kind === 'comment' ? COMMENT_RUN_PROPERTIES : NOTE_UNMODELED_RUN_PROPERTIES;
  const properties = newRunProperties(document, targetContext);
  for (const sourceMatches of eligiblePropertyGroups(
    sourceProperties,
    allowed,
  ).values()) {
    if (
      sourceMatches.length !== 1 ||
      hasRelationshipReference(sourceMatches[0])
    ) {
      continue;
    }
    const clone = cloneWordProperty(document, sourceMatches[0], targetContext);
    if (clone) properties.append(clone);
  }
  return properties.children.length ? properties : null;
}

function eligiblePropertyGroups(
  sourceProperties: Element,
  allowed: ReadonlySet<string>,
): Map<string, Element[]> {
  const sourceGroups = new Map<string, Element[]>();
  for (const property of Array.from(sourceProperties.children)) {
    if (!isWordElement(property) || !allowed.has(property.localName)) continue;
    const matches = sourceGroups.get(property.localName) ?? [];
    matches.push(property);
    sourceGroups.set(property.localName, matches);
  }
  return sourceGroups;
}

function createRunProperties(document: Document, run: Element): Element {
  const properties = newRunProperties(document, run);
  run.insertBefore(properties, run.firstChild);
  return properties;
}

function newRunProperties(document: Document, context: Element): Element {
  const namespace =
    context.namespaceURI ?? [...DOCX_WORDPROCESSING_NAMESPACES][0];
  const prefix = xmlNamespacePrefix(context, namespace);
  return document.createElementNS(namespace, prefix ? `${prefix}:rPr` : 'rPr');
}

function cloneWordProperty(
  document: Document,
  source: Element,
  target: Element,
): Element | null {
  const namespace = target.namespaceURI;
  if (!namespace || !DOCX_WORDPROCESSING_NAMESPACES.has(namespace)) return null;
  const prefix = xmlNamespacePrefix(target, namespace);
  let elementCount = 0;
  const clone = (element: Element, depth: number): Element | null => {
    elementCount += 1;
    if (
      depth > MAX_PROPERTY_DEPTH ||
      elementCount > MAX_PROPERTY_ELEMENTS ||
      !isWordElement(element)
    ) {
      return null;
    }
    const result = document.createElementNS(
      namespace,
      prefix ? `${prefix}:${element.localName}` : element.localName,
    );
    for (const item of Array.from(element.attributes)) {
      const itemNamespace = xmlAttributeNamespace(element, item);
      const localName = xmlAttributeLocalName(item);
      if (DOCX_WORDPROCESSING_NAMESPACES.has(itemNamespace ?? '')) {
        result.setAttributeNS(
          namespace,
          `${prefix ?? 'w'}:${localName}`,
          item.value,
        );
      } else if (itemNamespace === XML_NAMESPACE) {
        result.setAttributeNS(XML_NAMESPACE, `xml:${localName}`, item.value);
      } else if (!itemNamespace) {
        result.setAttribute(localName, item.value);
      }
    }
    for (const child of Array.from(element.childNodes)) {
      if (child.nodeType === Node.TEXT_NODE) {
        result.append(document.createTextNode(child.textContent ?? ''));
      } else if (child.nodeType === Node.ELEMENT_NODE) {
        const childElement = child as Element;
        if (!isWordElement(childElement)) continue;
        const childClone = clone(childElement, depth + 1);
        if (!childClone) return null;
        result.append(childClone);
      }
    }
    return result;
  };
  return clone(source, 1);
}

function hasRelationshipReference(root: Element): boolean {
  return [root, ...Array.from(root.querySelectorAll('*'))].some((element) =>
    Array.from(element.attributes).some((item) =>
      RELATIONSHIP_NAMESPACES.has(xmlAttributeNamespace(element, item) ?? ''),
    ),
  );
}

function wordDirectChildren(element: Element, localName: string): Element[] {
  return Array.from(element.children).filter(
    (child) => child.localName === localName && isWordElement(child),
  );
}

function isWordElement(element: Element): boolean {
  return DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '');
}
