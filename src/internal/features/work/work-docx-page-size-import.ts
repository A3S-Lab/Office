import {
  MAX_WORD_PAGE_TWIPS,
  MAX_WORD_PAPER_CODE,
  MAX_WORD_PAPER_SOURCE_CODE,
  type WorkDocumentPageGeometry,
  type WorkDocumentPaperSource,
  normalizeDocumentPageGeometry,
  normalizeDocumentPaperSource,
} from './work-document-page-size';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import {
  XMLNS_NAMESPACE,
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import {
  parseBoundedDocxInteger,
  parseDocxTwipsMeasure,
  STRICT_WORDPROCESSING_NAMESPACE,
} from './work-docx-twips';
import { directChildren } from './work-ooxml-package';

export interface InspectedDocxPageSize {
  status: 'absent' | 'valid' | 'invalid';
  pageGeometry?: WorkDocumentPageGeometry;
  invalidCount: number;
  spoofedCount: number;
}

export interface InspectedDocxPaperSource {
  status: 'absent' | 'valid' | 'invalid';
  paperSource?: WorkDocumentPaperSource;
  invalidCount: number;
  spoofedCount: number;
}

const PAGE_SIZE_ATTRIBUTE_SET = new Set(['w', 'h', 'orient', 'code']);
const PAPER_SOURCE_ATTRIBUTE_SET = new Set(['first', 'other']);

export function inspectDocxPageSize(
  sectionProperties: Element,
): InspectedDocxPageSize {
  if (!isWordSection(sectionProperties)) return absentPageSize();
  const inspected = inspectLeaf(sectionProperties, 'pgSz');
  if (inspected.status !== 'valid' || !inspected.element) {
    return {
      status: inspected.status,
      invalidCount: inspected.invalidCount,
      spoofedCount: inspected.spoofedCount,
    };
  }
  const attributes = wordAttributes(inspected.element, PAGE_SIZE_ATTRIBUTE_SET);
  if (!attributes?.has('w') || !attributes.has('h')) {
    return invalidPageSize(inspected.spoofedCount);
  }
  const strict =
    inspected.element.namespaceURI === STRICT_WORDPROCESSING_NAMESPACE;
  const width = parseDocxTwipsMeasure(attributes.get('w') ?? '', {
    minimum: 1,
    maximum: MAX_WORD_PAGE_TWIPS,
    signed: false,
    strict,
  });
  const height = parseDocxTwipsMeasure(attributes.get('h') ?? '', {
    minimum: 1,
    maximum: MAX_WORD_PAGE_TWIPS,
    signed: false,
    strict,
  });
  const orientation = attributes.get('orient');
  const codeSource = attributes.get('code');
  const code =
    codeSource === undefined
      ? undefined
      : parseBoundedDocxInteger(codeSource, {
          minimum: 0,
          maximum: MAX_WORD_PAPER_CODE,
        });
  if (
    width === null ||
    height === null ||
    (orientation !== undefined &&
      orientation !== 'portrait' &&
      orientation !== 'landscape') ||
    (codeSource !== undefined && code === null)
  ) {
    return invalidPageSize(inspected.spoofedCount);
  }
  const pageGeometry = normalizeDocumentPageGeometry({
    width,
    height,
    ...(orientation !== undefined ? { orientation } : {}),
    ...(code !== undefined ? { code } : {}),
  });
  return pageGeometry
    ? {
        status: 'valid',
        pageGeometry,
        invalidCount: 0,
        spoofedCount: inspected.spoofedCount,
      }
    : invalidPageSize(inspected.spoofedCount);
}

export function inspectDocxPaperSource(
  sectionProperties: Element,
): InspectedDocxPaperSource {
  if (!isWordSection(sectionProperties)) return absentPaperSource();
  const inspected = inspectLeaf(sectionProperties, 'paperSrc');
  if (inspected.status !== 'valid' || !inspected.element) {
    return {
      status: inspected.status,
      invalidCount: inspected.invalidCount,
      spoofedCount: inspected.spoofedCount,
    };
  }
  const attributes = wordAttributes(
    inspected.element,
    PAPER_SOURCE_ATTRIBUTE_SET,
  );
  if (!attributes) return invalidPaperSource(inspected.spoofedCount);
  const source: WorkDocumentPaperSource = {};
  for (const key of ['first', 'other'] as const) {
    const value = attributes.get(key);
    if (value === undefined) continue;
    const parsed = parseBoundedDocxInteger(value, {
      minimum: 0,
      maximum: MAX_WORD_PAPER_SOURCE_CODE,
    });
    if (parsed === null) return invalidPaperSource(inspected.spoofedCount);
    source[key] = parsed;
  }
  const paperSource = normalizeDocumentPaperSource(source);
  return paperSource
    ? {
        status: 'valid',
        paperSource,
        invalidCount: 0,
        spoofedCount: inspected.spoofedCount,
      }
    : invalidPaperSource(inspected.spoofedCount);
}

export function parseDocxPageGeometry(
  sectionProperties: Element,
  fallback?: WorkDocumentPageGeometry,
): WorkDocumentPageGeometry | null | undefined {
  const inspected = inspectDocxPageSize(sectionProperties);
  if (inspected.status === 'valid') return inspected.pageGeometry;
  if (inspected.status === 'invalid') return null;
  return normalizeDocumentPageGeometry(fallback) ?? undefined;
}

export function parseDocxPaperSource(
  sectionProperties: Element,
  fallback?: WorkDocumentPaperSource,
): WorkDocumentPaperSource | null | undefined {
  const inspected = inspectDocxPaperSource(sectionProperties);
  if (inspected.status === 'valid') return inspected.paperSource;
  if (inspected.status === 'invalid') return null;
  return normalizeDocumentPaperSource(fallback) ?? undefined;
}

function inspectLeaf(
  parent: Element,
  localName: 'paperSrc' | 'pgSz',
): {
  status: 'absent' | 'valid' | 'invalid';
  element?: Element;
  invalidCount: number;
  spoofedCount: number;
} {
  const named = directChildren(parent, localName);
  const candidates = named.filter(
    (element) => element.namespaceURI === parent.namespaceURI,
  );
  const spoofedCount = named.length - candidates.length;
  if (!candidates.length) {
    return { status: 'absent', invalidCount: 0, spoofedCount };
  }
  if (candidates.length !== 1) {
    return {
      status: 'invalid',
      invalidCount: candidates.length,
      spoofedCount,
    };
  }
  const element = candidates[0];
  if (!element || hasElementOrNonWhitespaceContent(element)) {
    return { status: 'invalid', invalidCount: 1, spoofedCount };
  }
  return {
    status: 'valid',
    element,
    invalidCount: 0,
    spoofedCount,
  };
}

function wordAttributes(
  element: Element,
  allowed: ReadonlySet<string>,
): Map<string, string> | null {
  const result = new Map<string, string>();
  for (const item of Array.from(element.attributes)) {
    if (isNamespaceDeclaration(item)) continue;
    const name = xmlAttributeLocalName(item);
    if (
      xmlAttributeNamespace(element, item) !== element.namespaceURI ||
      !allowed.has(name) ||
      result.has(name)
    ) {
      return null;
    }
    result.set(name, item.value.trim());
  }
  return result;
}

function isWordSection(element: Element): boolean {
  return (
    element.localName === 'sectPr' &&
    DOCX_WORDPROCESSING_NAMESPACES.has(element.namespaceURI ?? '')
  );
}

function hasElementOrNonWhitespaceContent(element: Element): boolean {
  return Array.from(element.childNodes).some(
    (node) =>
      node.nodeType === Node.ELEMENT_NODE ||
      ((node.nodeType === Node.TEXT_NODE ||
        node.nodeType === Node.CDATA_SECTION_NODE) &&
        Boolean(node.textContent?.trim())),
  );
}

function isNamespaceDeclaration(attribute: Attr): boolean {
  return (
    attribute.namespaceURI === XMLNS_NAMESPACE ||
    attribute.name === 'xmlns' ||
    attribute.name.startsWith('xmlns:')
  );
}

function absentPageSize(): InspectedDocxPageSize {
  return { status: 'absent', invalidCount: 0, spoofedCount: 0 };
}

function invalidPageSize(spoofedCount: number): InspectedDocxPageSize {
  return { status: 'invalid', invalidCount: 1, spoofedCount };
}

function absentPaperSource(): InspectedDocxPaperSource {
  return { status: 'absent', invalidCount: 0, spoofedCount: 0 };
}

function invalidPaperSource(spoofedCount: number): InspectedDocxPaperSource {
  return { status: 'invalid', invalidCount: 1, spoofedCount };
}
