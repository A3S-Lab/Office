import {
  DOCUMENT_PAGE_BORDER_EDGES,
  type WorkDocumentPageBorderDisplay,
  type WorkDocumentPageBorderEdge,
  type WorkDocumentPageBorderOffsetFrom,
  type WorkDocumentPageBorders,
  type WorkDocumentPageBorderZOrder,
  normalizeDocumentPageBorders,
} from './work-document-page-borders';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import { parseDocxBorderElement } from './work-docx-paragraph-borders-import';
import {
  XMLNS_NAMESPACE,
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import {
  type DocxThemeSource,
  resolveDocxThemeResolver,
} from './work-docx-theme';
import { directChildren } from './work-ooxml-package';

export interface InspectedDocxPageBorders {
  status: 'absent' | 'valid' | 'invalid';
  pageBorders?: WorkDocumentPageBorders;
  invalidCount: number;
  spoofedCount: number;
}

const PAGE_BORDER_EDGE_SET = new Set<string>(DOCUMENT_PAGE_BORDER_EDGES);
const PAGE_BORDER_ATTRIBUTES = new Set(['zOrder', 'display', 'offsetFrom']);
const PAGE_BORDER_DISPLAYS = new Set<WorkDocumentPageBorderDisplay>([
  'allPages',
  'firstPage',
  'notFirstPage',
]);
const PAGE_BORDER_OFFSETS = new Set<WorkDocumentPageBorderOffsetFrom>([
  'page',
  'text',
]);
const PAGE_BORDER_Z_ORDERS = new Set<WorkDocumentPageBorderZOrder>([
  'front',
  'back',
]);

export function inspectDocxPageBorders(
  sectionProperties: Element,
  themeSource?: DocxThemeSource,
): InspectedDocxPageBorders {
  if (
    sectionProperties.localName !== 'sectPr' ||
    !DOCX_WORDPROCESSING_NAMESPACES.has(sectionProperties.namespaceURI ?? '')
  ) {
    return {
      status: 'absent',
      invalidCount: 0,
      spoofedCount: 0,
    };
  }
  const named = directChildren(sectionProperties, 'pgBorders');
  const candidates = named.filter(
    (element) => element.namespaceURI === sectionProperties.namespaceURI,
  );
  const spoofedContainers = named.length - candidates.length;
  if (!candidates.length) {
    return {
      status: 'absent',
      invalidCount: 0,
      spoofedCount: spoofedContainers,
    };
  }
  if (candidates.length !== 1) {
    return {
      status: 'invalid',
      invalidCount: candidates.length,
      spoofedCount: spoofedContainers,
    };
  }
  const container = candidates[0];
  if (!container) {
    return {
      status: 'invalid',
      invalidCount: 1,
      spoofedCount: spoofedContainers,
    };
  }
  const attributes = wordContainerAttributes(container);
  if (!attributes || hasNonWhitespaceText(container)) {
    return {
      status: 'invalid',
      invalidCount: 1,
      spoofedCount: spoofedContainers,
    };
  }
  const display = pageBorderDisplay(attributes.get('display'));
  const offsetFrom = pageBorderOffset(attributes.get('offsetFrom'));
  const zOrder = pageBorderZOrder(attributes.get('zOrder'));
  if (display === null || offsetFrom === null || zOrder === null) {
    return {
      status: 'invalid',
      invalidCount: 1,
      spoofedCount: spoofedContainers,
    };
  }

  const theme = resolveDocxThemeResolver(themeSource);
  const edges: WorkDocumentPageBorders['edges'] = {};
  let invalidCount = 0;
  let spoofedCount = spoofedContainers;
  let previousIndex = -1;
  const seen = new Set<WorkDocumentPageBorderEdge>();
  for (const child of directChildren(container)) {
    if (child.namespaceURI !== container.namespaceURI) {
      spoofedCount += 1;
      continue;
    }
    const edge = child.localName as WorkDocumentPageBorderEdge;
    const index = DOCUMENT_PAGE_BORDER_EDGES.indexOf(edge);
    if (
      !PAGE_BORDER_EDGE_SET.has(edge) ||
      index < previousIndex ||
      seen.has(edge)
    ) {
      return {
        status: 'invalid',
        invalidCount: invalidCount + 1,
        spoofedCount,
      };
    }
    previousIndex = index;
    seen.add(edge);
    const border = parseDocxBorderElement(child, theme);
    if (!border) {
      invalidCount += 1;
      edges[edge] = { style: 'nil' };
    } else {
      edges[edge] = border;
    }
  }
  const pageBorders = normalizeDocumentPageBorders({
    ...(display ? { display } : {}),
    ...(offsetFrom ? { offsetFrom } : {}),
    ...(zOrder ? { zOrder } : {}),
    edges,
  });
  if (!pageBorders) {
    return {
      status: 'invalid',
      invalidCount: invalidCount + 1,
      spoofedCount,
    };
  }
  return {
    status: 'valid',
    pageBorders,
    invalidCount,
    spoofedCount,
  };
}

export function parseDocxPageBorders(
  sectionProperties: Element,
  themeSource?: DocxThemeSource,
): WorkDocumentPageBorders | null | undefined {
  const inspected = inspectDocxPageBorders(sectionProperties, themeSource);
  if (inspected.status === 'absent') return undefined;
  return inspected.status === 'valid' ? inspected.pageBorders : null;
}

function wordContainerAttributes(element: Element): Map<string, string> | null {
  const result = new Map<string, string>();
  for (const attribute of Array.from(element.attributes)) {
    if (isNamespaceDeclaration(attribute)) continue;
    const name = xmlAttributeLocalName(attribute);
    if (
      xmlAttributeNamespace(element, attribute) !== element.namespaceURI ||
      !PAGE_BORDER_ATTRIBUTES.has(name) ||
      result.has(name)
    ) {
      return null;
    }
    result.set(name, attribute.value);
  }
  return result;
}

function pageBorderDisplay(
  value: string | undefined,
): WorkDocumentPageBorderDisplay | null | undefined {
  if (value === undefined) return undefined;
  const normalized = value.trim() as WorkDocumentPageBorderDisplay;
  return PAGE_BORDER_DISPLAYS.has(normalized) ? normalized : null;
}

function pageBorderOffset(
  value: string | undefined,
): WorkDocumentPageBorderOffsetFrom | null | undefined {
  if (value === undefined) return undefined;
  const normalized = value.trim() as WorkDocumentPageBorderOffsetFrom;
  return PAGE_BORDER_OFFSETS.has(normalized) ? normalized : null;
}

function pageBorderZOrder(
  value: string | undefined,
): WorkDocumentPageBorderZOrder | null | undefined {
  if (value === undefined) return undefined;
  const normalized = value.trim() as WorkDocumentPageBorderZOrder;
  return PAGE_BORDER_Z_ORDERS.has(normalized) ? normalized : null;
}

function hasNonWhitespaceText(element: Element): boolean {
  return Array.from(element.childNodes).some(
    (node) => node.nodeType === Node.TEXT_NODE && node.textContent?.trim(),
  );
}

function isNamespaceDeclaration(attribute: Attr): boolean {
  return (
    attribute.namespaceURI === XMLNS_NAMESPACE ||
    attribute.name === 'xmlns' ||
    attribute.name.startsWith('xmlns:')
  );
}
