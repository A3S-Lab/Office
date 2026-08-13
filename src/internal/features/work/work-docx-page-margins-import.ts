import {
  type WorkDocumentPageMarginKey,
  type WorkDocumentPageMargins,
  normalizeDocumentPageMargins,
} from './work-document-page-margins';
import { DOCX_WORDPROCESSING_NAMESPACES } from './work-docx-ignorable-extension-preservation';
import {
  XMLNS_NAMESPACE,
  xmlAttributeLocalName,
  xmlAttributeNamespace,
} from './work-docx-settings-xml';
import { directChildren } from './work-ooxml-package';
import {
  parseDocxTwipsMeasure,
  STRICT_WORDPROCESSING_NAMESPACE,
} from './work-docx-twips';

export interface InspectedDocxPageMarginSettings {
  mirrorMargins?: boolean;
  gutterAtTop?: boolean;
  invalidCount: number;
  spoofedCount: number;
  incompatible: string[];
}

export interface InspectedDocxPageMargins {
  status: 'absent' | 'valid' | 'invalid';
  pageMargins?: WorkDocumentPageMargins;
  gutterOnRight?: boolean;
  invalidCount: number;
  spoofedCount: number;
}

const PAGE_MARGIN_ATTRIBUTE_SET = new Set([
  'top',
  'right',
  'bottom',
  'left',
  'header',
  'footer',
  'gutter',
]);
const PAGE_MARGIN_KEYS = [
  'top',
  'right',
  'bottom',
  'left',
  'header',
  'footer',
  'gutter',
] as const satisfies readonly WorkDocumentPageMarginKey[];
const MAX_UNSIGNED_WORD_TWIPS = 31_680;
const MIN_SIGNED_WORD_TWIPS = -2_147_483_648;
const MAX_SIGNED_WORD_TWIPS = 2_147_483_647;

export function inspectDocxPageMarginSettings(
  settings: Document | null | undefined,
): InspectedDocxPageMarginSettings {
  if (!settings) {
    return { invalidCount: 0, spoofedCount: 0, incompatible: [] };
  }
  const root = settings.documentElement;
  if (
    root.localName !== 'settings' ||
    !DOCX_WORDPROCESSING_NAMESPACES.has(root.namespaceURI ?? '')
  ) {
    return { invalidCount: 1, spoofedCount: 0, incompatible: [] };
  }
  const mirror = inspectOnOffChild(root, 'mirrorMargins');
  const gutterAtTop = inspectOnOffChild(root, 'gutterAtTop');
  const incompatible = [
    ...(mirror.value === true ? ['mirrorMargins'] : []),
    'bookFoldPrinting',
    'bookFoldRevPrinting',
    'printTwoOnOne',
  ].filter((name) => inspectOnOffChild(root, name).value === true);
  return {
    ...(mirror.value !== undefined ? { mirrorMargins: mirror.value } : {}),
    ...(gutterAtTop.value !== undefined
      ? { gutterAtTop: gutterAtTop.value }
      : {}),
    invalidCount: mirror.invalidCount + gutterAtTop.invalidCount,
    spoofedCount: mirror.spoofedCount + gutterAtTop.spoofedCount,
    incompatible,
  };
}

export function inspectDocxPageMargins(
  sectionProperties: Element,
  settings: InspectedDocxPageMarginSettings = {
    invalidCount: 0,
    spoofedCount: 0,
    incompatible: [],
  },
): InspectedDocxPageMargins {
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
  const rtlGutter = inspectOnOffChild(sectionProperties, 'rtlGutter');
  const named = directChildren(sectionProperties, 'pgMar');
  const candidates = named.filter(
    (element) => element.namespaceURI === sectionProperties.namespaceURI,
  );
  const spoofedCount =
    named.length - candidates.length + rtlGutter.spoofedCount;
  if (!candidates.length) {
    return {
      status: 'absent',
      invalidCount: rtlGutter.invalidCount,
      spoofedCount,
      ...(rtlGutter.value !== undefined
        ? { gutterOnRight: rtlGutter.value }
        : {}),
    };
  }
  if (candidates.length !== 1) {
    return {
      status: 'invalid',
      invalidCount: candidates.length + rtlGutter.invalidCount,
      spoofedCount,
    };
  }
  const element = candidates[0];
  const attributes = element ? wordPageMarginAttributes(element) : null;
  if (
    !element ||
    !attributes ||
    attributes.size !== PAGE_MARGIN_KEYS.length ||
    hasElementOrNonWhitespaceContent(element)
  ) {
    return {
      status: 'invalid',
      invalidCount: 1 + rtlGutter.invalidCount,
      spoofedCount,
    };
  }
  const values = {} as Record<WorkDocumentPageMarginKey, number>;
  for (const key of PAGE_MARGIN_KEYS) {
    const signed = key === 'top' || key === 'bottom';
    const value = parseDocxTwipsMeasure(attributes.get(key) ?? '', {
      signed,
      strict: element.namespaceURI === STRICT_WORDPROCESSING_NAMESPACE,
      minimum: signed ? MIN_SIGNED_WORD_TWIPS : 0,
      maximum: signed ? MAX_SIGNED_WORD_TWIPS : MAX_UNSIGNED_WORD_TWIPS,
    });
    if (value === null) {
      return {
        status: 'invalid',
        invalidCount: 1 + rtlGutter.invalidCount,
        spoofedCount,
      };
    }
    values[key] = value;
  }
  const pageMargins = applyPageMarginSettings(
    values,
    settings,
    rtlGutter.value,
  );
  return pageMargins
    ? {
        status: 'valid',
        pageMargins,
        invalidCount: rtlGutter.invalidCount,
        spoofedCount,
      }
    : {
        status: 'invalid',
        invalidCount: 1 + rtlGutter.invalidCount,
        spoofedCount,
      };
}

export function parseDocxPageMargins(
  sectionProperties: Element,
  settings: InspectedDocxPageMarginSettings,
  fallback?: WorkDocumentPageMargins,
): WorkDocumentPageMargins | null | undefined {
  const inspected = inspectDocxPageMargins(sectionProperties, settings);
  if (inspected.status === 'valid') return inspected.pageMargins;
  if (inspected.status === 'invalid') return null;
  const inherited = normalizeDocumentPageMargins(fallback);
  if (!inherited && !inspected.pageMargins) return undefined;
  return applyPageMarginSettings(inherited, settings, inspected.gutterOnRight);
}

function applyPageMarginSettings(
  source: Partial<Record<WorkDocumentPageMarginKey, number>> | null,
  settings: InspectedDocxPageMarginSettings,
  gutterOnRight: boolean | undefined,
): WorkDocumentPageMargins | undefined {
  if (!source) return undefined;
  const pageMargins = normalizeDocumentPageMargins({
    ...source,
    ...(settings.mirrorMargins !== undefined
      ? { mirrorMargins: settings.mirrorMargins }
      : {}),
    ...(settings.gutterAtTop !== undefined
      ? { gutterAtTop: settings.gutterAtTop }
      : {}),
    ...(gutterOnRight !== undefined ? { gutterOnRight } : {}),
  });
  return pageMargins ?? undefined;
}

function wordPageMarginAttributes(
  element: Element,
): Map<string, string> | null {
  const result = new Map<string, string>();
  for (const item of Array.from(element.attributes)) {
    if (isNamespaceDeclaration(item)) continue;
    const name = xmlAttributeLocalName(item);
    if (
      xmlAttributeNamespace(element, item) !== element.namespaceURI ||
      !PAGE_MARGIN_ATTRIBUTE_SET.has(name) ||
      result.has(name)
    ) {
      return null;
    }
    result.set(name, item.value.trim());
  }
  return result;
}

function inspectOnOffChild(
  parent: Element,
  localName: string,
): {
  value?: boolean;
  invalidCount: number;
  spoofedCount: number;
} {
  const named = directChildren(parent, localName);
  const candidates = named.filter(
    (element) => element.namespaceURI === parent.namespaceURI,
  );
  const spoofedCount = named.length - candidates.length;
  if (!candidates.length) return { invalidCount: 0, spoofedCount };
  if (candidates.length !== 1) {
    return {
      invalidCount: candidates.length,
      spoofedCount,
    };
  }
  const element = candidates[0];
  if (!element || hasElementOrNonWhitespaceContent(element)) {
    return { invalidCount: 1, spoofedCount };
  }
  let source: string | undefined;
  for (const item of Array.from(element.attributes)) {
    if (isNamespaceDeclaration(item)) continue;
    if (
      xmlAttributeNamespace(element, item) !== element.namespaceURI ||
      xmlAttributeLocalName(item) !== 'val' ||
      source !== undefined
    ) {
      return { invalidCount: 1, spoofedCount };
    }
    source = item.value.trim().toLowerCase();
  }
  if (
    source === undefined ||
    source === '1' ||
    source === 'true' ||
    source === 'on'
  ) {
    return { value: true, invalidCount: 0, spoofedCount };
  }
  if (source === '0' || source === 'false' || source === 'off') {
    return { value: false, invalidCount: 0, spoofedCount };
  }
  return { invalidCount: 1, spoofedCount };
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
