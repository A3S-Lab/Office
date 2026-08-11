import { normalizeDocumentImageIdentity } from './work-document-image-identity';
import {
  DOCUMENT_PARAGRAPH_ID_ATTRIBUTE,
  DOCUMENT_PARAGRAPH_TEXT_ID_ATTRIBUTE,
  normalizeDocumentParagraphIdentity,
} from './work-document-paragraph-identity';
import {
  DOCUMENT_TABLE_ROW_ID_ATTRIBUTE,
  DOCUMENT_TABLE_ROW_TEXT_ID_ATTRIBUTE,
  normalizeDocumentTableRowIdentity,
} from './work-document-table-row-identity';
import type {
  WorkDocumentPageChrome,
  WorkDocumentPageChromeContent,
  WorkDocumentPageChromeVariant,
  WorkDocumentSectionLayout,
} from './work-types';

interface LegacyPageChrome {
  headerText?: string;
  footerText?: string;
  showPageNumbers?: boolean;
}

export interface ResolvedDocumentPageChrome
  extends WorkDocumentPageChromeContent {
  variant: WorkDocumentPageChromeVariant;
}

const EMPTY_CONTENT: WorkDocumentPageChromeContent = {
  headerHtml: '',
  footerHtml: '',
  showPageNumber: false,
};
const IMAGE_IDENTITY_ATTRIBUTES = [
  'data-office-image-object-id',
  'data-office-image-doc-properties-id',
  'data-office-image-anchor-id',
  'data-office-image-edit-id',
] as const;
const PARAGRAPH_IDENTITY_ATTRIBUTES = [
  DOCUMENT_PARAGRAPH_ID_ATTRIBUTE,
  DOCUMENT_PARAGRAPH_TEXT_ID_ATTRIBUTE,
] as const;
const TABLE_ROW_IDENTITY_ATTRIBUTES = [
  DOCUMENT_TABLE_ROW_ID_ATTRIBUTE,
  DOCUMENT_TABLE_ROW_TEXT_ID_ATTRIBUTE,
] as const;

const ALLOWED_TAGS = new Set([
  'a',
  'b',
  'blockquote',
  'br',
  'div',
  'em',
  'i',
  'img',
  'li',
  'ol',
  'p',
  's',
  'span',
  'strike',
  'strong',
  'sub',
  'sup',
  'table',
  'tbody',
  'td',
  'th',
  'thead',
  'tr',
  'u',
  'ul',
]);

export function normalizeDocumentPageChrome(
  source?: Partial<WorkDocumentPageChrome> | null,
  legacy: LegacyPageChrome = {},
): WorkDocumentPageChrome {
  const defaultContent = normalizePageChromeContent(source?.default, {
    headerHtml: plainTextToPageChromeHtml(legacy.headerText),
    footerHtml: plainTextToPageChromeHtml(legacy.footerText),
    showPageNumber: Boolean(legacy.showPageNumbers),
  });
  return {
    differentFirstPage: Boolean(source?.differentFirstPage),
    differentOddEvenPages: Boolean(source?.differentOddEvenPages),
    default: defaultContent,
    first: normalizePageChromeContent(source?.first, EMPTY_CONTENT),
    even: normalizePageChromeContent(source?.even, EMPTY_CONTENT),
  };
}

export function serializeDocumentPageChrome(
  chrome: WorkDocumentPageChrome,
): string {
  return JSON.stringify(normalizeDocumentPageChrome(chrome));
}

export function parseDocumentPageChrome(
  source: string | undefined,
  legacy: LegacyPageChrome = {},
  fallback?: WorkDocumentPageChrome,
): WorkDocumentPageChrome {
  if (source?.trim()) {
    try {
      const parsed = JSON.parse(source) as Partial<WorkDocumentPageChrome>;
      return normalizeDocumentPageChrome(parsed, legacy);
    } catch {
      // Fall through to the compatible legacy representation.
    }
  }
  if (
    legacy.headerText ||
    legacy.footerText ||
    legacy.showPageNumbers !== undefined
  ) {
    return normalizeDocumentPageChrome(undefined, legacy);
  }
  return normalizeDocumentPageChrome(fallback);
}

export function resolveDocumentPageChrome(
  layout: WorkDocumentSectionLayout,
  sectionPage: number,
  physicalPage: number,
): ResolvedDocumentPageChrome {
  const chrome = normalizeDocumentPageChrome(layout.pageChrome, layout);
  const variant: WorkDocumentPageChromeVariant =
    chrome.differentFirstPage && sectionPage === 1
      ? 'first'
      : chrome.differentOddEvenPages && physicalPage % 2 === 0
        ? 'even'
        : 'default';
  return { variant, ...chrome[variant] };
}

export function documentPageChromeLegacyFields(
  chrome: WorkDocumentPageChrome,
): LegacyPageChrome {
  const normalized = normalizeDocumentPageChrome(chrome);
  return {
    headerText: pageChromePlainText(normalized.default.headerHtml) || undefined,
    footerText: pageChromePlainText(normalized.default.footerHtml) || undefined,
    showPageNumbers: normalized.default.showPageNumber,
  };
}

export function updateDocumentPageChromeVariant(
  chrome: WorkDocumentPageChrome,
  variant: WorkDocumentPageChromeVariant,
  patch: Partial<WorkDocumentPageChromeContent>,
): WorkDocumentPageChrome {
  const normalized = normalizeDocumentPageChrome(chrome);
  return {
    ...normalized,
    [variant]: normalizePageChromeContent(
      { ...normalized[variant], ...patch },
      EMPTY_CONTENT,
    ),
  };
}

export function pageChromePlainText(source: string): string {
  if (!source.trim()) return '';
  const document = new DOMParser().parseFromString(source, 'text/html');
  const blocks = Array.from(document.body.querySelectorAll('p, div, li'))
    .map((element) => element.textContent?.replace(/\s+/g, ' ').trim() ?? '')
    .filter(Boolean);
  return (
    blocks.length ? blocks.join('\n') : (document.body.textContent ?? '')
  ).trim();
}

export function sanitizeDocumentPageChromeHtml(
  source: string | undefined,
): string {
  if (!source?.trim()) return '';
  const document = new DOMParser().parseFromString(source, 'text/html');
  for (const element of Array.from(
    document.body.querySelectorAll('script, iframe, object, embed, link, meta'),
  )) {
    element.remove();
  }
  for (const font of Array.from(
    document.body.querySelectorAll<HTMLElement>('font'),
  )) {
    const span = document.createElement('span');
    const color = font.getAttribute('color')?.trim();
    if (font.getAttribute('style'))
      span.setAttribute('style', font.getAttribute('style') ?? '');
    if (color && !span.style.color) span.style.color = color;
    span.append(...Array.from(font.childNodes));
    font.replaceWith(span);
  }
  for (const element of Array.from(
    document.body.querySelectorAll<HTMLElement>('*'),
  )) {
    const tag = element.tagName.toLowerCase();
    if (!ALLOWED_TAGS.has(tag)) {
      element.replaceWith(...Array.from(element.childNodes));
      continue;
    }
    sanitizeAttributes(element, tag);
  }
  if (
    !(document.body.textContent ?? '').trim() &&
    !document.body.querySelector('img, table')
  )
    return '';
  return document.body.innerHTML;
}

export function plainTextToPageChromeHtml(source: string | undefined): string {
  const lines = source
    ?.split(/\r?\n/)
    .map((line) => line.trim())
    .filter(Boolean);
  if (!lines?.length) return '';
  return lines
    .map((line) => `<p style="text-align: center">${escapeHtml(line)}</p>`)
    .join('');
}

function normalizePageChromeContent(
  source: Partial<WorkDocumentPageChromeContent> | undefined,
  fallback: WorkDocumentPageChromeContent,
): WorkDocumentPageChromeContent {
  return {
    headerHtml: sanitizeDocumentPageChromeHtml(
      source?.headerHtml ?? fallback.headerHtml,
    ),
    footerHtml: sanitizeDocumentPageChromeHtml(
      source?.footerHtml ?? fallback.footerHtml,
    ),
    showPageNumber: source?.showPageNumber ?? fallback.showPageNumber,
  };
}

function sanitizeAttributes(element: HTMLElement, tag: string) {
  for (const attribute of Array.from(element.attributes)) {
    if (attribute.name.toLowerCase().startsWith('on'))
      element.removeAttribute(attribute.name);
  }
  const textAlign = ['left', 'center', 'right', 'justify'].includes(
    element.style.textAlign,
  )
    ? element.style.textAlign
    : '';
  const color = element.style.color;
  const direction = element.getAttribute('dir')?.trim().toLowerCase();
  element.removeAttribute('style');
  const styles = [
    textAlign ? `text-align: ${textAlign}` : '',
    color ? `color: ${color}` : '',
  ].filter(Boolean);
  if (styles.length) element.setAttribute('style', styles.join('; '));

  if (tag === 'a') {
    const href = element.getAttribute('href')?.trim() ?? '';
    if (!/^(?:https?:|mailto:|#)/i.test(href)) element.removeAttribute('href');
  } else if (tag === 'img') {
    const source = element.getAttribute('src')?.trim() ?? '';
    if (!/^(?:https?:|blob:|data:image\/)/i.test(source))
      element.removeAttribute('src');
    normalizeImageIdentityAttributes(element);
  } else if (tag === 'ol') {
    const start = Number(element.getAttribute('start'));
    if (!Number.isSafeInteger(start) || start <= 0)
      element.removeAttribute('start');
    const type = element.getAttribute('type');
    if (type && !['1', 'A', 'a', 'I', 'i'].includes(type))
      element.removeAttribute('type');
  }
  if (tag === 'p') normalizeParagraphIdentityAttributes(element);
  if (tag === 'tr') normalizeTableRowIdentityAttributes(element);
  if (direction === 'ltr' || direction === 'rtl')
    element.setAttribute('dir', direction);
  else element.removeAttribute('dir');
  const allowed =
    tag === 'a'
      ? new Set(['dir', 'href', 'title', 'style'])
      : tag === 'img'
        ? new Set([
            'dir',
            'src',
            'alt',
            'title',
            'width',
            'height',
            'style',
            ...IMAGE_IDENTITY_ATTRIBUTES,
          ])
        : tag === 'ol'
          ? new Set(['dir', 'start', 'style', 'type'])
          : tag === 'p'
            ? new Set(['dir', 'style', ...PARAGRAPH_IDENTITY_ATTRIBUTES])
            : tag === 'tr'
              ? new Set(['dir', 'style', ...TABLE_ROW_IDENTITY_ATTRIBUTES])
              : new Set(['colspan', 'dir', 'rowspan', 'style']);
  for (const attribute of Array.from(element.attributes)) {
    if (!allowed.has(attribute.name.toLowerCase()))
      element.removeAttribute(attribute.name);
  }
}

function normalizeTableRowIdentityAttributes(element: HTMLElement): void {
  const identity = normalizeDocumentTableRowIdentity({
    rowId: element.getAttribute(TABLE_ROW_IDENTITY_ATTRIBUTES[0]),
    rowTextId: element.getAttribute(TABLE_ROW_IDENTITY_ATTRIBUTES[1]),
  });
  if (!identity) {
    for (const name of TABLE_ROW_IDENTITY_ATTRIBUTES) {
      element.removeAttribute(name);
    }
    return;
  }
  element.setAttribute(TABLE_ROW_IDENTITY_ATTRIBUTES[0], identity.rowId);
  element.setAttribute(TABLE_ROW_IDENTITY_ATTRIBUTES[1], identity.rowTextId);
}

function normalizeParagraphIdentityAttributes(element: HTMLElement): void {
  const identity = normalizeDocumentParagraphIdentity({
    paragraphId: element.getAttribute(PARAGRAPH_IDENTITY_ATTRIBUTES[0]),
    textId: element.getAttribute(PARAGRAPH_IDENTITY_ATTRIBUTES[1]),
  });
  if (!identity) {
    for (const name of PARAGRAPH_IDENTITY_ATTRIBUTES) {
      element.removeAttribute(name);
    }
    return;
  }
  element.setAttribute(PARAGRAPH_IDENTITY_ATTRIBUTES[0], identity.paragraphId);
  element.setAttribute(PARAGRAPH_IDENTITY_ATTRIBUTES[1], identity.textId);
}

function normalizeImageIdentityAttributes(element: HTMLElement): void {
  const identity = normalizeDocumentImageIdentity({
    objectId: element.getAttribute(IMAGE_IDENTITY_ATTRIBUTES[0]),
    docPropertiesId: element.getAttribute(IMAGE_IDENTITY_ATTRIBUTES[1]),
    anchorId: element.getAttribute(IMAGE_IDENTITY_ATTRIBUTES[2]),
    editId: element.getAttribute(IMAGE_IDENTITY_ATTRIBUTES[3]),
  });
  if (!identity) {
    for (const name of IMAGE_IDENTITY_ATTRIBUTES) element.removeAttribute(name);
    return;
  }
  element.setAttribute(IMAGE_IDENTITY_ATTRIBUTES[0], identity.objectId);
  element.setAttribute(
    IMAGE_IDENTITY_ATTRIBUTES[1],
    String(identity.docPropertiesId),
  );
  element.setAttribute(IMAGE_IDENTITY_ATTRIBUTES[2], identity.anchorId);
  element.setAttribute(IMAGE_IDENTITY_ATTRIBUTES[3], identity.editId);
}

function escapeHtml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');
}
