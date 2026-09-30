import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import { stableJson } from './work-document-compare-stability';
import {
  pageChromePlainText,
  parseDocumentPageChrome,
  serializeDocumentPageChrome,
  updateDocumentPageChromeVariant,
} from './work-document-page-chrome';
import type {
  WorkDocumentPageChrome,
  WorkDocumentPageChromeVariant,
} from './work-types';

const SECTION_LAYOUT_IGNORED_ATTRIBUTES = new Set([
  'id',
  'headerText',
  'footerText',
  'propertyRevisionOmml',
  'sectionChangeKind',
  'sectionChangeId',
  'sectionChangeAuthor',
  'sectionChangeDate',
  'sectionChangeBefore',
]);

const PAGE_CHROME_VARIANTS = ['default', 'first', 'even'] as const;
const PAGE_CHROME_PARTS = ['header', 'footer'] as const;

type PageChromePart = (typeof PAGE_CHROME_PARTS)[number];

/**
 * Page geometry only. Header and footer prose, and leftover section-revision
 * metadata, are compared as content or ignored so a one-word copy is not
 * refused as 「第 1 节」 layout.
 */
export function sectionLayoutSignature(section: ProseMirrorNode): string {
  const attributes = filteredAttributes(
    section.attrs,
    SECTION_LAYOUT_IGNORED_ATTRIBUTES,
  );
  if (typeof attributes.pageChrome === 'string') {
    attributes.pageChrome = pageChromeLayoutSignature(attributes.pageChrome);
  }
  return stableJson(attributes);
}

export function comparedPageChromeAttributes(
  current: ProseMirrorNode,
  revised: ProseMirrorNode,
  renderRevisionHtml: (
    currentText: string,
    revisedText: string,
  ) => string | null,
): Record<string, unknown> {
  let chrome = sectionPageChrome(current);
  const patch: Record<string, unknown> = {};
  let changed = false;
  for (const variant of PAGE_CHROME_VARIANTS) {
    for (const part of PAGE_CHROME_PARTS) {
      const currentText = chromePartText(current, chrome, variant, part);
      const revisedText = chromePartText(
        revised,
        sectionPageChrome(revised),
        variant,
        part,
      );
      if (currentText === revisedText) continue;
      const html = renderRevisionHtml(currentText, revisedText);
      if (!html) continue;
      chrome = updateDocumentPageChromeVariant(chrome, variant, {
        [part === 'header' ? 'headerHtml' : 'footerHtml']: html,
      });
      changed = true;
      if (variant === 'default') {
        patch[part === 'header' ? 'headerText' : 'footerText'] = revisedText;
      }
    }
  }
  if (!changed) return patch;
  return {
    ...patch,
    pageChrome: serializeDocumentPageChrome(chrome),
  };
}

function pageChromeLayoutSignature(serialized: string): string {
  if (!serialized.trim()) return '';
  try {
    const parsed = JSON.parse(serialized) as Partial<WorkDocumentPageChrome>;
    const showPageNumber = (
      content: WorkDocumentPageChrome['default'] | undefined,
    ) => Boolean(content?.showPageNumber);
    return stableJson({
      differentFirstPage: Boolean(parsed.differentFirstPage),
      differentOddEvenPages: Boolean(parsed.differentOddEvenPages),
      showPageNumber: {
        default: showPageNumber(parsed.default),
        even: showPageNumber(parsed.even),
        first: showPageNumber(parsed.first),
      },
    });
  } catch {
    return serialized;
  }
}

function sectionPageChrome(section: ProseMirrorNode): WorkDocumentPageChrome {
  return parseDocumentPageChrome(
    typeof section.attrs.pageChrome === 'string'
      ? section.attrs.pageChrome
      : '',
    {
      headerText: stringAttribute(section.attrs.headerText),
      footerText: stringAttribute(section.attrs.footerText),
      showPageNumbers: section.attrs.showPageNumbers === true,
    },
  );
}

function chromePartText(
  section: ProseMirrorNode,
  chrome: WorkDocumentPageChrome,
  variant: WorkDocumentPageChromeVariant,
  part: PageChromePart,
): string {
  const html = chrome[variant][part === 'header' ? 'headerHtml' : 'footerHtml'];
  const fromHtml = pageChromePlainText(html);
  if (fromHtml) return fromHtml;
  if (variant !== 'default') return '';
  return stringAttribute(
    section.attrs[part === 'header' ? 'headerText' : 'footerText'],
  );
}

function stringAttribute(value: unknown): string {
  return typeof value === 'string' ? value : '';
}

function filteredAttributes(
  attributes: Record<string, unknown>,
  ignored: ReadonlySet<string>,
): Record<string, unknown> {
  return Object.fromEntries(
    Object.entries(attributes).filter(([name]) => !ignored.has(name)),
  );
}
