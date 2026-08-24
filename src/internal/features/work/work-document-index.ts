import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import type { WorkDocumentFieldContextResolver } from './work-document-fields';

export type WorkDocumentIndexLeader = 'dot' | 'dash' | 'underline' | 'none';
export type WorkDocumentIndexFormat = 'indented' | 'run-in';

export interface WorkDocumentIndexEntryDraft {
  mainEntry: string;
  subEntry: string;
  crossReference: string;
  pageBold: boolean;
  pageItalic: boolean;
}

export interface WorkDocumentIndexEntry extends WorkDocumentIndexEntryDraft {
  id: string;
}

export interface WorkDocumentIndexPage {
  pageNumber: number;
  pageBold: boolean;
  pageItalic: boolean;
  targetIds: string[];
}

export interface WorkDocumentIndexGeneratedEntry {
  mainEntry: string;
  subEntry: string;
  crossReference: string;
  pages: WorkDocumentIndexPage[];
}

export interface WorkDocumentIndexOptions {
  columns: number;
  format: WorkDocumentIndexFormat;
  rightAlignPageNumbers: boolean;
  leader: WorkDocumentIndexLeader;
}

export interface WorkDocumentIndexValue {
  id: string;
  options: WorkDocumentIndexOptions;
  entries: WorkDocumentIndexGeneratedEntry[];
  truncated: boolean;
}

export interface WorkDocumentIndexBuildOptions {
  resolveContext?: WorkDocumentFieldContextResolver;
}

export const DEFAULT_DOCUMENT_INDEX_OPTIONS: WorkDocumentIndexOptions = {
  columns: 1,
  format: 'indented',
  rightAlignPageNumbers: true,
  leader: 'dot',
};

export const MAX_DOCUMENT_INDEX_ENTRIES = 512;
export const MAX_DOCUMENT_INDEX_MARKERS = 2_048;
export const MAX_DOCUMENT_INDEX_TERM_LENGTH = 240;

const INDEX_ENTRY_SELECTOR = '[data-document-index-entry]';
const INDEX_SELECTOR = '[data-document-index]';
const INDEX_ID_PATTERN = /^[a-z0-9][a-z0-9._:-]{0,127}$/i;

export function normalizeDocumentIndexEntryDraft(
  source: Partial<WorkDocumentIndexEntryDraft> | null | undefined,
): WorkDocumentIndexEntryDraft | null {
  const mainEntry = normalizedIndexTerm(source?.mainEntry);
  if (!mainEntry) return null;
  const subEntry = normalizedIndexTerm(source?.subEntry);
  const crossReference = normalizedIndexTerm(source?.crossReference);
  return {
    mainEntry,
    subEntry,
    crossReference,
    pageBold: !crossReference && Boolean(source?.pageBold),
    pageItalic: !crossReference && Boolean(source?.pageItalic),
  };
}

export function normalizeDocumentIndexEntry(
  source: Partial<WorkDocumentIndexEntry> | null | undefined,
  fallbackId = 'index-entry',
): WorkDocumentIndexEntry | null {
  const value = normalizeDocumentIndexEntryDraft(source);
  if (!value) return null;
  return {
    id: validIndexId(source?.id) ?? fallbackId,
    ...value,
  };
}

export function normalizeDocumentIndexOptions(
  source: Partial<WorkDocumentIndexOptions> | null | undefined,
): WorkDocumentIndexOptions {
  return {
    columns: boundedColumns(source?.columns),
    format: source?.format === 'run-in' ? 'run-in' : 'indented',
    rightAlignPageNumbers: source?.rightAlignPageNumbers !== false,
    leader: indexLeader(source?.leader) ?? 'dot',
  };
}

export function normalizeDocumentIndexValue(
  source: Omit<WorkDocumentIndexValue, 'truncated'> & { truncated?: boolean },
): WorkDocumentIndexValue {
  return {
    id: validIndexId(source.id) ?? 'document-index',
    options: normalizeDocumentIndexOptions(source.options),
    entries: normalizeGeneratedEntries(source.entries),
    truncated: Boolean(source.truncated),
  };
}

export function buildDocumentIndexEntries(
  document: ProseMirrorNode,
  options: WorkDocumentIndexBuildOptions = {},
): { entries: WorkDocumentIndexGeneratedEntry[]; truncated: boolean } {
  const grouped = new Map<
    string,
    {
      entry: WorkDocumentIndexGeneratedEntry;
      pages: Map<number, WorkDocumentIndexPage>;
    }
  >();
  let markerCount = 0;
  document.descendants((node, position) => {
    if (node.type.name !== 'documentIndexEntry') return;
    markerCount += 1;
    if (markerCount > MAX_DOCUMENT_INDEX_MARKERS) return;
    const marker = normalizeDocumentIndexEntry(
      node.attrs,
      `index-entry-${markerCount}`,
    );
    if (!marker) return;
    const key = indexEntryKey(marker);
    let group = grouped.get(key);
    if (!group) {
      group = {
        entry: {
          mainEntry: marker.mainEntry,
          subEntry: marker.subEntry,
          crossReference: marker.crossReference,
          pages: [],
        },
        pages: new Map(),
      };
      grouped.set(key, group);
    }
    if (marker.crossReference) return;
    const pageNumber =
      positiveInteger(options.resolveContext?.(position)?.pageNumber) ??
      fallbackDocumentPageNumber(document, position);
    const existing = group.pages.get(pageNumber);
    if (existing) {
      existing.pageBold ||= marker.pageBold;
      existing.pageItalic ||= marker.pageItalic;
      if (!existing.targetIds.includes(marker.id)) {
        existing.targetIds.push(marker.id);
      }
      return;
    }
    group.pages.set(pageNumber, {
      pageNumber,
      pageBold: marker.pageBold,
      pageItalic: marker.pageItalic,
      targetIds: [marker.id],
    });
  });

  const collator = new Intl.Collator(undefined, {
    numeric: true,
    sensitivity: 'base',
    usage: 'sort',
  });
  const allEntries = Array.from(grouped.values())
    .map(({ entry, pages }) => ({
      ...entry,
      pages: Array.from(pages.values()).sort(
        (left, right) => left.pageNumber - right.pageNumber,
      ),
    }))
    .sort(
      (left, right) =>
        collator.compare(left.mainEntry, right.mainEntry) ||
        collator.compare(left.subEntry, right.subEntry) ||
        collator.compare(left.crossReference, right.crossReference),
    );
  return {
    entries: allEntries.slice(0, MAX_DOCUMENT_INDEX_ENTRIES),
    truncated:
      markerCount > MAX_DOCUMENT_INDEX_MARKERS ||
      allEntries.length > MAX_DOCUMENT_INDEX_ENTRIES,
  };
}

export function documentIndexEntryHtml(
  source: Partial<WorkDocumentIndexEntry>,
): string {
  const value = normalizeDocumentIndexEntry(source);
  if (!value) return '';
  const detail = indexEntryDisplay(value);
  return [
    `<span data-document-index-entry="true" data-index-entry-id="${escapeHtmlAttribute(value.id)}" data-index-main-entry="${escapeHtmlAttribute(value.mainEntry)}" data-index-sub-entry="${escapeHtmlAttribute(value.subEntry)}" data-index-cross-reference="${escapeHtmlAttribute(value.crossReference)}" data-index-page-bold="${String(value.pageBold)}" data-index-page-italic="${String(value.pageItalic)}" class="work-document-index-entry" contenteditable="false" aria-label="索引项：${escapeHtmlAttribute(detail)}">`,
    '<span aria-hidden="true">索引项</span>',
    `<strong>${escapeHtml(detail)}</strong>`,
    '</span>',
  ].join('');
}

export function documentIndexHtml(
  source: Omit<WorkDocumentIndexValue, 'truncated'> & { truncated?: boolean },
): string {
  const value = normalizeDocumentIndexValue(source);
  const rows = value.entries.length
    ? value.entries.map(documentIndexRowHtml)
    : ['<li class="work-document-index-empty">没有已标记的索引项</li>'];
  const status = value.truncated
    ? `仅显示前 ${MAX_DOCUMENT_INDEX_ENTRIES} 项`
    : `${value.entries.length} 项`;
  return [
    `<div data-document-index="true" data-index-id="${escapeHtmlAttribute(value.id)}" data-index-columns="${value.options.columns}" data-index-format="${value.options.format}" data-index-right-align-page-numbers="${String(value.options.rightAlignPageNumbers)}" data-index-leader="${value.options.leader}" data-index-entries="${escapeHtmlAttribute(JSON.stringify(value.entries))}" data-index-truncated="${String(value.truncated)}" class="work-document-index" contenteditable="false" aria-label="索引">`,
    `<div class="work-document-index-header"><strong>索引</strong><span>${status}</span></div>`,
    '<ol class="work-document-index-list">',
    ...rows,
    '</ol></div>',
  ].join('');
}

export function documentIndexEntryFromElement(
  element: HTMLElement,
  index = 1,
): WorkDocumentIndexEntry | null {
  return normalizeDocumentIndexEntry(
    {
      id: element.dataset.indexEntryId,
      mainEntry: element.dataset.indexMainEntry,
      subEntry: element.dataset.indexSubEntry,
      crossReference: element.dataset.indexCrossReference,
      pageBold: element.dataset.indexPageBold === 'true',
      pageItalic: element.dataset.indexPageItalic === 'true',
    },
    `index-entry-${index}`,
  );
}

export function documentIndexValueFromElement(
  element: HTMLElement,
  index = 1,
): WorkDocumentIndexValue {
  return normalizeDocumentIndexValue({
    id: validIndexId(element.dataset.indexId) ?? `document-index-${index}`,
    options: {
      columns: Number(element.dataset.indexColumns),
      format: element.dataset.indexFormat === 'run-in' ? 'run-in' : 'indented',
      rightAlignPageNumbers:
        element.dataset.indexRightAlignPageNumbers !== 'false',
      leader: indexLeader(element.dataset.indexLeader) ?? 'dot',
    },
    entries: parseGeneratedEntries(element.dataset.indexEntries),
    truncated: element.dataset.indexTruncated === 'true',
  });
}

export function normalizeDocumentIndexesHtml(source: string): string {
  const document = new DOMParser().parseFromString(source, 'text/html');
  const usedEntryIds = new Set<string>();
  for (const [index, element] of Array.from(
    document.body.querySelectorAll<HTMLElement>(INDEX_ENTRY_SELECTOR),
  ).entries()) {
    const value = documentIndexEntryFromElement(element, index + 1);
    if (!value) {
      element.remove();
      continue;
    }
    value.id = uniqueIndexId(value.id, 'index-entry', index + 1, usedEntryIds);
    element.replaceWith(
      document
        .createRange()
        .createContextualFragment(documentIndexEntryHtml(value)),
    );
  }
  const usedIndexIds = new Set<string>();
  for (const [index, element] of Array.from(
    document.body.querySelectorAll<HTMLElement>(INDEX_SELECTOR),
  ).entries()) {
    const value = documentIndexValueFromElement(element, index + 1);
    value.id = uniqueIndexId(
      value.id,
      'document-index',
      index + 1,
      usedIndexIds,
    );
    element.replaceWith(
      document.createRange().createContextualFragment(documentIndexHtml(value)),
    );
  }
  return document.body.innerHTML;
}

export function indexLeader(value: unknown): WorkDocumentIndexLeader | null {
  return value === 'dot' ||
    value === 'dash' ||
    value === 'underline' ||
    value === 'none'
    ? value
    : null;
}

function documentIndexRowHtml(entry: WorkDocumentIndexGeneratedEntry): string {
  const term = entry.subEntry
    ? `<span class="work-document-index-main">${escapeHtml(entry.mainEntry)}</span><span class="work-document-index-sub">${escapeHtml(entry.subEntry)}</span>`
    : `<span class="work-document-index-main">${escapeHtml(entry.mainEntry)}</span>`;
  const pages = entry.crossReference
    ? `<span class="work-document-index-cross-reference">参见 ${escapeHtml(entry.crossReference)}</span>`
    : entry.pages
        .map((page) => {
          const target = page.targetIds[0] ?? '';
          const classes = [
            page.pageBold ? 'bold' : '',
            page.pageItalic ? 'italic' : '',
          ]
            .filter(Boolean)
            .join(' ');
          return `<a href="#${escapeHtmlAttribute(target)}" data-index-target="${escapeHtmlAttribute(target)}" tabindex="-1"${classes ? ` class="${classes}"` : ''}>${page.pageNumber}</a>`;
        })
        .join('<span aria-hidden="true">, </span>');
  return `<li data-index-main-entry="${escapeHtmlAttribute(entry.mainEntry)}" data-index-sub-entry="${escapeHtmlAttribute(entry.subEntry)}">${term}<i aria-hidden="true"></i><span class="work-document-index-pages">${pages}</span></li>`;
}

function normalizeGeneratedEntries(
  source: readonly unknown[],
): WorkDocumentIndexGeneratedEntry[] {
  const entries: WorkDocumentIndexGeneratedEntry[] = [];
  for (const item of source.slice(0, MAX_DOCUMENT_INDEX_ENTRIES)) {
    if (!item || typeof item !== 'object') continue;
    const candidate = item as Record<string, unknown>;
    const draft = normalizeDocumentIndexEntryDraft({
      mainEntry: normalizedIndexTerm(candidate.mainEntry),
      subEntry: normalizedIndexTerm(candidate.subEntry),
      crossReference: normalizedIndexTerm(candidate.crossReference),
    });
    if (!draft) continue;
    entries.push({
      mainEntry: draft.mainEntry,
      subEntry: draft.subEntry,
      crossReference: draft.crossReference,
      pages: draft.crossReference ? [] : normalizeIndexPages(candidate.pages),
    });
  }
  return entries;
}

function normalizeIndexPages(source: unknown): WorkDocumentIndexPage[] {
  if (!Array.isArray(source)) return [];
  const pages: WorkDocumentIndexPage[] = [];
  const seen = new Set<number>();
  for (const item of source.slice(0, MAX_DOCUMENT_INDEX_MARKERS)) {
    if (!item || typeof item !== 'object') continue;
    const candidate = item as Record<string, unknown>;
    const pageNumber = positiveInteger(candidate.pageNumber);
    if (!pageNumber || seen.has(pageNumber)) continue;
    const targetIds = Array.isArray(candidate.targetIds)
      ? candidate.targetIds
          .flatMap((value) => (validIndexId(value) ? [String(value)] : []))
          .slice(0, 64)
      : [];
    if (!targetIds.length) continue;
    seen.add(pageNumber);
    pages.push({
      pageNumber,
      pageBold: Boolean(candidate.pageBold),
      pageItalic: Boolean(candidate.pageItalic),
      targetIds,
    });
  }
  return pages.sort((left, right) => left.pageNumber - right.pageNumber);
}

function parseGeneratedEntries(
  source: string | undefined,
): WorkDocumentIndexGeneratedEntry[] {
  if (!source) return [];
  try {
    const parsed: unknown = JSON.parse(source);
    return Array.isArray(parsed) ? normalizeGeneratedEntries(parsed) : [];
  } catch {
    return [];
  }
}

function fallbackDocumentPageNumber(
  document: ProseMirrorNode,
  position: number,
): number {
  let pageNumber = 1;
  document.descendants((node, offset) => {
    if (offset >= position) return false;
    if (node.type.name === 'pageBreak') pageNumber += 1;
    return true;
  });
  return pageNumber;
}

function indexEntryKey(entry: WorkDocumentIndexEntryDraft): string {
  return [entry.mainEntry, entry.subEntry, entry.crossReference]
    .map((value) => value.normalize('NFKC').toLocaleLowerCase())
    .join('\u0000');
}

function indexEntryDisplay(entry: WorkDocumentIndexEntryDraft): string {
  const term = entry.subEntry
    ? `${entry.mainEntry} › ${entry.subEntry}`
    : entry.mainEntry;
  return entry.crossReference ? `${term} · 参见 ${entry.crossReference}` : term;
}

function normalizedIndexTerm(value: unknown): string {
  return typeof value === 'string'
    ? value.replace(/\s+/g, ' ').trim().slice(0, MAX_DOCUMENT_INDEX_TERM_LENGTH)
    : '';
}

function boundedColumns(value: unknown): number {
  const number = Number(value);
  return Number.isInteger(number) && number >= 1 && number <= 4 ? number : 1;
}

function positiveInteger(value: unknown): number | null {
  const number = Number(value);
  return Number.isSafeInteger(number) && number > 0
    ? Math.min(999_999, number)
    : null;
}

function validIndexId(value: unknown): string | null {
  return typeof value === 'string' && INDEX_ID_PATTERN.test(value)
    ? value
    : null;
}

function uniqueIndexId(
  source: string,
  prefix: string,
  index: number,
  used: Set<string>,
): string {
  if (!used.has(source)) {
    used.add(source);
    return source;
  }
  let suffix = index;
  while (used.has(`${prefix}-${suffix}`)) suffix += 1;
  const id = `${prefix}-${suffix}`;
  used.add(id);
  return id;
}

function escapeHtml(value: string): string {
  return value
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;');
}

function escapeHtmlAttribute(value: string): string {
  return escapeHtml(value).replaceAll('"', '&quot;').replaceAll("'", '&#39;');
}
