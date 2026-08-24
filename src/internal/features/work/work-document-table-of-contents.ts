import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import {
  collectWorkDocumentOutline,
  type WorkDocumentOutlineItem,
} from './work-document-outline';
import type { WorkDocumentFieldContextResolver } from './work-document-fields';

export type WorkDocumentTableOfContentsLeader =
  | 'dot'
  | 'dash'
  | 'underline'
  | 'none';

export interface WorkDocumentTableOfContentsOptions {
  minLevel: number;
  maxLevel: number;
  hyperlinks: boolean;
  showPageNumbers: boolean;
  rightAlignPageNumbers: boolean;
  leader: WorkDocumentTableOfContentsLeader;
}

export interface WorkDocumentTableOfContentsEntry {
  targetId: string;
  title: string;
  level: number;
  pageNumber: number;
}

export interface WorkDocumentTableOfContentsValue {
  id: string;
  options: WorkDocumentTableOfContentsOptions;
  entries: WorkDocumentTableOfContentsEntry[];
  truncated: boolean;
}

export interface WorkDocumentTableOfContentsBuildOptions {
  resolveContext?: WorkDocumentFieldContextResolver;
}

export type WorkDocumentTableOfContentsInstructionResult =
  | {
      supported: true;
      options: WorkDocumentTableOfContentsOptions;
    }
  | {
      supported: false;
      reason:
        | 'not-table-of-contents'
        | 'invalid-instruction'
        | 'invalid-level-range'
        | 'unsupported-switch'
        | 'unsupported-page-range'
        | 'unsupported-separator';
    };

export const DEFAULT_DOCUMENT_TABLE_OF_CONTENTS_OPTIONS: WorkDocumentTableOfContentsOptions =
  {
    minLevel: 1,
    maxLevel: 3,
    hyperlinks: true,
    showPageNumbers: true,
    rightAlignPageNumbers: true,
    leader: 'dot',
  };

export const MAX_DOCUMENT_TABLE_OF_CONTENTS_ENTRIES = 512;

const TABLE_OF_CONTENTS_SELECTOR = '[data-document-table-of-contents]';
const MAX_TABLE_OF_CONTENTS_TITLE_LENGTH = 512;
const TABLE_OF_CONTENTS_ID_PATTERN = /^[a-z0-9][a-z0-9._:-]{0,127}$/i;
const TABLE_OF_CONTENTS_TARGET_PATTERN = /^heading-[a-z0-9._:-]{1,128}$/i;

export function normalizeDocumentTableOfContentsOptions(
  source: Partial<WorkDocumentTableOfContentsOptions> | null | undefined,
): WorkDocumentTableOfContentsOptions {
  const minLevel = boundedLevel(source?.minLevel, 1);
  const maxLevel = Math.max(
    minLevel,
    boundedLevel(source?.maxLevel, Math.max(3, minLevel)),
  );
  const showPageNumbers = source?.showPageNumbers !== false;
  return {
    minLevel,
    maxLevel,
    hyperlinks: source?.hyperlinks !== false,
    showPageNumbers,
    rightAlignPageNumbers:
      showPageNumbers && source?.rightAlignPageNumbers !== false,
    leader: tableOfContentsLeader(source?.leader) ?? 'dot',
  };
}

export function buildDocumentTableOfContentsEntries(
  document: ProseMirrorNode,
  options: WorkDocumentTableOfContentsOptions,
  buildOptions: WorkDocumentTableOfContentsBuildOptions = {},
): { entries: WorkDocumentTableOfContentsEntry[]; truncated: boolean } {
  const normalized = normalizeDocumentTableOfContentsOptions(options);
  const outline = collectWorkDocumentOutline(document).filter(
    (item) =>
      item.text.length > 0 &&
      item.level >= normalized.minLevel &&
      item.level <= normalized.maxLevel,
  );
  const entries = outline
    .slice(0, MAX_DOCUMENT_TABLE_OF_CONTENTS_ENTRIES)
    .map((item) =>
      tableOfContentsEntryFromOutline(
        document,
        item,
        buildOptions.resolveContext,
      ),
    );
  return {
    entries,
    truncated: outline.length > MAX_DOCUMENT_TABLE_OF_CONTENTS_ENTRIES,
  };
}

export function documentTableOfContentsHtml(
  source: Omit<WorkDocumentTableOfContentsValue, 'truncated'> & {
    truncated?: boolean;
  },
): string {
  const value = normalizeDocumentTableOfContentsValue(source);
  const options = value.options;
  const rows = value.entries.length
    ? value.entries.map((entry) => tableOfContentsEntryHtml(entry, options))
    : [
        '<li class="work-document-table-of-contents-empty">没有符合级别范围的标题</li>',
      ];
  const status = value.truncated
    ? `仅显示前 ${MAX_DOCUMENT_TABLE_OF_CONTENTS_ENTRIES} 项`
    : `${value.entries.length} 项`;
  return [
    `<div data-document-table-of-contents="true" data-toc-id="${escapeHtmlAttribute(value.id)}" data-toc-min-level="${options.minLevel}" data-toc-max-level="${options.maxLevel}" data-toc-hyperlinks="${String(options.hyperlinks)}" data-toc-show-page-numbers="${String(options.showPageNumbers)}" data-toc-right-align-page-numbers="${String(options.rightAlignPageNumbers)}" data-toc-leader="${options.leader}" data-toc-entries="${escapeHtmlAttribute(JSON.stringify(value.entries))}" data-toc-truncated="${String(value.truncated)}" class="work-document-table-of-contents" contenteditable="false" aria-label="目录">`,
    '<div class="work-document-table-of-contents-header"><strong>目录</strong>',
    `<span>${status}</span></div>`,
    '<ol class="work-document-table-of-contents-list">',
    ...rows,
    '</ol></div>',
  ].join('');
}

export function normalizeDocumentTableOfContentsHtml(source: string): string {
  const document = new DOMParser().parseFromString(source, 'text/html');
  const usedIds = new Set<string>();
  for (const [index, element] of Array.from(
    document.body.querySelectorAll<HTMLElement>(TABLE_OF_CONTENTS_SELECTOR),
  ).entries()) {
    const value = documentTableOfContentsValueFromElement(element, index + 1);
    value.id = uniqueTableOfContentsId(value.id, index + 1, usedIds);
    const replacement = document
      .createRange()
      .createContextualFragment(documentTableOfContentsHtml(value));
    element.replaceWith(replacement);
  }
  return document.body.innerHTML;
}

export function documentTableOfContentsValueFromElement(
  element: HTMLElement,
  index = 1,
): WorkDocumentTableOfContentsValue {
  const entries = parseTableOfContentsEntries(element.dataset.tocEntries);
  return normalizeDocumentTableOfContentsValue({
    id: validTableOfContentsId(element.dataset.tocId) ?? `toc-${index}`,
    options: {
      minLevel: Number(element.dataset.tocMinLevel),
      maxLevel: Number(element.dataset.tocMaxLevel),
      hyperlinks: element.dataset.tocHyperlinks !== 'false',
      showPageNumbers: element.dataset.tocShowPageNumbers !== 'false',
      rightAlignPageNumbers:
        element.dataset.tocRightAlignPageNumbers !== 'false',
      leader: tableOfContentsLeader(element.dataset.tocLeader) ?? 'dot',
    },
    entries,
    truncated: element.dataset.tocTruncated === 'true',
  });
}

export function parseDocumentTableOfContentsInstruction(
  instruction: string,
): WorkDocumentTableOfContentsInstructionResult {
  const command = /^\s*TOC\b/i.exec(instruction);
  if (!command) return { supported: false, reason: 'not-table-of-contents' };
  const source = instruction.slice(command[0].length);
  const switches = new Map<string, string | null>();
  const matcher = /\\([a-z])(?:\s+("([^"]*)"|([^\\\s]+)))?/gi;
  let cursor = 0;
  for (let match = matcher.exec(source); match; match = matcher.exec(source)) {
    if (source.slice(cursor, match.index).trim()) {
      return { supported: false, reason: 'invalid-instruction' };
    }
    const name = match[1]?.toLowerCase() ?? '';
    if (!name || switches.has(name)) {
      return { supported: false, reason: 'invalid-instruction' };
    }
    switches.set(name, match[3] ?? match[4] ?? null);
    cursor = matcher.lastIndex;
  }
  if (source.slice(cursor).trim()) {
    return { supported: false, reason: 'invalid-instruction' };
  }
  for (const name of switches.keys()) {
    if (!['o', 'h', 'z', 'u', 'n', 'p'].includes(name)) {
      return { supported: false, reason: 'unsupported-switch' };
    }
  }

  const range = switches.has('o')
    ? parseTableOfContentsRange(switches.get('o'))
    : { minLevel: 1, maxLevel: 3 };
  if (!range) return { supported: false, reason: 'invalid-level-range' };
  const pageRange = switches.has('n')
    ? parseOptionalTableOfContentsRange(switches.get('n'))
    : null;
  if (switches.has('n') && pageRange === undefined) {
    return { supported: false, reason: 'invalid-level-range' };
  }
  if (
    pageRange &&
    (pageRange.minLevel !== range.minLevel ||
      pageRange.maxLevel !== range.maxLevel)
  ) {
    return { supported: false, reason: 'unsupported-page-range' };
  }
  const separator = switches.get('p');
  if (
    switches.has('p') &&
    (separator === null || separator === undefined || separator.trim())
  ) {
    return { supported: false, reason: 'unsupported-separator' };
  }
  const showPageNumbers = !switches.has('n');
  const rightAlignPageNumbers = showPageNumbers && !switches.has('p');
  return {
    supported: true,
    options: {
      ...range,
      hyperlinks: switches.has('h'),
      showPageNumbers,
      rightAlignPageNumbers,
      leader: rightAlignPageNumbers ? 'dot' : 'none',
    },
  };
}

export function documentTableOfContentsInstruction(
  source: WorkDocumentTableOfContentsOptions,
): string {
  const options = normalizeDocumentTableOfContentsOptions(source);
  const switches = [`\\o "${options.minLevel}-${options.maxLevel}"`];
  if (options.hyperlinks) switches.push('\\h');
  if (!options.showPageNumbers) {
    switches.push(`\\n "${options.minLevel}-${options.maxLevel}"`);
  } else if (!options.rightAlignPageNumbers) {
    switches.push('\\p " "');
  }
  switches.push('\\z', '\\u');
  return `TOC ${switches.join(' ')}`;
}

export function tableOfContentsLeader(
  value: unknown,
): WorkDocumentTableOfContentsLeader | null {
  return value === 'dot' ||
    value === 'dash' ||
    value === 'underline' ||
    value === 'none'
    ? value
    : null;
}

export function normalizeDocumentTableOfContentsValue(
  source: Omit<WorkDocumentTableOfContentsValue, 'truncated'> & {
    truncated?: boolean;
  },
): WorkDocumentTableOfContentsValue {
  return {
    id: validTableOfContentsId(source.id) ?? 'document-table-of-contents',
    options: normalizeDocumentTableOfContentsOptions(source.options),
    entries: normalizeTableOfContentsEntries(source.entries),
    truncated: Boolean(source.truncated),
  };
}

function tableOfContentsEntryFromOutline(
  document: ProseMirrorNode,
  item: WorkDocumentOutlineItem,
  resolveContext: WorkDocumentFieldContextResolver | undefined,
): WorkDocumentTableOfContentsEntry {
  return {
    targetId: item.id,
    title: item.text.slice(0, MAX_TABLE_OF_CONTENTS_TITLE_LENGTH),
    level: item.level,
    pageNumber:
      positiveInteger(resolveContext?.(item.from)?.pageNumber) ??
      fallbackDocumentPageNumber(document, item.from),
  };
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

function tableOfContentsEntryHtml(
  entry: WorkDocumentTableOfContentsEntry,
  options: WorkDocumentTableOfContentsOptions,
): string {
  const title = escapeHtml(entry.title);
  const targetId = escapeHtmlAttribute(entry.targetId);
  const label = options.hyperlinks
    ? `<a href="#${targetId}" data-toc-target="${targetId}" tabindex="-1">${title}</a>`
    : `<span>${title}</span>`;
  const page = options.showPageNumbers
    ? `<span class="work-document-table-of-contents-page">${entry.pageNumber}</span>`
    : '';
  return `<li data-toc-target="${targetId}" data-toc-level="${entry.level}" style="--work-toc-level:${entry.level}">${label}<i aria-hidden="true"></i>${page}</li>`;
}

function parseTableOfContentsEntries(
  source: string | undefined,
): WorkDocumentTableOfContentsEntry[] {
  if (!source) return [];
  try {
    const parsed: unknown = JSON.parse(source);
    return Array.isArray(parsed) ? normalizeTableOfContentsEntries(parsed) : [];
  } catch {
    return [];
  }
}

function normalizeTableOfContentsEntries(
  source: readonly unknown[],
): WorkDocumentTableOfContentsEntry[] {
  const entries: WorkDocumentTableOfContentsEntry[] = [];
  for (const value of source.slice(0, MAX_DOCUMENT_TABLE_OF_CONTENTS_ENTRIES)) {
    if (!value || typeof value !== 'object') continue;
    const candidate = value as Record<string, unknown>;
    const targetId =
      typeof candidate.targetId === 'string' ? candidate.targetId.trim() : '';
    const title =
      typeof candidate.title === 'string'
        ? candidate.title.replace(/\s+/g, ' ').trim()
        : '';
    const level = boundedLevel(candidate.level, 0);
    const pageNumber = positiveInteger(candidate.pageNumber);
    if (
      !TABLE_OF_CONTENTS_TARGET_PATTERN.test(targetId) ||
      !title ||
      !level ||
      !pageNumber
    ) {
      continue;
    }
    entries.push({
      targetId,
      title: title.slice(0, MAX_TABLE_OF_CONTENTS_TITLE_LENGTH),
      level,
      pageNumber,
    });
  }
  return entries;
}

function parseTableOfContentsRange(
  value: string | null | undefined,
): { minLevel: number; maxLevel: number } | null {
  if (value === null || value === undefined) return null;
  const match = /^([1-9])-([1-9])$/.exec(value.trim());
  if (!match) return null;
  const minLevel = Number(match[1]);
  const maxLevel = Number(match[2]);
  return minLevel <= maxLevel ? { minLevel, maxLevel } : null;
}

function parseOptionalTableOfContentsRange(
  value: string | null | undefined,
): { minLevel: number; maxLevel: number } | null | undefined {
  if (value === null) return null;
  return parseTableOfContentsRange(value) ?? undefined;
}

function boundedLevel(value: unknown, fallback: number): number {
  const number = Number(value);
  return Number.isInteger(number) && number >= 1 && number <= 9
    ? number
    : fallback;
}

function positiveInteger(value: unknown): number | null {
  const number = Number(value);
  return Number.isSafeInteger(number) && number > 0
    ? Math.min(999_999, number)
    : null;
}

function validTableOfContentsId(value: unknown): string | null {
  return typeof value === 'string' && TABLE_OF_CONTENTS_ID_PATTERN.test(value)
    ? value
    : null;
}

function uniqueTableOfContentsId(
  source: string,
  index: number,
  usedIds: Set<string>,
): string {
  if (!usedIds.has(source)) {
    usedIds.add(source);
    return source;
  }
  let suffix = index;
  while (usedIds.has(`document-table-of-contents-${suffix}`)) suffix += 1;
  const id = `document-table-of-contents-${suffix}`;
  usedIds.add(id);
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
