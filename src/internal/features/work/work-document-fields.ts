import { documentWordCount } from './work-document-text-statistics';

export type WorkDocumentFieldKind =
  | 'page'
  | 'numPages'
  | 'section'
  | 'sectionPages'
  | 'date'
  | 'time'
  | 'wordCount'
  | 'characterCount'
  | 'pageReference';

export interface WorkDocumentFieldContext {
  pageNumber: number;
  totalPages: number;
  sectionNumber: number;
  sectionPages: number;
  /** Number of visible body words, excluding generated field results. */
  wordCount?: number;
  /** Number of visible body characters, including spaces. */
  characterCount?: number;
  /** Page number resolved for a PAGEREF target, when the target is present. */
  referencePageNumber?: number | null;
  /** Page numbers keyed by `id:<bookmark-id>` or `name:<bookmark-name>`. */
  bookmarkPageNumbers?: ReadonlyMap<string, number>;
  now?: Date;
}

export type WorkDocumentFieldContextResolver = (
  position: number,
) => WorkDocumentFieldContext | null;

export interface WorkDocumentFieldRefreshOptions {
  resolveContext?: WorkDocumentFieldContextResolver;
  now?: Date;
  addToHistory?: boolean;
  updateClock?: boolean;
}

export interface WorkDocumentFieldInsertOptions {
  /** Bookmark identity used to keep a page reference stable across renames. */
  targetId?: string;
  /** Bookmark name emitted in the native PAGEREF instruction. */
  targetName?: string;
}

type WorkDocumentNumericFieldFormat =
  | 'arabic'
  | 'roman'
  | 'romanLower'
  | 'alphabetic'
  | 'alphabeticLower'
  | 'ordinal';

const FIELD_SELECTOR = 'span[data-document-field]';
const FIELD_TEXT_BOUNDARY = '\uFFFC';

const FIELD_COMMANDS: Record<WorkDocumentFieldKind, string> = {
  page: 'PAGE',
  numPages: 'NUMPAGES',
  section: 'SECTION',
  sectionPages: 'SECTIONPAGES',
  date: 'DATE \\@ "yyyy年M月d日"',
  time: 'TIME \\@ "HH:mm"',
  wordCount: 'NUMWORDS',
  characterCount: 'NUMCHARS',
  pageReference: 'PAGEREF',
};

const FIELD_LABELS: Record<WorkDocumentFieldKind, string> = {
  page: '当前页码',
  numPages: '总页数',
  section: '当前节号',
  sectionPages: '本节页数',
  date: '当前日期',
  time: '当前时间',
  wordCount: '字数',
  characterCount: '字符数',
  pageReference: '目标页码',
};

export function documentFieldKind(
  value: string | undefined,
): WorkDocumentFieldKind | null {
  if (
    value === 'page' ||
    value === 'numPages' ||
    value === 'section' ||
    value === 'sectionPages' ||
    value === 'date' ||
    value === 'time' ||
    value === 'wordCount' ||
    value === 'characterCount' ||
    value === 'pageReference'
  ) {
    return value;
  }
  return null;
}

export function docxDocumentFieldKind(
  instruction: string,
): WorkDocumentFieldKind | null {
  const command = /^\s*([a-z][a-z0-9]*)\b/i
    .exec(instruction)?.[1]
    ?.toUpperCase();
  if (command === 'PAGE') return 'page';
  if (command === 'NUMPAGES') return 'numPages';
  if (command === 'SECTION') return 'section';
  if (command === 'SECTIONPAGES') return 'sectionPages';
  if (command === 'DATE') return 'date';
  if (command === 'TIME') return 'time';
  if (command === 'NUMWORDS') return 'wordCount';
  if (command === 'NUMCHARS') return 'characterCount';
  if (command === 'PAGEREF') return 'pageReference';
  return null;
}

export function documentFieldInstruction(
  kind: WorkDocumentFieldKind,
  options: WorkDocumentFieldInsertOptions = {},
): string {
  if (kind === 'pageReference') {
    return documentPageReferenceInstruction(options.targetName, '', true);
  }
  return FIELD_COMMANDS[kind];
}

/** Builds a stable PAGEREF instruction while retaining the supported switches. */
export function documentPageReferenceInstruction(
  targetName: unknown,
  source = '',
  defaultHyperlink = false,
): string {
  const target = normalizeFieldTarget(targetName);
  if (!target) return FIELD_COMMANDS.pageReference;
  const switches: string[] = [];
  if (/(?:^|\s)\\h(?:\s|$)/i.test(source)) switches.push('\\h');
  const numericFormat = numericFieldFormatSwitch(source);
  if (numericFormat) switches.push(numericFormat);
  if (/(?:^|\s)\\\*\s+MERGEFORMAT(?:\s|$)/i.test(source)) {
    switches.push('\\* MERGEFORMAT');
  }
  if (!switches.length && defaultHyperlink) switches.push('\\h');
  return `PAGEREF ${target}${switches.length ? ` ${switches.join(' ')}` : ''}`;
}

export function documentFieldLabel(kind: WorkDocumentFieldKind): string {
  return FIELD_LABELS[kind];
}

export function documentFieldDisplay(
  kind: WorkDocumentFieldKind,
  context: WorkDocumentFieldContext,
  instruction = documentFieldInstruction(kind),
  cachedValue = '',
): string {
  if (kind === 'page')
    return formatNumericFieldValue(
      positiveInteger(context.pageNumber),
      instruction,
    );
  if (kind === 'numPages')
    return formatNumericFieldValue(
      positiveInteger(context.totalPages),
      instruction,
    );
  if (kind === 'section')
    return formatNumericFieldValue(
      positiveInteger(context.sectionNumber),
      instruction,
    );
  if (kind === 'sectionPages')
    return formatNumericFieldValue(
      positiveInteger(context.sectionPages),
      instruction,
    );
  if (kind === 'wordCount')
    return String(nonNegativeInteger(context.wordCount, cachedValue));
  if (kind === 'characterCount')
    return String(nonNegativeInteger(context.characterCount, cachedValue));
  if (kind === 'pageReference') {
    const target = docxDocumentFieldTarget(instruction);
    const hasResolutionContext =
      context.referencePageNumber !== undefined ||
      context.bookmarkPageNumbers !== undefined;
    const page =
      context.referencePageNumber ??
      (target
        ? context.bookmarkPageNumbers?.get(`name:${target.toLowerCase()}`)
        : undefined);
    return page && Number.isSafeInteger(page) && page > 0
      ? formatNumericFieldValue(page, instruction)
      : hasResolutionContext
        ? '引用缺失'
        : cachedValue.trim() || '引用缺失';
  }
  const now = validDate(context.now) ?? new Date();
  const format =
    dateFormatSwitch(instruction) ??
    (kind === 'date' ? 'yyyy年M月d日' : 'HH:mm');
  const display = formatWordDate(now, format);
  return display || cachedValue || documentFieldLabel(kind);
}

export function normalizeDocumentFieldsHtml(source: string): string {
  const document = new DOMParser().parseFromString(source, 'text/html');
  const usedIds = new Set<string>();
  for (const [index, element] of Array.from(
    document.body.querySelectorAll<HTMLElement>(FIELD_SELECTOR),
  ).entries()) {
    const instruction = element.dataset.fieldInstruction?.trim() ?? '';
    const kind =
      documentFieldKind(element.dataset.fieldKind) ??
      docxDocumentFieldKind(instruction);
    if (!kind) {
      element.replaceWith(document.createTextNode(element.textContent ?? ''));
      continue;
    }
    const display =
      element.dataset.fieldDisplay?.trim() ||
      element.textContent?.trim() ||
      documentFieldLabel(kind);
    element.dataset.documentField = 'true';
    element.dataset.fieldId = uniqueFieldId(
      element.dataset.fieldId,
      index + 1,
      usedIds,
    );
    element.dataset.fieldKind = kind;
    element.dataset.fieldInstruction =
      instruction || documentFieldInstruction(kind);
    if (kind === 'pageReference') {
      const targetName =
        normalizeFieldTarget(element.dataset.fieldTargetName) ??
        docxDocumentFieldTarget(instruction);
      if (targetName) {
        element.dataset.fieldTargetName = targetName;
        element.dataset.fieldInstruction = documentPageReferenceInstruction(
          targetName,
          instruction,
        );
        delete element.dataset.fieldOrphaned;
      } else {
        delete element.dataset.fieldTargetName;
        element.dataset.fieldOrphaned = 'true';
      }
    }
    element.dataset.fieldDisplay = display;
    element.classList.add('work-document-field');
    element.textContent = display;
  }
  return document.body.innerHTML;
}

export function resolveDocumentFieldsHtml(
  source: string,
  context: WorkDocumentFieldContext,
): string {
  const document = new DOMParser().parseFromString(
    normalizeDocumentFieldsHtml(source),
    'text/html',
  );
  for (const element of Array.from(
    document.body.querySelectorAll<HTMLElement>(FIELD_SELECTOR),
  )) {
    const kind = documentFieldKind(element.dataset.fieldKind);
    if (!kind) continue;
    const referencePageNumber =
      kind === 'pageReference'
        ? (context.referencePageNumber ??
          (element.dataset.fieldTargetId
            ? context.bookmarkPageNumbers?.get(
                `id:${element.dataset.fieldTargetId.trim()}`,
              )
            : undefined))
        : context.referencePageNumber;
    const display = documentFieldDisplay(
      kind,
      { ...context, referencePageNumber },
      element.dataset.fieldInstruction,
      element.dataset.fieldDisplay,
    );
    element.dataset.fieldDisplay = display;
    element.textContent = display;
  }
  return document.body.innerHTML;
}

function uniqueFieldId(
  source: string | undefined,
  index: number,
  usedIds: Set<string>,
): string {
  const candidate = source?.trim();
  if (candidate && !usedIds.has(candidate)) {
    usedIds.add(candidate);
    return candidate;
  }
  let suffix = index;
  while (usedIds.has(`document-field-${suffix}`)) suffix += 1;
  const id = `document-field-${suffix}`;
  usedIds.add(id);
  return id;
}

function dateFormatSwitch(instruction: string): string | null {
  return /\\@\s+"([^"]+)"/i.exec(instruction)?.[1] ?? null;
}

/** Returns the bookmark name used by a bounded PAGEREF instruction. */
export function docxDocumentFieldTarget(instruction: string): string | null {
  const kind = docxDocumentFieldKind(instruction);
  if (kind !== 'pageReference') return null;
  return normalizeFieldTarget(
    /^\s*PAGEREF\s+([^\s\\]+)/i.exec(instruction)?.[1],
  );
}

/**
 * The editor intentionally accepts only switches whose result can be
 * represented by one deterministic inline atom.  Unknown switches remain
 * cached text instead of being presented as an editable field with altered
 * semantics.
 */
export function supportedDocxDocumentFieldInstruction(
  instruction: string,
): boolean {
  const kind = docxDocumentFieldKind(instruction);
  if (!kind) return false;
  const source = instruction.trim();
  if (kind === 'pageReference') {
    const match = /^PAGEREF\s+([^\s\\]+)([\s\S]*)$/i.exec(source);
    const target = normalizeFieldTarget(match?.[1]);
    if (!target || !match) return false;
    let rest = match[2] ?? '';
    const hyperlink = /^\s+\\h\b/i.exec(rest);
    if (hyperlink) rest = rest.slice(hyperlink[0].length);
    return onlyNumericFieldSwitches(rest);
  }
  if (kind === 'date' || kind === 'time') {
    const command = kind === 'date' ? 'DATE' : 'TIME';
    const match = new RegExp(`^${command}\\b([\\s\\S]*)$`, 'i').exec(source);
    if (!match) return false;
    let rest = match[1] ?? '';
    const format = /^\s+\\@\s+"[^"\r\n]{1,128}"/i.exec(rest);
    if (format) rest = rest.slice(format[0].length);
    return onlyMergeFormatSwitch(rest);
  }
  const command = FIELD_COMMANDS[kind];
  const match = new RegExp(`^${command}\\b([\\s\\S]*)$`, 'i').exec(source);
  return Boolean(
    match &&
      (isNumericFieldKind(kind)
        ? onlyNumericFieldSwitches(match[1] ?? '')
        : onlyMergeFormatSwitch(match[1] ?? '')),
  );
}

/** Returns the supported numeric display switch, preserving its source case. */
export function numericFieldFormatSwitch(source: string): string | null {
  return /(?:^|\s)(\\\*\s+(?:Arabic|ROMAN|roman|ALPHABETIC|alphabetic|Ordinal)\b)/i.exec(
    source,
  )?.[1] ?? null;
}

function isNumericFieldKind(kind: WorkDocumentFieldKind): boolean {
  return (
    kind === 'page' ||
    kind === 'numPages' ||
    kind === 'section' ||
    kind === 'sectionPages' ||
    kind === 'pageReference'
  );
}

function onlyNumericFieldSwitches(source: string): boolean {
  let rest = source.trim();
  let formatSeen = false;
  let mergeFormatSeen = false;
  while (rest) {
    const format = /^\\\*\s+(Arabic|ROMAN|roman|ALPHABETIC|alphabetic|Ordinal)\b/i.exec(
      rest,
    );
    if (format && !formatSeen) {
      formatSeen = true;
      rest = rest.slice(format[0].length).trim();
      continue;
    }
    const mergeFormat = /^\\\*\s+MERGEFORMAT\b/i.exec(rest);
    if (mergeFormat && !mergeFormatSeen) {
      mergeFormatSeen = true;
      rest = rest.slice(mergeFormat[0].length).trim();
      continue;
    }
    return false;
  }
  return true;
}

function formatNumericFieldValue(value: number, instruction: string): string {
  const format = numericFieldFormatSwitchValue(instruction);
  if (format === 'roman' || format === 'romanLower') {
    const roman = toRoman(value);
    return format === 'romanLower' ? roman.toLowerCase() : roman;
  }
  if (format === 'alphabetic' || format === 'alphabeticLower') {
    const alphabetic = toAlphabetic(value);
    return format === 'alphabeticLower'
      ? alphabetic.toLowerCase()
      : alphabetic;
  }
  if (format === 'ordinal') return ordinal(value);
  return String(value);
}

function numericFieldFormatSwitchValue(
  instruction: string,
): WorkDocumentNumericFieldFormat {
  const token = /\\\*\s+(Arabic|ROMAN|roman|ALPHABETIC|alphabetic|Ordinal)\b/i.exec(
    instruction,
  )?.[1];
  if (token === 'ROMAN') return 'roman';
  if (token && token.toLowerCase() === 'roman') return 'romanLower';
  if (token === 'ALPHABETIC') return 'alphabetic';
  if (token && token.toLowerCase() === 'alphabetic') {
    return 'alphabeticLower';
  }
  if (token && token.toLowerCase() === 'ordinal') {
    return 'ordinal';
  }
  return 'arabic';
}

function toRoman(value: number): string {
  const digits: Array<[number, string]> = [
    [1000, 'M'],
    [900, 'CM'],
    [500, 'D'],
    [400, 'CD'],
    [100, 'C'],
    [90, 'XC'],
    [50, 'L'],
    [40, 'XL'],
    [10, 'X'],
    [9, 'IX'],
    [5, 'V'],
    [4, 'IV'],
    [1, 'I'],
  ];
  let remaining = Math.max(1, Math.min(3999, Math.trunc(value)));
  let result = '';
  for (const [unit, symbol] of digits) {
    while (remaining >= unit) {
      result += symbol;
      remaining -= unit;
    }
  }
  return result;
}

function toAlphabetic(value: number): string {
  let remaining = Math.max(1, Math.trunc(value));
  let result = '';
  while (remaining > 0) {
    remaining -= 1;
    result = String.fromCharCode(65 + (remaining % 26)) + result;
    remaining = Math.floor(remaining / 26);
  }
  return result;
}

function ordinal(value: number): string {
  const mod100 = value % 100;
  const suffix =
    mod100 >= 11 && mod100 <= 13
      ? 'th'
      : ({ 1: 'st', 2: 'nd', 3: 'rd' } as Record<number, string>)[
            value % 10
          ] ?? 'th';
  return `${value}${suffix}`;
}

function onlyMergeFormatSwitch(source: string): boolean {
  return /^(?:\s+\\\*\s+MERGEFORMAT)?\s*$/i.test(source);
}

export function documentFieldStatisticsFromText(source: string): {
  wordCount: number;
  characterCount: number;
} {
  const normalized = source.replace(/\r\n?/g, '\n');
  let characterCount = 0;
  for (let index = 0; index < normalized.length; index += 1) {
    const codePoint = normalized.codePointAt(index);
    if (codePoint === undefined) continue;
    if (codePoint > 0xffff) index += 1;
    if (codePoint === 0x0a || codePoint === 0xfffc) continue;
    characterCount += 1;
  }
  return {
    wordCount: documentWordCount(normalized),
    characterCount,
  };
}

/** Computes bounded body statistics while excluding generated field results. */
export function documentFieldStatisticsFromHtml(source: string): {
  wordCount: number;
  characterCount: number;
} {
  const document = new DOMParser().parseFromString(source, 'text/html');
  for (const field of Array.from(
    document.body.querySelectorAll('[data-document-field]'),
  )) {
    field.replaceWith(document.createTextNode(FIELD_TEXT_BOUNDARY));
  }
  return documentFieldStatisticsFromText(document.body.textContent ?? '');
}

function formatWordDate(date: Date, format: string): string {
  const hour12 = date.getHours() % 12 || 12;
  const replacements: Record<string, string> = {
    'AM/PM': date.getHours() < 12 ? 'AM' : 'PM',
    'am/pm': date.getHours() < 12 ? 'am' : 'pm',
    yyyy: String(date.getFullYear()).padStart(4, '0'),
    yy: String(date.getFullYear() % 100).padStart(2, '0'),
    MMMM: new Intl.DateTimeFormat('zh-CN', { month: 'long' }).format(date),
    MMM: new Intl.DateTimeFormat('zh-CN', { month: 'short' }).format(date),
    MM: String(date.getMonth() + 1).padStart(2, '0'),
    M: String(date.getMonth() + 1),
    dddd: new Intl.DateTimeFormat('zh-CN', { weekday: 'long' }).format(date),
    ddd: new Intl.DateTimeFormat('zh-CN', { weekday: 'short' }).format(date),
    dd: String(date.getDate()).padStart(2, '0'),
    d: String(date.getDate()),
    HH: String(date.getHours()).padStart(2, '0'),
    H: String(date.getHours()),
    hh: String(hour12).padStart(2, '0'),
    h: String(hour12),
    mm: String(date.getMinutes()).padStart(2, '0'),
    m: String(date.getMinutes()),
    ss: String(date.getSeconds()).padStart(2, '0'),
    s: String(date.getSeconds()),
  };
  return format.replace(
    /AM\/PM|am\/pm|yyyy|MMMM|dddd|MMM|ddd|yy|MM|dd|HH|hh|mm|ss|M|d|H|h|m|s/g,
    (token) => replacements[token] ?? token,
  );
}

function positiveInteger(value: number): number {
  return Number.isSafeInteger(value) && value > 0 ? value : 1;
}

function nonNegativeInteger(
  value: number | undefined,
  cachedValue: string,
): number {
  if (value !== undefined && Number.isSafeInteger(value) && value >= 0) {
    return value;
  }
  const cached = Number(cachedValue.trim());
  return Number.isSafeInteger(cached) && cached >= 0 ? cached : 0;
}

function normalizeFieldTarget(value: unknown): string | null {
  if (typeof value !== 'string') return null;
  const target = value.trim();
  return /^[\p{L}_][\p{L}\p{N}_]*$/u.test(target) ? target : null;
}

function validDate(value: Date | undefined): Date | null {
  return value && Number.isFinite(value.getTime()) ? value : null;
}
