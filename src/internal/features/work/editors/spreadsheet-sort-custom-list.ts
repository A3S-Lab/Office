export const MAX_SPREADSHEET_SORT_CUSTOM_LIST_ENTRIES = 256;
export const MAX_SPREADSHEET_SORT_CUSTOM_LIST_ENTRY_CODE_POINTS = 128;
export const MAX_SPREADSHEET_SORT_CUSTOM_LIST_CODE_POINTS = 4_096;
export const MAX_SPREADSHEET_SORT_SESSION_CUSTOM_LISTS = 32;

export type SpreadsheetSortCustomListSource = 'built-in' | 'session';

export interface SpreadsheetSortCustomList {
  entries: readonly string[];
  label: string;
  source: SpreadsheetSortCustomListSource;
}

export type SpreadsheetSortCustomListErrorCode =
  | 'duplicate-entry'
  | 'entry-too-long'
  | 'invalid-entry'
  | 'list-too-long'
  | 'not-enough-entries'
  | 'too-many-entries';

export type SpreadsheetSortCustomListValidationResult =
  | { entries: string[]; ok: true }
  | {
      code: SpreadsheetSortCustomListErrorCode;
      message: string;
      ok: false;
    };

export function parseSpreadsheetSortCustomList(
  text: string,
): SpreadsheetSortCustomListValidationResult {
  return validateSpreadsheetSortCustomList(
    text
      .split(/\r?\n|,|，/u)
      .map((entry) => entry.trim())
      .filter(Boolean),
  );
}

export function validateSpreadsheetSortCustomList(
  input: unknown,
): SpreadsheetSortCustomListValidationResult {
  if (!Array.isArray(input)) return customListError('invalid-entry');
  if (input.length < 2) return customListError('not-enough-entries');
  if (input.length > MAX_SPREADSHEET_SORT_CUSTOM_LIST_ENTRIES) {
    return customListError('too-many-entries');
  }

  const entries: string[] = [];
  const seen = new Set<string>();
  let totalCodePoints = 0;
  for (const candidate of input) {
    if (typeof candidate !== 'string') return customListError('invalid-entry');
    const entry = candidate.trim();
    if (!entry) return customListError('invalid-entry');
    const codePoints = Array.from(entry).length;
    if (codePoints > MAX_SPREADSHEET_SORT_CUSTOM_LIST_ENTRY_CODE_POINTS) {
      return customListError('entry-too-long');
    }
    totalCodePoints += codePoints;
    if (totalCodePoints > MAX_SPREADSHEET_SORT_CUSTOM_LIST_CODE_POINTS) {
      return customListError('list-too-long');
    }
    const matchKey = spreadsheetSortCustomListMatchKey(entry);
    if (seen.has(matchKey)) return customListError('duplicate-entry');
    seen.add(matchKey);
    entries.push(entry);
  }
  return { ok: true, entries };
}

export function createSpreadsheetSortCustomList(
  entries: readonly string[],
  source: SpreadsheetSortCustomListSource = 'session',
): SpreadsheetSortCustomList | null {
  const validation = validateSpreadsheetSortCustomList(entries);
  if (!validation.ok) return null;
  const normalized = Object.freeze([...validation.entries]);
  return Object.freeze({
    source,
    entries: normalized,
    label: spreadsheetSortCustomListLabel(normalized),
  });
}

export function mergeSpreadsheetSortCustomLists(
  lists: readonly SpreadsheetSortCustomList[],
): readonly SpreadsheetSortCustomList[] {
  const merged: SpreadsheetSortCustomList[] = [
    ...SPREADSHEET_SORT_BUILT_IN_CUSTOM_LISTS,
  ];
  let sessionCount = 0;
  for (const candidate of lists) {
    if (candidate.source !== 'session') continue;
    const list = createSpreadsheetSortCustomList(candidate.entries, 'session');
    if (
      !list ||
      merged.some((item) =>
        spreadsheetSortCustomListsEqual(item.entries, list.entries),
      )
    ) {
      continue;
    }
    if (sessionCount >= MAX_SPREADSHEET_SORT_SESSION_CUSTOM_LISTS) break;
    merged.push(list);
    sessionCount += 1;
  }
  return Object.freeze(merged);
}

export function spreadsheetSortCustomListsEqual(
  left: readonly string[],
  right: readonly string[],
): boolean {
  return (
    left.length === right.length &&
    left.every(
      (entry, index) =>
        spreadsheetSortCustomListMatchKey(entry) ===
        spreadsheetSortCustomListMatchKey(right[index] ?? ''),
    )
  );
}

export function spreadsheetSortCustomListMatchKey(
  value: number | string,
): string {
  return String(value).normalize('NFKC').trim().toLocaleLowerCase('zh-CN');
}

function spreadsheetSortCustomListLabel(entries: readonly string[]): string {
  const visible = entries
    .slice(0, 3)
    .map((entry) => truncateSpreadsheetSortCustomListLabelEntry(entry));
  return `${visible.join(' → ')}${entries.length > 3 ? ' → …' : ''}`;
}

function truncateSpreadsheetSortCustomListLabelEntry(entry: string): string {
  const characters = Array.from(entry);
  return characters.length > 18
    ? `${characters.slice(0, 17).join('')}…`
    : entry;
}

function customListError(
  code: SpreadsheetSortCustomListErrorCode,
): Extract<SpreadsheetSortCustomListValidationResult, { ok: false }> {
  const messages: Record<SpreadsheetSortCustomListErrorCode, string> = {
    'duplicate-entry': '自定义序列不能包含重复项（忽略大小写和全半角差异）。',
    'entry-too-long': `每个序列项最多包含 ${MAX_SPREADSHEET_SORT_CUSTOM_LIST_ENTRY_CODE_POINTS} 个字符。`,
    'invalid-entry': '自定义序列只能包含非空文本项。',
    'list-too-long': `一个自定义序列最多包含 ${MAX_SPREADSHEET_SORT_CUSTOM_LIST_CODE_POINTS.toLocaleString('en-US')} 个字符。`,
    'not-enough-entries': '自定义序列至少需要两个项目。',
    'too-many-entries': `一个自定义序列最多包含 ${MAX_SPREADSHEET_SORT_CUSTOM_LIST_ENTRIES} 个项目。`,
  };
  return { ok: false, code, message: messages[code] };
}

function builtInCustomList(
  label: string,
  entries: readonly string[],
): SpreadsheetSortCustomList {
  return Object.freeze({
    source: 'built-in',
    entries: Object.freeze([...entries]),
    label,
  });
}

export const SPREADSHEET_SORT_BUILT_IN_CUSTOM_LISTS = Object.freeze([
  builtInCustomList('月份（一月 → 十二月）', [
    '一月',
    '二月',
    '三月',
    '四月',
    '五月',
    '六月',
    '七月',
    '八月',
    '九月',
    '十月',
    '十一月',
    '十二月',
  ]),
  builtInCustomList('星期（星期日 → 星期六）', [
    '星期日',
    '星期一',
    '星期二',
    '星期三',
    '星期四',
    '星期五',
    '星期六',
  ]),
  builtInCustomList('周（周日 → 周六）', [
    '周日',
    '周一',
    '周二',
    '周三',
    '周四',
    '周五',
    '周六',
  ]),
  builtInCustomList('月份（January → December）', [
    'January',
    'February',
    'March',
    'April',
    'May',
    'June',
    'July',
    'August',
    'September',
    'October',
    'November',
    'December',
  ]),
  builtInCustomList('月份（Jan → Dec）', [
    'Jan',
    'Feb',
    'Mar',
    'Apr',
    'May',
    'Jun',
    'Jul',
    'Aug',
    'Sep',
    'Oct',
    'Nov',
    'Dec',
  ]),
  builtInCustomList('星期（Sunday → Saturday）', [
    'Sunday',
    'Monday',
    'Tuesday',
    'Wednesday',
    'Thursday',
    'Friday',
    'Saturday',
  ]),
  builtInCustomList('星期（Sun → Sat）', [
    'Sun',
    'Mon',
    'Tue',
    'Wed',
    'Thu',
    'Fri',
    'Sat',
  ]),
] as const);
