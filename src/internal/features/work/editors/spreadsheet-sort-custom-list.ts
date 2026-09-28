import {
  officeMessage,
  resolveOfficeMessages,
  type OfficeMessageKey,
} from '../../../i18n/office-locale';
import {
  SPREADSHEET_SORT_BUILTIN_EN_MONTHS,
  SPREADSHEET_SORT_BUILTIN_EN_MONTHS_ABBR,
  SPREADSHEET_SORT_BUILTIN_EN_WEEKDAYS,
  SPREADSHEET_SORT_BUILTIN_EN_WEEKDAYS_ABBR,
  SPREADSHEET_SORT_BUILTIN_ZH_MONTHS,
  SPREADSHEET_SORT_BUILTIN_ZH_WEEKDAYS,
  SPREADSHEET_SORT_BUILTIN_ZH_WEEKDAYS_SHORT,
} from '../../../i18n/spreadsheet-sort-builtin-list-entries';

export const MAX_SPREADSHEET_SORT_CUSTOM_LIST_ENTRIES = 256;
export const MAX_SPREADSHEET_SORT_CUSTOM_LIST_ENTRY_CODE_POINTS = 128;
export const MAX_SPREADSHEET_SORT_CUSTOM_LIST_CODE_POINTS = 4_096;
export const MAX_SPREADSHEET_SORT_USER_CUSTOM_LISTS = 32;

export type SpreadsheetSortCustomListSource = 'built-in' | 'stored' | 'session';

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
    ...spreadsheetSortBuiltInCustomLists(),
  ];
  let userListCount = 0;
  for (const candidate of lists) {
    if (candidate.source !== 'stored' && candidate.source !== 'session') {
      continue;
    }
    const list = createSpreadsheetSortCustomList(
      candidate.entries,
      candidate.source,
    );
    if (
      !list ||
      merged.some((item) =>
        spreadsheetSortCustomListsEqual(item.entries, list.entries),
      )
    ) {
      continue;
    }
    if (userListCount >= MAX_SPREADSHEET_SORT_USER_CUSTOM_LISTS) break;
    merged.push(list);
    userListCount += 1;
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
  const catalog = resolveOfficeMessages();
  const keys: Record<SpreadsheetSortCustomListErrorCode, OfficeMessageKey> = {
    'duplicate-entry': 'spreadsheet.sort.customList.error.duplicateEntry',
    'entry-too-long': 'spreadsheet.sort.customList.error.entryTooLong',
    'invalid-entry': 'spreadsheet.sort.customList.error.invalidEntry',
    'list-too-long': 'spreadsheet.sort.customList.error.listTooLong',
    'not-enough-entries': 'spreadsheet.sort.customList.error.notEnoughEntries',
    'too-many-entries': 'spreadsheet.sort.customList.error.tooManyEntries',
  };
  const params =
    code === 'entry-too-long'
      ? { n: String(MAX_SPREADSHEET_SORT_CUSTOM_LIST_ENTRY_CODE_POINTS) }
      : code === 'list-too-long'
        ? {
            n: MAX_SPREADSHEET_SORT_CUSTOM_LIST_CODE_POINTS.toLocaleString(
              'en-US',
            ),
          }
        : code === 'too-many-entries'
          ? { n: String(MAX_SPREADSHEET_SORT_CUSTOM_LIST_ENTRIES) }
          : undefined;
  return {
    ok: false,
    code,
    message: officeMessage(catalog, keys[code], params),
  };
}

function builtInCustomList(
  labelKey: OfficeMessageKey,
  entries: readonly string[],
  catalog = resolveOfficeMessages(),
): SpreadsheetSortCustomList {
  return Object.freeze({
    source: 'built-in',
    entries: Object.freeze([...entries]),
    label: officeMessage(catalog, labelKey),
  });
}

export function spreadsheetSortBuiltInCustomLists(
  catalog = resolveOfficeMessages(),
): readonly SpreadsheetSortCustomList[] {
  return Object.freeze([
    builtInCustomList(
      'spreadsheet.sort.builtin.zhMonths',
      SPREADSHEET_SORT_BUILTIN_ZH_MONTHS,
      catalog,
    ),
    builtInCustomList(
      'spreadsheet.sort.builtin.zhWeekdays',
      SPREADSHEET_SORT_BUILTIN_ZH_WEEKDAYS,
      catalog,
    ),
    builtInCustomList(
      'spreadsheet.sort.builtin.zhWeekdaysShort',
      SPREADSHEET_SORT_BUILTIN_ZH_WEEKDAYS_SHORT,
      catalog,
    ),
    builtInCustomList(
      'spreadsheet.sort.builtin.enMonths',
      SPREADSHEET_SORT_BUILTIN_EN_MONTHS,
      catalog,
    ),
    builtInCustomList(
      'spreadsheet.sort.builtin.enMonthsAbbr',
      SPREADSHEET_SORT_BUILTIN_EN_MONTHS_ABBR,
      catalog,
    ),
    builtInCustomList(
      'spreadsheet.sort.builtin.enWeekdays',
      SPREADSHEET_SORT_BUILTIN_EN_WEEKDAYS,
      catalog,
    ),
    builtInCustomList(
      'spreadsheet.sort.builtin.enWeekdaysAbbr',
      SPREADSHEET_SORT_BUILTIN_EN_WEEKDAYS_ABBR,
      catalog,
    ),
  ]);
}

/** @deprecated Prefer spreadsheetSortBuiltInCustomLists() for locale-aware labels. */
export const SPREADSHEET_SORT_BUILT_IN_CUSTOM_LISTS =
  spreadsheetSortBuiltInCustomLists();
