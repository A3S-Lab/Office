import {
  MAX_SPREADSHEET_SORT_USER_CUSTOM_LISTS,
  spreadsheetSortCustomListsEqual,
  validateSpreadsheetSortCustomList,
} from './spreadsheet-sort-custom-list';

const SPREADSHEET_SORT_CUSTOM_LIST_STORAGE_VERSION = 1;

export const DEFAULT_SPREADSHEET_SORT_CUSTOM_LIST_STORAGE_KEY =
  'a3s-office.spreadsheet-sort-custom-lists.v1';

export interface SpreadsheetSortCustomListStore {
  load(): readonly (readonly string[])[];
  save(lists: readonly (readonly string[])[]): void;
}

export class LocalStorageSpreadsheetSortCustomListStore
  implements SpreadsheetSortCustomListStore
{
  constructor(
    private readonly storage: Pick<Storage, 'getItem' | 'setItem'>,
    private readonly key = DEFAULT_SPREADSHEET_SORT_CUSTOM_LIST_STORAGE_KEY,
  ) {}

  load(): readonly (readonly string[])[] {
    try {
      const source = this.storage.getItem(this.key);
      if (!source) return Object.freeze([]);
      const payload: unknown = JSON.parse(source);
      if (
        !isRecord(payload) ||
        payload.version !== SPREADSHEET_SORT_CUSTOM_LIST_STORAGE_VERSION
      ) {
        return Object.freeze([]);
      }
      return normalizeStoredSpreadsheetSortCustomLists(payload.lists);
    } catch {
      return Object.freeze([]);
    }
  }

  save(lists: readonly (readonly string[])[]): void {
    const normalized = normalizeStoredSpreadsheetSortCustomLists(lists);
    this.storage.setItem(
      this.key,
      JSON.stringify({
        version: SPREADSHEET_SORT_CUSTOM_LIST_STORAGE_VERSION,
        lists: normalized,
      }),
    );
  }
}

export function normalizeStoredSpreadsheetSortCustomLists(
  input: unknown,
): readonly (readonly string[])[] {
  if (!Array.isArray(input)) return Object.freeze([]);
  const lists: Array<readonly string[]> = [];
  for (const candidate of input) {
    const validation = validateSpreadsheetSortCustomList(candidate);
    if (
      !validation.ok ||
      lists.some((entries) =>
        spreadsheetSortCustomListsEqual(entries, validation.entries),
      )
    ) {
      continue;
    }
    lists.push(Object.freeze([...validation.entries]));
    if (lists.length >= MAX_SPREADSHEET_SORT_USER_CUSTOM_LISTS) break;
  }
  return Object.freeze(lists);
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value && typeof value === 'object' && !Array.isArray(value));
}
