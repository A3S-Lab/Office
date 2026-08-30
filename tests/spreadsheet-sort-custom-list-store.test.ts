import { describe, expect, test } from '@rstest/core';
import {
  DEFAULT_SPREADSHEET_SORT_CUSTOM_LIST_STORAGE_KEY,
  LocalStorageSpreadsheetSortCustomListStore,
} from '../src/internal/features/work/editors/spreadsheet-sort-custom-list-store';

describe('spreadsheet sort custom-list store', () => {
  test('round-trips canonical local lists across store instances', () => {
    const storage = new MemoryStorage();
    const first = new LocalStorageSpreadsheetSortCustomListStore(storage);

    first.save([
      ['有风险', '进行中', '正常', '已完成'],
      [' High ', 'Medium', 'Low'],
      ['有风险', '进行中', '正常', '已完成'],
    ]);

    expect(
      new LocalStorageSpreadsheetSortCustomListStore(storage).load(),
    ).toEqual([
      ['有风险', '进行中', '正常', '已完成'],
      ['High', 'Medium', 'Low'],
    ]);
    expect(
      JSON.parse(
        storage.getItem(DEFAULT_SPREADSHEET_SORT_CUSTOM_LIST_STORAGE_KEY) ?? '',
      ),
    ).toEqual({
      version: 1,
      lists: [
        ['有风险', '进行中', '正常', '已完成'],
        ['High', 'Medium', 'Low'],
      ],
    });
  });

  test('fails closed for malformed or unsupported persisted payloads', () => {
    const storage = new MemoryStorage();
    const store = new LocalStorageSpreadsheetSortCustomListStore(storage);
    storage.setItem(DEFAULT_SPREADSHEET_SORT_CUSTOM_LIST_STORAGE_KEY, '{');
    expect(store.load()).toEqual([]);

    storage.setItem(
      DEFAULT_SPREADSHEET_SORT_CUSTOM_LIST_STORAGE_KEY,
      JSON.stringify({ version: 2, lists: [['High', 'Low']] }),
    );
    expect(store.load()).toEqual([]);

    storage.setItem(
      DEFAULT_SPREADSHEET_SORT_CUSTOM_LIST_STORAGE_KEY,
      JSON.stringify({
        version: 1,
        lists: [['Only one'], ['North', 'South'], ['Ａ', 'A']],
      }),
    );
    expect(store.load()).toEqual([['North', 'South']]);
  });
});

class MemoryStorage implements Pick<Storage, 'getItem' | 'setItem'> {
  readonly values = new Map<string, string>();

  getItem(key: string): string | null {
    return this.values.get(key) ?? null;
  }

  setItem(key: string, value: string): void {
    this.values.set(key, value);
  }
}
