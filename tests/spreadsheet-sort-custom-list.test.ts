import { describe, expect, test } from '@rstest/core';
import {
  createSpreadsheetSortCustomList,
  MAX_SPREADSHEET_SORT_CUSTOM_LIST_CODE_POINTS,
  MAX_SPREADSHEET_SORT_CUSTOM_LIST_ENTRIES,
  MAX_SPREADSHEET_SORT_CUSTOM_LIST_ENTRY_CODE_POINTS,
  MAX_SPREADSHEET_SORT_SESSION_CUSTOM_LISTS,
  mergeSpreadsheetSortCustomLists,
  parseSpreadsheetSortCustomList,
  SPREADSHEET_SORT_BUILT_IN_CUSTOM_LISTS,
  spreadsheetSortCustomListsEqual,
  validateSpreadsheetSortCustomList,
} from '../src/internal/features/work/editors/spreadsheet-sort-custom-list';

describe('spreadsheet sort custom lists', () => {
  test('normalizes newline and comma input into one reusable bounded list', () => {
    expect(
      parseSpreadsheetSortCustomList(' 有风险\n进行中，正常, 已完成\n'),
    ).toEqual({
      ok: true,
      entries: ['有风险', '进行中', '正常', '已完成'],
    });

    expect(
      createSpreadsheetSortCustomList(
        ['有风险', '进行中', '正常', '已完成'],
        'session',
      ),
    ).toEqual({
      source: 'session',
      label: '有风险 → 进行中 → 正常 → …',
      entries: ['有风险', '进行中', '正常', '已完成'],
    });
  });

  test('rejects ambiguous, duplicate, overlong, and over-count lists', () => {
    expect(validateSpreadsheetSortCustomList(['Only one'])).toMatchObject({
      ok: false,
      code: 'not-enough-entries',
    });
    expect(
      validateSpreadsheetSortCustomList(['High', ' high ', 'Low']),
    ).toMatchObject({ ok: false, code: 'duplicate-entry' });
    expect(validateSpreadsheetSortCustomList(['Ａ', 'A'])).toMatchObject({
      ok: false,
      code: 'duplicate-entry',
    });
    expect(
      validateSpreadsheetSortCustomList([
        'A'.repeat(MAX_SPREADSHEET_SORT_CUSTOM_LIST_ENTRY_CODE_POINTS + 1),
        'B',
      ]),
    ).toMatchObject({ ok: false, code: 'entry-too-long' });
    expect(
      validateSpreadsheetSortCustomList(
        Array.from(
          { length: MAX_SPREADSHEET_SORT_CUSTOM_LIST_ENTRIES + 1 },
          (_, index) => `Item ${index}`,
        ),
      ),
    ).toMatchObject({ ok: false, code: 'too-many-entries' });
    expect(
      validateSpreadsheetSortCustomList(
        Array.from(
          {
            length:
              MAX_SPREADSHEET_SORT_CUSTOM_LIST_CODE_POINTS /
                MAX_SPREADSHEET_SORT_CUSTOM_LIST_ENTRY_CODE_POINTS +
              1,
          },
          (_, index) =>
            `${String.fromCodePoint(0x4e00 + index)}${'A'.repeat(
              MAX_SPREADSHEET_SORT_CUSTOM_LIST_ENTRY_CODE_POINTS - 1,
            )}`,
        ),
      ),
    ).toMatchObject({ ok: false, code: 'list-too-long' });
  });

  test('ships distinct month and weekday orders without mutable aliases', () => {
    expect(SPREADSHEET_SORT_BUILT_IN_CUSTOM_LISTS).toHaveLength(7);
    for (const list of SPREADSHEET_SORT_BUILT_IN_CUSTOM_LISTS) {
      expect(validateSpreadsheetSortCustomList(list.entries)).toMatchObject({
        ok: true,
      });
      expect(list.source).toBe('built-in');
      expect(Object.isFrozen(list)).toBe(true);
      expect(Object.isFrozen(list.entries)).toBe(true);
    }
    expect(
      spreadsheetSortCustomListsEqual(
        [' January ', 'FEBRUARY'],
        ['january', 'february'],
      ),
    ).toBe(true);
    expect(
      spreadsheetSortCustomListsEqual(
        SPREADSHEET_SORT_BUILT_IN_CUSTOM_LISTS[0]?.entries ?? [],
        SPREADSHEET_SORT_BUILT_IN_CUSTOM_LISTS[1]?.entries ?? [],
      ),
    ).toBe(false);
  });

  test('deduplicates and bounds the mounted-editor session registry', () => {
    const sessionLists = Array.from(
      { length: MAX_SPREADSHEET_SORT_SESSION_CUSTOM_LISTS + 4 },
      (_, index) =>
        createSpreadsheetSortCustomList(
          [`First ${index}`, `Second ${index}`],
          'session',
        ),
    ).filter((list) => list !== null);
    const merged = mergeSpreadsheetSortCustomLists([
      ...sessionLists,
      ...(sessionLists[0] ? [sessionLists[0]] : []),
    ]);

    expect(merged).toHaveLength(
      SPREADSHEET_SORT_BUILT_IN_CUSTOM_LISTS.length +
        MAX_SPREADSHEET_SORT_SESSION_CUSTOM_LISTS,
    );
    expect(Object.isFrozen(merged)).toBe(true);
  });
});
