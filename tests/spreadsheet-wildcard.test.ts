import { expect, test } from '@rstest/core';
import {
  workSpreadsheetFilterTextCharacters,
  workSpreadsheetFilterTextIsBounded,
} from '../src/internal/features/work/work-spreadsheet-filter-contract';
import {
  workSpreadsheetHasUnescapedWildcard,
  workSpreadsheetWildcardMatcher,
} from '../src/internal/features/work/work-spreadsheet-wildcard';

test('matches WPS wildcard expressions over normalized Unicode text', () => {
  expect(workSpreadsheetWildcardMatcher('')('')).toBe(true);
  expect(workSpreadsheetWildcardMatcher('')('King')).toBe(false);
  expect(workSpreadsheetWildcardMatcher('Ｋ?ＮＧ*')('KingSoft')).toBe(true);
  expect(workSpreadsheetWildcardMatcher('K?ng')('Kang')).toBe(true);
  expect(workSpreadsheetWildcardMatcher('K?ng')('K👑ng')).toBe(true);
  expect(workSpreadsheetWildcardMatcher('K??ng')('K👑ng')).toBe(false);
  expect(workSpreadsheetWildcardMatcher('*')('')).toBe(true);
  expect(workSpreadsheetWildcardMatcher('**')('multiple stars')).toBe(true);
  expect(workSpreadsheetWildcardMatcher('*draft*')('line\ndraft\nvalue')).toBe(
    true,
  );
});

test('treats WPS tilde escapes and regular-expression punctuation literally', () => {
  expect(workSpreadsheetWildcardMatcher('King~*')('King*')).toBe(true);
  expect(workSpreadsheetWildcardMatcher('King~?')('King?')).toBe(true);
  expect(workSpreadsheetWildcardMatcher('King~~')('King~')).toBe(true);
  expect(workSpreadsheetWildcardMatcher('King~x')('King~x')).toBe(true);
  expect(workSpreadsheetWildcardMatcher('[draft](1)')('[draft](1)')).toBe(true);
  expect(workSpreadsheetWildcardMatcher('King~*')('KingSoft')).toBe(false);
});

test('detects only wildcard operators that are not escaped', () => {
  expect(workSpreadsheetHasUnescapedWildcard('K?ng*')).toBe(true);
  expect(workSpreadsheetHasUnescapedWildcard('King~*')).toBe(false);
  expect(workSpreadsheetHasUnescapedWildcard('King~?')).toBe(false);
  expect(workSpreadsheetHasUnescapedWildcard('King~~')).toBe(false);
  expect(workSpreadsheetHasUnescapedWildcard('King~~*')).toBe(true);
  expect(workSpreadsheetHasUnescapedWildcard('King~~~*')).toBe(false);
});

test('enforces the shared expression ceiling in Unicode code points', () => {
  expect(workSpreadsheetFilterTextCharacters('👑')).toBe(1);
  expect(workSpreadsheetFilterTextIsBounded('👑'.repeat(32_767))).toBe(true);
  expect(workSpreadsheetFilterTextIsBounded('👑'.repeat(32_768))).toBe(false);
});

test('handles bounded long wildcard expressions without quadratic rescans', () => {
  const literalMatcher = workSpreadsheetWildcardMatcher(
    `*${'a'.repeat(16_382)}b`,
  );
  const singleMatcher = workSpreadsheetWildcardMatcher(
    `*?${'a'.repeat(16_380)}b*`,
  );
  const value = 'a'.repeat(32_767);
  expect(literalMatcher(value)).toBe(false);
  expect(singleMatcher(value)).toBe(false);
});
