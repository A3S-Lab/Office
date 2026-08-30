export type SpreadsheetSortTextMethod = 'pinyin' | 'stroke';

export interface SpreadsheetSortTextOptions {
  caseSensitive: boolean;
  textMethod: SpreadsheetSortTextMethod;
}

export const DEFAULT_SPREADSHEET_SORT_TEXT_OPTIONS: SpreadsheetSortTextOptions =
  Object.freeze({
    caseSensitive: false,
    textMethod: 'pinyin',
  });

export function isSpreadsheetSortTextMethod(
  value: unknown,
): value is SpreadsheetSortTextMethod {
  return value === 'pinyin' || value === 'stroke';
}

export function createSpreadsheetSortTextComparator(
  options: SpreadsheetSortTextOptions,
): ((left: string, right: string) => number) | null {
  try {
    const collator = new Intl.Collator(`zh-CN-u-co-${options.textMethod}`, {
      caseFirst: 'lower',
      numeric: false,
      sensitivity: options.caseSensitive ? 'case' : 'base',
      usage: 'sort',
    });
    if (collator.resolvedOptions().collation !== options.textMethod) {
      return null;
    }
    return collator.compare;
  } catch {
    return null;
  }
}
