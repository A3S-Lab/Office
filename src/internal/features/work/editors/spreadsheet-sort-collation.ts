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

const SPREADSHEET_SORT_COLLATION_PROBES: Readonly<
  Record<SpreadsheetSortTextMethod, readonly string[]>
> = Object.freeze({
  pinyin: Object.freeze(['阿', '丁', '王', '赵']),
  stroke: Object.freeze(['丁', '王', '安', '阿', '赵']),
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
    if (!spreadsheetSortCollatorSupportsMethod(collator, options.textMethod)) {
      return null;
    }
    return collator.compare;
  } catch {
    return null;
  }
}

function spreadsheetSortCollatorSupportsMethod(
  collator: Intl.Collator,
  textMethod: SpreadsheetSortTextMethod,
): boolean {
  const probe = SPREADSHEET_SORT_COLLATION_PROBES[textMethod];
  for (let index = 1; index < probe.length; index += 1) {
    const previous = probe[index - 1];
    const current = probe[index];
    if (previous === undefined || current === undefined) return false;
    if (collator.compare(previous, current) >= 0) return false;
  }
  return true;
}
