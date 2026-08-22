export const spreadsheetFormatCellsTabs = [
  { id: 'number', label: '数字' },
  { id: 'alignment', label: '对齐' },
  { id: 'font', label: '字体' },
  { id: 'border', label: '边框' },
  { id: 'fill', label: '填充' },
  { id: 'protection', label: '保护' },
] as const;

export type SpreadsheetFormatCellsTabId =
  (typeof spreadsheetFormatCellsTabs)[number]['id'];

export type SpreadsheetFormatCellsInitialFocus = 'fontFamily' | 'fontSize';

export type SpreadsheetFormatCellsOpenIntent =
  | {
      tab: 'font';
      focus?: SpreadsheetFormatCellsInitialFocus;
    }
  | {
      tab: Exclude<SpreadsheetFormatCellsTabId, 'font'>;
      focus?: never;
    };

export const defaultSpreadsheetFormatCellsOpenIntent = {
  tab: 'number',
} as const satisfies SpreadsheetFormatCellsOpenIntent;

const spreadsheetFormatCellsTabIds = new Set<SpreadsheetFormatCellsTabId>(
  spreadsheetFormatCellsTabs.map(({ id }) => id),
);

export function normalizeSpreadsheetFormatCellsOpenIntent(
  value: unknown,
): SpreadsheetFormatCellsOpenIntent | null {
  if (value === undefined) return defaultSpreadsheetFormatCellsOpenIntent;
  if (!isRecord(value)) return null;
  const keys = Object.keys(value);
  if (keys.some((key) => key !== 'tab' && key !== 'focus')) return null;
  if (
    typeof value.tab !== 'string' ||
    !spreadsheetFormatCellsTabIds.has(value.tab as SpreadsheetFormatCellsTabId)
  ) {
    return null;
  }
  if (value.tab !== 'font') {
    return value.focus === undefined
      ? {
          tab: value.tab as Exclude<SpreadsheetFormatCellsTabId, 'font'>,
        }
      : null;
  }
  if (value.focus === undefined) return { tab: 'font' };
  return value.focus === 'fontFamily' || value.focus === 'fontSize'
    ? { tab: 'font', focus: value.focus }
    : null;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return Boolean(value && typeof value === 'object' && !Array.isArray(value));
}
