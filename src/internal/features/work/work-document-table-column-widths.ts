export type DocumentTableColumnWidthType = 'pixels' | 'percent';

export const MIN_DOCUMENT_TABLE_COLUMN_PERCENT = 1;
export const MAX_DOCUMENT_TABLE_COLUMN_PERCENT = 100;

export function normalizeDocumentTableColumnPercent(
  value: unknown,
): number | null {
  const number = Number(value);
  if (
    !Number.isFinite(number) ||
    number < MIN_DOCUMENT_TABLE_COLUMN_PERCENT ||
    number > MAX_DOCUMENT_TABLE_COLUMN_PERCENT
  ) {
    return null;
  }
  return Math.round(number * 100) / 100;
}

export function normalizeDocumentTableColumnPercentages(
  value: unknown,
  expectedCount?: number,
): number[] | null {
  const values = Array.isArray(value)
    ? value
    : typeof value === 'string'
      ? value.split(',')
      : null;
  if (
    !values ||
    (expectedCount !== undefined && values.length !== expectedCount)
  ) {
    return null;
  }
  const normalized = values.map(normalizeDocumentTableColumnPercent);
  return normalized.every((entry) => entry !== null)
    ? (normalized as number[])
    : null;
}

export function documentTableColumnPercentagesFromElement(
  cell: HTMLElement,
): number[] | null {
  return normalizeDocumentTableColumnPercentages(
    cell.dataset.officeColumnWidthsPercent,
    Math.max(1, Number(cell.getAttribute('colspan') ?? 1)),
  );
}

export function renderDocumentTableColumnPercentages(
  value: unknown,
): Record<string, string> {
  const percentages = normalizeDocumentTableColumnPercentages(value);
  return percentages
    ? { 'data-office-column-widths-percent': percentages.join(',') }
    : {};
}
