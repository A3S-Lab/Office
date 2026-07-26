export const DEFAULT_DOCUMENT_PAGE_COLOR = '#ffffff';

export function normalizeDocumentPageColor(
  value: string | null | undefined,
): string | undefined {
  const trimmed = value?.trim();
  if (!trimmed) return undefined;
  if (/^#[0-9a-f]{3}$/i.test(trimmed)) {
    return `#${trimmed
      .slice(1)
      .split('')
      .map((part) => `${part}${part}`)
      .join('')}`.toLowerCase();
  }
  return /^#[0-9a-f]{6}$/i.test(trimmed) ? trimmed.toLowerCase() : undefined;
}

export function documentPageColor(value: string | null | undefined): string {
  return normalizeDocumentPageColor(value) ?? DEFAULT_DOCUMENT_PAGE_COLOR;
}
