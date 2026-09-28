export type PlaygroundOfficeLocale = 'zh-CN' | 'en-US';

const SUPPORTED: readonly PlaygroundOfficeLocale[] = ['zh-CN', 'en-US'];

/**
 * Playground UI locale for editors. Defaults to product zh-CN; pass
 * `?locale=en-US` to exercise the English catalog in visual/functional gates.
 */
export function resolvePlaygroundLocale(
  search: string = typeof window !== 'undefined' ? window.location.search : '',
): PlaygroundOfficeLocale {
  const value = new URLSearchParams(search).get('locale');
  if (value && (SUPPORTED as readonly string[]).includes(value)) {
    return value as PlaygroundOfficeLocale;
  }
  return 'zh-CN';
}

/**
 * Opt-in A3S virtual grid for Spreadsheet (`?virtualGrid=1`). Off by default
 * so Fortune remains the GA path until Phase 4 exit is proven.
 */
export function resolvePlaygroundVirtualGrid(
  search: string = typeof window !== 'undefined' ? window.location.search : '',
): boolean {
  const value = new URLSearchParams(search).get('virtualGrid');
  return value === '1' || value === 'true';
}
