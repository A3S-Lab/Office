export type WorkDocumentTextCase = 'none' | 'all-caps' | 'small-caps';

export const DOCUMENT_TEXT_CASE_ATTRIBUTE = 'data-office-text-case';

export const documentTextCaseKeyboardShortcuts = {
  allCaps: 'Mod-Shift-a',
  smallCaps: 'Mod-Shift-k',
} as const;

export function normalizeDocumentTextCase(
  value: unknown,
): WorkDocumentTextCase | null {
  return value === 'none' || value === 'all-caps' || value === 'small-caps'
    ? value
    : null;
}

export function documentTextCaseFromWordFlags(
  allCaps: boolean | undefined,
  smallCaps: boolean | undefined,
): WorkDocumentTextCase | null {
  if (allCaps === undefined && smallCaps === undefined) return null;
  if (allCaps) return 'all-caps';
  if (smallCaps) return 'small-caps';
  return 'none';
}

export function applyDocumentTextCaseStyle(
  element: HTMLElement,
  textCase: WorkDocumentTextCase,
): void {
  element.dataset.officeTextCase = textCase;
  element.style.textTransform = textCase === 'all-caps' ? 'uppercase' : 'none';
  element.style.fontVariantCaps =
    textCase === 'small-caps' ? 'small-caps' : 'normal';
}

export function documentTextCaseCss(textCase: WorkDocumentTextCase): string {
  return [
    `text-transform: ${textCase === 'all-caps' ? 'uppercase' : 'none'}`,
    `font-variant-caps: ${textCase === 'small-caps' ? 'small-caps' : 'normal'}`,
  ].join('; ');
}
