export type WorkDocumentTextCase = 'none' | 'all-caps' | 'small-caps';

export type DocumentSelectionCase = 'lower' | 'upper' | 'title';

export const DOCUMENT_TEXT_CASE_ATTRIBUTE = 'data-office-text-case';

export const documentTextCaseKeyboardShortcuts = {
  allCaps: 'Mod-Shift-a',
  smallCaps: 'Mod-Shift-k',
  /** WPS/Word Shift+F3: cycle selected text lower → UPPER → Title. */
  changeCase: 'Shift-F3',
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

export function detectDocumentSelectionCase(
  text: string,
): DocumentSelectionCase {
  const letters = text.replace(/[^\p{L}]/gu, '');
  if (!letters) return 'lower';
  if (letters === letters.toLocaleUpperCase()) return 'upper';
  if (letters === letters.toLocaleLowerCase()) return 'lower';
  if (isTitleCaseSelection(text)) return 'title';
  return 'lower';
}

export function nextDocumentSelectionCase(
  current: DocumentSelectionCase,
): DocumentSelectionCase {
  if (current === 'lower') return 'upper';
  if (current === 'upper') return 'title';
  return 'lower';
}

export function transformDocumentSelectionCase(
  text: string,
  next: DocumentSelectionCase,
): string {
  if (next === 'upper') return text.toLocaleUpperCase();
  if (next === 'lower') return text.toLocaleLowerCase();
  return toTitleCaseSelection(text);
}

export function cycleDocumentSelectionCaseText(text: string): string {
  return transformDocumentSelectionCase(
    text,
    nextDocumentSelectionCase(detectDocumentSelectionCase(text)),
  );
}

function isTitleCaseSelection(text: string): boolean {
  const words = text.match(/\p{L}[\p{L}'’]*/gu);
  if (!words?.length) return false;
  return words.every((word) => {
    const first = word.charAt(0);
    const rest = word.slice(1);
    return (
      first === first.toLocaleUpperCase() && rest === rest.toLocaleLowerCase()
    );
  });
}

function toTitleCaseSelection(text: string): string {
  return text.replace(/\p{L}[\p{L}'’]*/gu, (word) => {
    const first = word.charAt(0).toLocaleUpperCase();
    const rest = word.slice(1).toLocaleLowerCase();
    return `${first}${rest}`;
  });
}
