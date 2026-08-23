export const DOCUMENT_HIDDEN_TEXT_ATTRIBUTE = 'data-office-hidden-text';

export const documentHiddenTextKeyboardShortcut = 'Mod-Shift-h';

export function normalizeDocumentHiddenText(value: unknown): boolean | null {
  if (value === true || value === 'true' || value === '1') return true;
  if (value === false || value === 'false' || value === '0') return false;
  return null;
}

export function documentHiddenTextFromElement(
  element: HTMLElement,
): boolean | null {
  return normalizeDocumentHiddenText(
    element.getAttribute(DOCUMENT_HIDDEN_TEXT_ATTRIBUTE),
  );
}

export function documentHiddenTextDomAttributes(
  value: unknown,
): Record<string, string> {
  const hiddenText = normalizeDocumentHiddenText(value);
  return hiddenText === null
    ? {}
    : { [DOCUMENT_HIDDEN_TEXT_ATTRIBUTE]: String(hiddenText) };
}
