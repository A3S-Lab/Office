/**
 * WPS/PowerPoint slideshow digit buffer: type a 1-based slide number, then
 * Enter jumps. Escape clears a non-empty buffer without exiting the show.
 */

const MAX_GOTO_DIGITS = 4;

export function appendPresentationGotoDigit(
  buffer: string,
  key: string,
): string | null {
  if (!/^\d$/.test(key)) return null;
  if (buffer.length >= MAX_GOTO_DIGITS) return buffer;
  return `${buffer}${key}`;
}

/** Resolve a typed digit buffer to a 0-based slide index, or null. */
export function resolvePresentationGotoIndex(
  buffer: string,
  slideCount: number,
): number | null {
  if (!buffer || slideCount <= 0) return null;
  const page = Number.parseInt(buffer, 10);
  if (!Number.isFinite(page) || page < 1) return null;
  return Math.min(page, slideCount) - 1;
}

export type PresentationGotoEscapeAction = 'clear-buffer' | 'exit';

export function presentationGotoEscapeAction(
  buffer: string,
): PresentationGotoEscapeAction {
  return buffer.length > 0 ? 'clear-buffer' : 'exit';
}
