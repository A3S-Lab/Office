import { type KeyboardEvent, useEffect, useState } from 'react';

const ACTIVE_OVERLAY_SELECTOR =
  '.ds-dialog-backdrop, .ds-popover.open, [role="menu"]';
const ESCAPE_CONSUMER_SELECTOR = '[data-office-escape-consumer="true"]';

export const OFFICE_TASK_PANE_MODAL_QUERY = '(max-width: 900px)';

type OfficeTaskPaneClose = () => unknown;

function officeTaskPaneOwnsEscape(
  event: Pick<globalThis.KeyboardEvent, 'defaultPrevented' | 'key' | 'target'>,
) {
  return (
    event.key === 'Escape' &&
    !event.defaultPrevented &&
    !(
      event.target instanceof Element &&
      event.target.closest(ESCAPE_CONSUMER_SELECTOR)
    ) &&
    !document.querySelector(ACTIVE_OVERLAY_SELECTOR)
  );
}

export function handleOfficeTaskPaneKeyDown(
  event: KeyboardEvent<HTMLElement>,
  onClose: OfficeTaskPaneClose,
) {
  if (!officeTaskPaneOwnsEscape(event.nativeEvent)) return;
  event.preventDefault();
  event.stopPropagation();
  void onClose();
}

export function useOfficeTaskPaneEscape(
  active: boolean,
  onClose: OfficeTaskPaneClose,
) {
  useEffect(() => {
    if (!active) return;
    const closeFromEscape = (event: globalThis.KeyboardEvent) => {
      if (!officeTaskPaneOwnsEscape(event)) return;
      event.preventDefault();
      event.stopPropagation();
      void onClose();
    };
    document.addEventListener('keydown', closeFromEscape, true);
    return () => document.removeEventListener('keydown', closeFromEscape, true);
  }, [active, onClose]);
}

export function useOfficeTaskPaneModal(
  query = OFFICE_TASK_PANE_MODAL_QUERY,
): boolean {
  const [matches, setMatches] = useState(() =>
    officeTaskPaneModalMatches(query),
  );

  useEffect(() => {
    if (
      typeof window === 'undefined' ||
      typeof window.matchMedia !== 'function'
    ) {
      return;
    }
    const mediaQuery = window.matchMedia(query);
    const update = () => setMatches(mediaQuery.matches);
    update();
    if (typeof mediaQuery.addEventListener === 'function') {
      mediaQuery.addEventListener('change', update);
      return () => mediaQuery.removeEventListener('change', update);
    }
    mediaQuery.addListener(update);
    return () => mediaQuery.removeListener(update);
  }, [query]);

  return matches;
}

function officeTaskPaneModalMatches(query: string): boolean {
  if (
    typeof window === 'undefined' ||
    typeof window.matchMedia !== 'function'
  ) {
    return false;
  }
  return window.matchMedia(query).matches;
}
