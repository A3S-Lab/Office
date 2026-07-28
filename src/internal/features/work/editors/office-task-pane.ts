import { type KeyboardEvent, useEffect } from 'react';

const ACTIVE_OVERLAY_SELECTOR =
  '.ds-dialog-backdrop, .ds-popover.open, [role="menu"]';

type OfficeTaskPaneClose = () => unknown;

function officeTaskPaneOwnsEscape(
  event: Pick<globalThis.KeyboardEvent, 'defaultPrevented' | 'key'>,
) {
  return (
    event.key === 'Escape' &&
    !event.defaultPrevented &&
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
