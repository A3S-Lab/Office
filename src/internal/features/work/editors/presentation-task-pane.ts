import { type KeyboardEvent, useEffect } from 'react';

const OVERLAY_SELECTOR = '.ds-dialog-backdrop, .ds-popover.open, [role="menu"]';

function presentationTaskPaneOwnsEscape(
  event: Pick<globalThis.KeyboardEvent, 'defaultPrevented' | 'key'>,
) {
  return (
    event.key === 'Escape' &&
    !event.defaultPrevented &&
    !document.querySelector(OVERLAY_SELECTOR)
  );
}

export function handlePresentationTaskPaneKeyDown(
  event: KeyboardEvent<HTMLElement>,
  onClose: () => void,
) {
  if (!presentationTaskPaneOwnsEscape(event.nativeEvent)) return;
  event.preventDefault();
  event.stopPropagation();
  onClose();
}

export function usePresentationTaskPaneEscape(
  active: boolean,
  onClose: () => void,
) {
  useEffect(() => {
    if (!active) return;
    const closeFromEscape = (event: globalThis.KeyboardEvent) => {
      if (!presentationTaskPaneOwnsEscape(event)) return;
      event.preventDefault();
      event.stopPropagation();
      onClose();
    };
    document.addEventListener('keydown', closeFromEscape, true);
    return () => document.removeEventListener('keydown', closeFromEscape, true);
  }, [active, onClose]);
}
