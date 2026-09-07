import type { KeyboardEvent } from 'react';

/**
 * Apply the shared WPS-style horizontal focus contract to an Office toolbar.
 *
 * Form controls keep their own cursor and option navigation. Only ordinary
 * toolbar buttons participate in the roving focus behavior.
 */
export function moveOfficeToolbarFocus(
  event: KeyboardEvent<HTMLElement>,
): void {
  if (!['ArrowLeft', 'ArrowRight', 'Home', 'End'].includes(event.key)) return;

  const activeElement = event.currentTarget.ownerDocument.activeElement;
  const source =
    event.target instanceof HTMLButtonElement ? event.target : activeElement;
  if (
    !(source instanceof HTMLButtonElement) ||
    source.getAttribute('role') === 'combobox' ||
    source.getAttribute('role') === 'slider'
  ) {
    return;
  }

  const controls = [
    ...event.currentTarget.querySelectorAll<HTMLElement>(
      'button:not(:disabled), [role="slider"]:not([aria-disabled="true"])',
    ),
  ];
  if (!controls.length) return;
  const currentIndex = controls.indexOf(source);
  if (currentIndex < 0) return;
  const nextIndex =
    event.key === 'Home'
      ? 0
      : event.key === 'End'
        ? controls.length - 1
        : event.key === 'ArrowRight'
          ? (currentIndex + 1) % controls.length
          : (currentIndex - 1 + controls.length) % controls.length;
  event.preventDefault();
  controls[nextIndex]?.focus();
}
