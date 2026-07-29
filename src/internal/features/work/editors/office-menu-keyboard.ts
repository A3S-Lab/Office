import type { KeyboardEvent } from 'react';

const OFFICE_MENU_NAVIGATION_KEYS: readonly string[] = [
  'ArrowDown',
  'ArrowUp',
  'Home',
  'End',
] as const;

export function moveOfficeMenuFocus(
  event: KeyboardEvent<HTMLElement>,
): boolean {
  if (!OFFICE_MENU_NAVIGATION_KEYS.includes(event.key)) return false;
  const buttons = [
    ...event.currentTarget.querySelectorAll<HTMLButtonElement>(
      [
        'button[role="menuitem"]:not(:disabled)',
        'button[role="menuitemradio"]:not(:disabled)',
        'button[role="menuitemcheckbox"]:not(:disabled)',
      ].join(', '),
    ),
  ];
  if (!buttons.length) return false;
  event.preventDefault();
  const current = buttons.indexOf(document.activeElement as HTMLButtonElement);
  const next =
    event.key === 'Home'
      ? 0
      : event.key === 'End'
        ? buttons.length - 1
        : event.key === 'ArrowDown'
          ? (current + 1 + buttons.length) % buttons.length
          : (current - 1 + buttons.length) % buttons.length;
  buttons[next]?.focus();
  return true;
}
