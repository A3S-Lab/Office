import type { KeyboardEvent } from 'react';

const OFFICE_MENU_NAVIGATION_KEYS: readonly string[] = [
  'ArrowDown',
  'ArrowUp',
  'Home',
  'End',
] as const;

const OFFICE_GRID_MENU_NAVIGATION_KEYS: readonly string[] = [
  ...OFFICE_MENU_NAVIGATION_KEYS,
  'ArrowLeft',
  'ArrowRight',
] as const;

const OFFICE_MENU_BUTTON_SELECTOR = [
  'button[role="menuitem"]:not(:disabled)',
  'button[role="menuitemradio"]:not(:disabled)',
  'button[role="menuitemcheckbox"]:not(:disabled)',
].join(', ');

export function moveOfficeMenuFocus(
  event: KeyboardEvent<HTMLElement>,
): boolean {
  if (!OFFICE_MENU_NAVIGATION_KEYS.includes(event.key)) return false;
  const buttons = officeMenuButtons(event.currentTarget);
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

export function moveOfficeGridMenuFocus(
  event: KeyboardEvent<HTMLElement>,
  columns: number,
): boolean {
  if (!OFFICE_GRID_MENU_NAVIGATION_KEYS.includes(event.key)) return false;
  const buttons = officeMenuButtons(event.currentTarget);
  if (!buttons.length) return false;
  event.preventDefault();
  const currentButton = document.activeElement as HTMLButtonElement;
  const current = Math.max(0, buttons.indexOf(currentButton));
  if (event.key === 'Home' || event.key === 'End') {
    buttons[event.key === 'Home' ? 0 : buttons.length - 1]?.focus();
    return true;
  }
  if (event.key === 'ArrowLeft' || event.key === 'ArrowRight') {
    const delta = event.key === 'ArrowRight' ? 1 : -1;
    buttons[(current + delta + buttons.length) % buttons.length]?.focus();
    return true;
  }

  const groups = [
    ...event.currentTarget.querySelectorAll<HTMLElement>(
      '[data-office-menu-grid]',
    ),
  ];
  const groupIndex = groups.findIndex((group) => group.contains(currentButton));
  if (groupIndex < 0) {
    const delta = event.key === 'ArrowDown' ? columns : -columns;
    buttons[(current + delta + buttons.length) % buttons.length]?.focus();
    return true;
  }
  const groupButtons = officeMenuButtons(groups[groupIndex]);
  const localIndex = Math.max(0, groupButtons.indexOf(currentButton));
  const column = localIndex % columns;
  if (event.key === 'ArrowDown') {
    const nextRow = localIndex - column + columns;
    if (nextRow < groupButtons.length) {
      groupButtons[
        Math.min(nextRow + column, groupButtons.length - 1)
      ]?.focus();
      return true;
    }
    const nextGroup = groups[(groupIndex + 1) % groups.length];
    const nextButtons = officeMenuButtons(nextGroup);
    nextButtons[Math.min(column, nextButtons.length - 1)]?.focus();
    return true;
  }
  const previousRow = localIndex - column - columns;
  if (previousRow >= 0) {
    groupButtons[previousRow + column]?.focus();
    return true;
  }
  const previousGroup =
    groups[(groupIndex - 1 + groups.length) % groups.length];
  const previousButtons = officeMenuButtons(previousGroup);
  const previousRowStart =
    Math.floor((previousButtons.length - 1) / columns) * columns;
  previousButtons[
    Math.min(previousRowStart + column, previousButtons.length - 1)
  ]?.focus();
  return true;
}

function officeMenuButtons(root: HTMLElement): HTMLButtonElement[] {
  return [
    ...root.querySelectorAll<HTMLButtonElement>(OFFICE_MENU_BUTTON_SELECTOR),
  ];
}
