import type { WorkSpreadsheetSheet } from '../work-types';

export const SPREADSHEET_AUTO_FILTER_MENU_ITEMS: string[] = [
  'filter-by-condition',
  '|',
  'filter-by-color',
  '|',
  'filter-by-value',
];

export function enhanceSpreadsheetAutoFilterSurface(
  container: HTMLElement,
  sheet: WorkSpreadsheetSheet | undefined,
  invoker: HTMLElement | null,
): HTMLElement | null {
  const range = sheet?.filter_select;
  const startRow = range?.row?.[0];
  const startColumn = range?.column?.[0];
  const triggers = [
    ...container.querySelectorAll<HTMLElement>('.luckysheet-filter-options'),
  ];
  for (const [index, trigger] of triggers.entries()) {
    const column = Number(startColumn) + index;
    if (!Number.isFinite(column) || !Number.isFinite(startRow)) continue;
    const label = spreadsheetAutoFilterColumnLabel(
      sheet,
      Number(startRow),
      column,
    );
    trigger.dataset.filterColumn = String(column);
    trigger.dataset.officeShortcuts = 'ignore';
    trigger.setAttribute('role', 'button');
    trigger.setAttribute('aria-label', `${label} 筛选`);
    trigger.setAttribute('aria-haspopup', 'dialog');
    trigger.setAttribute('aria-expanded', 'false');
    trigger.setAttribute('title', `${label} 筛选（Alt+↓）`);
  }

  const menu = container.querySelector<HTMLElement>('.fortune-filter-menu');
  if (!menu) return null;
  const trigger =
    invoker ?? triggers.find((candidate) => candidate.matches(':focus'));
  const triggerLabel = trigger
    ?.getAttribute('aria-label')
    ?.replace(/\s*筛选$/, '');
  const label = triggerLabel || '列';
  menu.setAttribute('role', 'dialog');
  menu.dataset.officeShortcuts = 'ignore';
  menu.setAttribute('aria-label', `${label} 筛选`);
  menu.setAttribute('aria-modal', 'false');
  trigger?.setAttribute('aria-expanded', 'true');

  const search = menu.querySelector<HTMLInputElement>(
    '.filtermenu-input-container input, input:not([type="checkbox"])',
  );
  search?.setAttribute('aria-label', '搜索筛选值');
  for (const checkbox of menu.querySelectorAll<HTMLInputElement>(
    'input[type="checkbox"]',
  )) {
    const item = checkbox.closest('.select-item');
    const itemLabel = item
      ? [...item.children]
          .filter(
            (child) =>
              child !== checkbox &&
              !child.classList.contains('count') &&
              !child.classList.contains('filter-caret'),
          )
          .map((child) => child.textContent?.trim())
          .find(Boolean)
      : undefined;
    checkbox.setAttribute('aria-label', `显示 ${itemLabel || '筛选值'}`);
  }
  for (const action of menu.querySelectorAll<HTMLElement>(
    '.luckysheet-cols-menuitem, .fortune-byvalue-btn, .button-basic',
  )) {
    action.setAttribute('role', 'button');
    const actionLabel = action.textContent?.replace(/\s+/g, ' ').trim();
    if (actionLabel) action.setAttribute('aria-label', actionLabel);
    if (
      actionLabel === '按条件过滤' ||
      actionLabel === '按條件過濾' ||
      actionLabel === 'Filter by condition'
    ) {
      action.dataset.a3sAutoFilterCondition = '';
      action.setAttribute('aria-haspopup', 'dialog');
    }
  }
  return menu;
}

export function focusSpreadsheetAutoFilterMenu(container: HTMLElement): void {
  requestAnimationFrame(() => {
    const menu = container.querySelector<HTMLElement>('.fortune-filter-menu');
    if (!menu) return;
    spreadsheetAutoFilterMenuFocusable(menu)[0]?.focus({
      preventScroll: true,
    });
  });
}

export function spreadsheetAutoFilterMenuFocusable(
  menu: HTMLElement,
): HTMLElement[] {
  return [
    ...menu.querySelectorAll<HTMLElement>(
      '.luckysheet-cols-menuitem, .fortune-byvalue-btn, .select-item[tabindex], input:not([disabled]), .button-basic',
    ),
  ].filter((element) => element.tabIndex >= 0 && !element.hidden);
}

export function spreadsheetAutoFilterTrigger(
  target: EventTarget | null,
): HTMLElement | null {
  return target instanceof Element
    ? target.closest<HTMLElement>('.luckysheet-filter-options')
    : null;
}

export function spreadsheetAutoFilterMenu(
  target: EventTarget | null,
): HTMLElement | null {
  return target instanceof Element
    ? target.closest<HTMLElement>('.fortune-filter-menu')
    : null;
}

export function spreadsheetAutoFilterConditionAction(
  target: EventTarget | null,
): HTMLElement | null {
  return target instanceof Element
    ? target.closest<HTMLElement>('[data-a3s-auto-filter-condition]')
    : null;
}

export function spreadsheetAutoFilterColumnLabel(
  sheet: WorkSpreadsheetSheet | undefined,
  row: number,
  column: number,
): string {
  const cell = sheet?.data
    ? sheet.data[row]?.[column]
    : sheet?.celldata?.find(
        (candidate) => candidate.r === row && candidate.c === column,
      )?.v;
  const value = cell?.m ?? cell?.v;
  const label =
    value === undefined || value === null ? '' : String(value).trim();
  return label || `列 ${spreadsheetColumnName(column)}`;
}

function spreadsheetColumnName(column: number): string {
  let value = Math.max(0, Math.floor(column)) + 1;
  let name = '';
  while (value > 0) {
    value -= 1;
    name = String.fromCharCode(65 + (value % 26)) + name;
    value = Math.floor(value / 26);
  }
  return name;
}
