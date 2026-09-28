import {
  officeMessage,
  resolveOfficeMessages,
} from '../../../i18n/office-locale';
import type {
  WorkSpreadsheetDateSystem,
  WorkSpreadsheetSheet,
} from '../work-types';
import { normalizedWorkSpreadsheetAutoFilterRange } from '../work-spreadsheet-auto-filter';
import { spreadsheetAutoFilterColumnIsNumeric } from './spreadsheet-auto-filter';

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
  dateSystem: WorkSpreadsheetDateSystem = '1900',
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
    const catalog = resolveOfficeMessages();
    trigger.dataset.filterLabel = label;
    trigger.setAttribute(
      'aria-label',
      officeMessage(catalog, 'spreadsheet.autoFilter.triggerAria', { label }),
    );
    trigger.setAttribute('aria-haspopup', 'dialog');
    trigger.setAttribute('aria-expanded', 'false');
    trigger.setAttribute(
      'title',
      officeMessage(catalog, 'spreadsheet.autoFilter.triggerTitle', { label }),
    );
  }

  const menu = container.querySelector<HTMLElement>('.fortune-filter-menu');
  if (!menu) return null;
  const trigger =
    invoker ?? triggers.find((candidate) => candidate.matches(':focus'));
  const catalog = resolveOfficeMessages();
  const label =
    trigger?.dataset.filterLabel ||
    officeMessage(catalog, 'spreadsheet.autoFilter.columnFallback');
  menu.setAttribute('role', 'dialog');
  menu.dataset.officeShortcuts = 'ignore';
  menu.setAttribute(
    'aria-label',
    officeMessage(catalog, 'spreadsheet.autoFilter.menuAria', { label }),
  );
  menu.setAttribute('aria-modal', 'false');
  trigger?.setAttribute('aria-expanded', 'true');

  const search = menu.querySelector<HTMLInputElement>(
    '.filtermenu-input-container input, input:not([type="checkbox"])',
  );
  search?.setAttribute(
    'aria-label',
    officeMessage(catalog, 'spreadsheet.autoFilter.searchAria'),
  );
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
    checkbox.setAttribute(
      'aria-label',
      officeMessage(catalog, 'spreadsheet.autoFilter.showValueAria', {
        label:
          itemLabel ||
          officeMessage(catalog, 'spreadsheet.autoFilter.valueFallback'),
      }),
    );
  }
  for (const action of menu.querySelectorAll<HTMLElement>(
    '.luckysheet-cols-menuitem, .fortune-byvalue-btn, .button-basic',
  )) {
    action.setAttribute('role', 'button');
    const actionLabel = action.textContent?.replace(/\s+/g, ' ').trim();
    if (actionLabel) action.setAttribute('aria-label', actionLabel);
    if (
      actionLabel === '\u6309\u6761\u4ef6\u8fc7\u6ee4' ||
      actionLabel === '\u6309\u689d\u4ef6\u904e\u6ffe' ||
      actionLabel === 'Filter by condition'
    ) {
      action.dataset.a3sAutoFilterCondition = '';
      action.setAttribute('aria-haspopup', 'dialog');
    }
  }
  synchronizeSpreadsheetAutoFilterRankAction(menu, sheet, trigger, dateSystem);
  return menu;
}

function synchronizeSpreadsheetAutoFilterRankAction(
  menu: HTMLElement,
  sheet: WorkSpreadsheetSheet | undefined,
  trigger: HTMLElement | undefined,
  dateSystem: WorkSpreadsheetDateSystem,
): void {
  const existing = menu.querySelector<HTMLElement>(
    '[data-a3s-auto-filter-rank]',
  );
  const range = normalizedWorkSpreadsheetAutoFilterRange(sheet?.filter_select);
  const column = Number(trigger?.dataset.filterColumn);
  const numeric = Boolean(
    sheet &&
      range &&
      Number.isSafeInteger(column) &&
      spreadsheetAutoFilterColumnIsNumeric(sheet, range, column, dateSystem),
  );
  if (!numeric) {
    existing?.remove();
    return;
  }
  if (existing) {
    existing.dataset.filterColumn = String(column);
    return;
  }
  const action = document.createElement('div');
  action.className = 'luckysheet-cols-menuitem';
  action.dataset.a3sAutoFilterRank = '';
  action.dataset.filterColumn = String(column);
  action.dataset.officeShortcuts = 'ignore';
  action.tabIndex = 0;
  const catalog = resolveOfficeMessages();
  const top10 = officeMessage(catalog, 'spreadsheet.autoFilter.top10');
  action.textContent = top10;
  action.setAttribute('role', 'button');
  action.setAttribute('aria-label', top10);
  action.setAttribute('aria-haspopup', 'dialog');
  const condition = menu.querySelector<HTMLElement>(
    '[data-a3s-auto-filter-condition]',
  );
  condition ? condition.before(action) : menu.prepend(action);
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

export function spreadsheetAutoFilterRankAction(
  target: EventTarget | null,
): HTMLElement | null {
  return target instanceof Element
    ? target.closest<HTMLElement>('[data-a3s-auto-filter-rank]')
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
  if (label) return label;
  return officeMessage(
    resolveOfficeMessages(),
    'spreadsheet.autoFilter.columnWithName',
    { column: spreadsheetColumnName(column) },
  );
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
