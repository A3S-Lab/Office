import type { Selection } from '@fortune-sheet/core';
import { type RefObject, useCallback, useEffect, useMemo, useRef } from 'react';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetSheet,
} from '../work-types';
import {
  spreadsheetAutoFilterHeaderColumn,
  toggleSpreadsheetAutoFilter,
} from './spreadsheet-auto-filter';
import type {
  SpreadsheetAutoFilterCommandPort,
  SpreadsheetWorkbookCommandPort,
} from './spreadsheet-command-controller';

export interface UseSpreadsheetAutoFilterOptions {
  canvasRef: RefObject<HTMLElement | null>;
  content: WorkSpreadsheetContent;
  editable: boolean;
  mountRevision?: number;
  onChange: (content: WorkSpreadsheetContent) => void;
  selection: Selection;
  sheetId: string;
  workbook?: Pick<SpreadsheetWorkbookCommandPort, 'getSelection'> | null;
}

export interface SpreadsheetAutoFilterController {
  active: boolean;
  commandPort: SpreadsheetAutoFilterCommandPort;
  reserveAltKey: () => boolean;
  status: string;
}

export function useSpreadsheetAutoFilter({
  canvasRef,
  content,
  editable,
  mountRevision,
  onChange,
  selection,
  sheetId,
  workbook = null,
}: UseSpreadsheetAutoFilterOptions): SpreadsheetAutoFilterController {
  const contentRef = useRef(content);
  const editableRef = useRef(editable);
  const onChangeRef = useRef(onChange);
  const selectionRef = useRef(selection);
  const sheetIdRef = useRef(sheetId);
  const workbookRef = useRef(workbook);
  const keyboardInvokerRef = useRef<HTMLElement | null>(null);
  const restoreInvokerFocusRef = useRef(false);
  const menuWasOpenRef = useRef(false);
  contentRef.current = content;
  editableRef.current = editable;
  onChangeRef.current = onChange;
  selectionRef.current = selection;
  sheetIdRef.current = sheetId;
  workbookRef.current = workbook;

  const sheet = content.sheets.find((candidate) => candidate.id === sheetId);
  const active = Boolean(sheet?.filter_select);
  const canToggle = Boolean(editable && sheet);
  const canOpenMenu = Boolean(editable && sheet?.filter_select);

  const toggle = useCallback((): boolean => {
    if (!editableRef.current) return false;
    const next = toggleSpreadsheetAutoFilter(
      contentRef.current,
      sheetIdRef.current,
      currentSpreadsheetAutoFilterSelection(
        selectionRef.current,
        workbookRef.current,
      ),
    );
    if (!next) return false;
    contentRef.current = next;
    onChangeRef.current(next);
    return true;
  }, []);

  const openMenu = useCallback((): boolean => {
    if (!editableRef.current) return false;
    const container = canvasRef.current;
    const activeSheet = contentRef.current.sheets.find(
      (candidate) => candidate.id === sheetIdRef.current,
    );
    const column = spreadsheetAutoFilterHeaderColumn(
      activeSheet,
      currentSpreadsheetAutoFilterSelection(
        selectionRef.current,
        workbookRef.current,
      ),
    );
    const startColumn = activeSheet?.filter_select?.column?.[0];
    if (!container || column === null || !Number.isFinite(startColumn)) {
      return false;
    }
    enhanceSpreadsheetAutoFilterSurface(
      container,
      activeSheet,
      keyboardInvokerRef.current,
    );
    const triggers = [
      ...container.querySelectorAll<HTMLElement>('.luckysheet-filter-options'),
    ];
    const trigger = triggers[column - Number(startColumn)];
    if (!trigger) return false;
    keyboardInvokerRef.current = trigger;
    restoreInvokerFocusRef.current = true;
    trigger.click();
    enhanceSpreadsheetAutoFilterSurface(container, activeSheet, trigger);
    focusSpreadsheetAutoFilterMenu(container);
    return true;
  }, [canvasRef]);

  const reserveAltKey = useCallback((): boolean => {
    if (!editableRef.current) return false;
    const activeSheet = contentRef.current.sheets.find(
      (candidate) => candidate.id === sheetIdRef.current,
    );
    return (
      spreadsheetAutoFilterHeaderColumn(
        activeSheet,
        currentSpreadsheetAutoFilterSelection(
          selectionRef.current,
          workbookRef.current,
        ),
      ) !== null
    );
  }, []);

  useEffect(() => {
    const container = canvasRef.current;
    if (!container || !editable) return;
    const enhance = () => {
      const activeSheet = contentRef.current.sheets.find(
        (candidate) => candidate.id === sheetIdRef.current,
      );
      const menu = enhanceSpreadsheetAutoFilterSurface(
        container,
        activeSheet,
        keyboardInvokerRef.current,
      );
      const menuOpen = Boolean(menu);
      if (
        menuWasOpenRef.current &&
        !menuOpen &&
        restoreInvokerFocusRef.current
      ) {
        const invoker = keyboardInvokerRef.current;
        requestAnimationFrame(() => {
          if (invoker?.isConnected) invoker.focus({ preventScroll: true });
        });
        restoreInvokerFocusRef.current = false;
      }
      menuWasOpenRef.current = menuOpen;
    };
    const handlePointerDown = (event: PointerEvent) => {
      const trigger = spreadsheetAutoFilterTrigger(event.target);
      if (!trigger) return;
      keyboardInvokerRef.current = trigger;
      restoreInvokerFocusRef.current = false;
    };
    const handleKeyDown = (event: KeyboardEvent) => {
      const trigger = spreadsheetAutoFilterTrigger(event.target);
      if (trigger && (event.key === 'Enter' || event.key === ' ')) {
        event.preventDefault();
        event.stopPropagation();
        keyboardInvokerRef.current = trigger;
        restoreInvokerFocusRef.current = true;
        trigger.click();
        enhanceSpreadsheetAutoFilterSurface(
          container,
          contentRef.current.sheets.find(
            (candidate) => candidate.id === sheetIdRef.current,
          ),
          trigger,
        );
        focusSpreadsheetAutoFilterMenu(container);
        return;
      }
      const menu = spreadsheetAutoFilterMenu(event.target);
      if (!menu) return;
      if (event.key === 'Escape') {
        event.preventDefault();
        event.stopPropagation();
        menu.querySelector<HTMLElement>('.button-default')?.click();
        if (!menu.isConnected && restoreInvokerFocusRef.current) {
          const invoker = keyboardInvokerRef.current;
          invoker?.setAttribute('aria-expanded', 'false');
          requestAnimationFrame(() => {
            if (invoker?.isConnected) invoker.focus({ preventScroll: true });
          });
          menuWasOpenRef.current = false;
          restoreInvokerFocusRef.current = false;
        }
        return;
      }
      const target = event.target;
      if (
        (event.key === 'Enter' || event.key === ' ') &&
        target instanceof HTMLElement &&
        !(target instanceof HTMLInputElement) &&
        target.tabIndex >= 0
      ) {
        event.preventDefault();
        event.stopPropagation();
        target.click();
        return;
      }
      if (
        (event.key === 'ArrowDown' || event.key === 'ArrowUp') &&
        !(target instanceof HTMLInputElement && target.type !== 'checkbox')
      ) {
        const focusable = spreadsheetAutoFilterMenuFocusable(menu);
        const current = focusable.indexOf(
          document.activeElement as HTMLElement,
        );
        if (!focusable.length) return;
        event.preventDefault();
        event.stopPropagation();
        const offset = event.key === 'ArrowDown' ? 1 : -1;
        const next =
          (Math.max(current, 0) + offset + focusable.length) % focusable.length;
        focusable[next]?.focus({ preventScroll: true });
      }
    };

    enhance();
    container.addEventListener('pointerdown', handlePointerDown, true);
    container.addEventListener('keydown', handleKeyDown, true);
    if (typeof MutationObserver === 'undefined') {
      return () => {
        container.removeEventListener('pointerdown', handlePointerDown, true);
        container.removeEventListener('keydown', handleKeyDown, true);
      };
    }
    const observer = new MutationObserver(enhance);
    observer.observe(container, { childList: true, subtree: true });
    return () => {
      observer.disconnect();
      container.removeEventListener('pointerdown', handlePointerDown, true);
      container.removeEventListener('keydown', handleKeyDown, true);
    };
  }, [canvasRef, editable, mountRevision]);

  const commandPort = useMemo<SpreadsheetAutoFilterCommandPort>(
    () => ({
      active,
      canOpenMenu,
      canToggle,
      openMenu,
      toggle,
    }),
    [active, canOpenMenu, canToggle, openMenu, toggle],
  );
  return {
    active,
    commandPort,
    reserveAltKey,
    status: active ? '自动筛选已开启；在表头按 Alt+向下箭头打开筛选菜单。' : '',
  };
}

function currentSpreadsheetAutoFilterSelection(
  fallback: Selection,
  workbook: Pick<SpreadsheetWorkbookCommandPort, 'getSelection'> | null,
): Selection {
  const live = workbook?.getSelection()?.at(-1);
  if (!live) return fallback;
  return {
    ...fallback,
    ...live,
    row: [...live.row],
    column: [...live.column],
  };
}

function enhanceSpreadsheetAutoFilterSurface(
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
  }
  return menu;
}

function focusSpreadsheetAutoFilterMenu(container: HTMLElement): void {
  requestAnimationFrame(() => {
    const menu = container.querySelector<HTMLElement>('.fortune-filter-menu');
    if (!menu) return;
    spreadsheetAutoFilterMenuFocusable(menu)[0]?.focus({
      preventScroll: true,
    });
  });
}

function spreadsheetAutoFilterMenuFocusable(menu: HTMLElement): HTMLElement[] {
  return [
    ...menu.querySelectorAll<HTMLElement>(
      '.luckysheet-cols-menuitem, .fortune-byvalue-btn, .select-item[tabindex], input:not([disabled]), .button-basic',
    ),
  ].filter((element) => element.tabIndex >= 0 && !element.hidden);
}

function spreadsheetAutoFilterTrigger(
  target: EventTarget | null,
): HTMLElement | null {
  return target instanceof Element
    ? target.closest<HTMLElement>('.luckysheet-filter-options')
    : null;
}

function spreadsheetAutoFilterMenu(
  target: EventTarget | null,
): HTMLElement | null {
  return target instanceof Element
    ? target.closest<HTMLElement>('.fortune-filter-menu')
    : null;
}

function spreadsheetAutoFilterColumnLabel(
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
