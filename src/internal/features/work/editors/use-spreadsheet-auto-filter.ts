import type { Selection } from '@fortune-sheet/core';
import { showToast } from '../../../state/app-state';
import {
  createElement,
  type ReactNode,
  type RefObject,
  useCallback,
  useEffect,
  useMemo,
  useRef,
  useState,
} from 'react';
import type {
  WorkSpreadsheetFilterCriteria,
  WorkSpreadsheetContent,
  WorkSpreadsheetSheet,
} from '../work-types';
import {
  spreadsheetAutoFilterCriteria,
  spreadsheetAutoFilterHeaderColumn,
  toggleSpreadsheetAutoFilter,
} from './spreadsheet-auto-filter';
import {
  SpreadsheetAutoFilterConditionDialog,
  type SpreadsheetAutoFilterConditionDialogSource,
} from './spreadsheet-auto-filter-condition-dialog';
import type {
  SpreadsheetAutoFilterCommandPort,
  SpreadsheetEditorCommands,
  SpreadsheetWorkbookCommandPort,
} from './spreadsheet-command-controller';
import {
  enhanceSpreadsheetAutoFilterSurface,
  focusSpreadsheetAutoFilterMenu,
  spreadsheetAutoFilterColumnLabel,
  spreadsheetAutoFilterConditionAction,
  spreadsheetAutoFilterMenu,
  spreadsheetAutoFilterMenuFocusable,
  spreadsheetAutoFilterTrigger,
} from './spreadsheet-auto-filter-menu';
import { spreadsheetSheetCellReader } from './spreadsheet-current-region';

export interface UseSpreadsheetAutoFilterOptions {
  canvasRef: RefObject<HTMLElement | null>;
  commandsRef: { current: SpreadsheetEditorCommands | null };
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
  dialog: ReactNode;
  reserveAltKey: () => boolean;
  selectionForChange: () => SpreadsheetAutoFilterSelectionState | null;
  status: string;
}

export interface SpreadsheetAutoFilterSelectionState {
  selections: Selection[];
  sheetId: string;
}

interface SpreadsheetAutoFilterConditionSurface
  extends SpreadsheetAutoFilterConditionDialogSource {
  column: number;
  filterRange: {
    column: [number, number];
    row: [number, number];
  };
  invoker: HTMLElement | null;
  sheetId: string;
}

export function useSpreadsheetAutoFilter({
  canvasRef,
  commandsRef,
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
  const dialogSelectionRef = useRef<SpreadsheetAutoFilterSelectionState | null>(
    null,
  );
  const applyingConditionRef = useRef(false);
  const restoreInvokerFocusRef = useRef(false);
  const menuWasOpenRef = useRef(false);
  const [conditionSurface, setConditionSurface] =
    useState<SpreadsheetAutoFilterConditionSurface | null>(null);
  const conditionSurfaceRef = useRef(conditionSurface);
  conditionSurfaceRef.current = conditionSurface;
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

  const openConditionDialog = useCallback((action: HTMLElement): boolean => {
    if (!editableRef.current || conditionSurfaceRef.current) return false;
    const activeSheet = contentRef.current.sheets.find(
      (candidate) => candidate.id === sheetIdRef.current,
    );
    const invoker = keyboardInvokerRef.current;
    const column = Number(invoker?.dataset.filterColumn);
    const range = normalizedSpreadsheetAutoFilterRange(
      activeSheet?.filter_select,
    );
    if (
      !activeSheet ||
      !range ||
      !Number.isSafeInteger(column) ||
      column < range.column[0] ||
      column > range.column[1]
    ) {
      return false;
    }
    const key = String(column - range.column[0]);
    const source: SpreadsheetAutoFilterConditionSurface = {
      column,
      columnLabel: spreadsheetAutoFilterColumnLabel(
        activeSheet,
        range.row[0],
        column,
      ),
      criteria: spreadsheetAutoFilterCriteria(activeSheet, column),
      filterRange: range,
      hasActiveFilter: Object.hasOwn(activeSheet.filter ?? {}, key),
      invoker,
      numeric: spreadsheetAutoFilterColumnIsNumeric(activeSheet, range, column),
      sheetId: activeSheet.id ?? sheetIdRef.current,
      sheetName: activeSheet.name,
    };
    dialogSelectionRef.current = {
      selections: [
        currentSpreadsheetAutoFilterSelection(
          selectionRef.current,
          workbookRef.current,
        ),
      ],
      sheetId: source.sheetId,
    };
    const menu = action.closest<HTMLElement>('.fortune-filter-menu');
    restoreInvokerFocusRef.current = false;
    menu?.querySelector<HTMLElement>('.button-default')?.click();
    conditionSurfaceRef.current = source;
    setConditionSurface(source);
    return true;
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
    const handleClick = (event: MouseEvent) => {
      const action = spreadsheetAutoFilterConditionAction(event.target);
      if (!action) return;
      event.preventDefault();
      event.stopPropagation();
      openConditionDialog(action);
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
    container.addEventListener('click', handleClick, true);
    container.addEventListener('pointerdown', handlePointerDown, true);
    container.addEventListener('keydown', handleKeyDown, true);
    if (typeof MutationObserver === 'undefined') {
      return () => {
        container.removeEventListener('click', handleClick, true);
        container.removeEventListener('pointerdown', handlePointerDown, true);
        container.removeEventListener('keydown', handleKeyDown, true);
      };
    }
    const observer = new MutationObserver(enhance);
    observer.observe(container, { childList: true, subtree: true });
    return () => {
      observer.disconnect();
      container.removeEventListener('click', handleClick, true);
      container.removeEventListener('pointerdown', handlePointerDown, true);
      container.removeEventListener('keydown', handleKeyDown, true);
    };
  }, [canvasRef, editable, mountRevision, openConditionDialog]);

  useEffect(() => {
    if (editable) return;
    conditionSurfaceRef.current = null;
    dialogSelectionRef.current = null;
    applyingConditionRef.current = false;
    setConditionSurface(null);
  }, [editable]);

  const closeConditionDialog = useCallback(() => {
    conditionSurfaceRef.current = null;
    dialogSelectionRef.current = null;
    setConditionSurface(null);
  }, []);

  const applyCondition = useCallback(
    (criteria: WorkSpreadsheetFilterCriteria): boolean => {
      const source = conditionSurfaceRef.current;
      if (!source || !editableRef.current) return false;
      applyingConditionRef.current = true;
      let handled = false;
      try {
        handled =
          commandsRef.current?.applyAutoFilterCriteria({
            sheetId: source.sheetId,
            column: source.column,
            filterRange: source.filterRange,
            criteria,
          }) ?? false;
      } finally {
        applyingConditionRef.current = false;
      }
      if (!handled) {
        showToast('筛选区域已发生变化，请重新打开筛选菜单。', 'error');
      }
      return handled;
    },
    [commandsRef],
  );

  const clearCondition = useCallback((): boolean => {
    const source = conditionSurfaceRef.current;
    if (!source || !editableRef.current) return false;
    applyingConditionRef.current = true;
    let handled = false;
    try {
      handled =
        commandsRef.current?.clearAutoFilterCriteria({
          sheetId: source.sheetId,
          column: source.column,
          filterRange: source.filterRange,
        }) ?? false;
    } finally {
      applyingConditionRef.current = false;
    }
    if (!handled) showToast('无法清除此列筛选。', 'error');
    return handled;
  }, [commandsRef]);

  const selectionForChange = useCallback(
    () => (applyingConditionRef.current ? dialogSelectionRef.current : null),
    [],
  );

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
    dialog: conditionSurface
      ? createElement(SpreadsheetAutoFilterConditionDialog, {
          source: conditionSurface,
          restoreFocusTarget: () =>
            conditionSurface.invoker?.isConnected
              ? conditionSurface.invoker
              : canvasRef.current,
          onApply: applyCondition,
          onClear: clearCondition,
          onClose: closeConditionDialog,
        })
      : null,
    reserveAltKey,
    selectionForChange,
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

function normalizedSpreadsheetAutoFilterRange(
  range: WorkSpreadsheetSheet['filter_select'] | undefined,
): SpreadsheetAutoFilterConditionSurface['filterRange'] | null {
  const rowStart = range?.row?.[0];
  const rowEnd = range?.row?.[1];
  const columnStart = range?.column?.[0];
  const columnEnd = range?.column?.[1];
  if (
    !Number.isSafeInteger(rowStart) ||
    !Number.isSafeInteger(rowEnd) ||
    !Number.isSafeInteger(columnStart) ||
    !Number.isSafeInteger(columnEnd)
  ) {
    return null;
  }
  return {
    row: [
      Math.min(Number(rowStart), Number(rowEnd)),
      Math.max(Number(rowStart), Number(rowEnd)),
    ],
    column: [
      Math.min(Number(columnStart), Number(columnEnd)),
      Math.max(Number(columnStart), Number(columnEnd)),
    ],
  };
}

function spreadsheetAutoFilterColumnIsNumeric(
  sheet: WorkSpreadsheetSheet,
  range: SpreadsheetAutoFilterConditionSurface['filterRange'],
  column: number,
): boolean {
  const cellAt = spreadsheetSheetCellReader(sheet);
  let values = 0;
  for (let row = range.row[0] + 1; row <= range.row[1]; row += 1) {
    const value = cellAt(row, column)?.v ?? cellAt(row, column)?.m;
    if (value === undefined || value === null || value === '') continue;
    values += 1;
    if (typeof value !== 'number' || !Number.isFinite(value)) return false;
  }
  return values > 0;
}
