import type { Cell } from '@fortune-sheet/core';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { showToast } from '../../../state/app-state';
import type { WorkSpreadsheetContent } from '../work-types';
import type { SpreadsheetCellRange } from './spreadsheet-cell-range';
import { spreadsheetCellRangesEqual } from './spreadsheet-cell-range';
import type {
  SpreadsheetEditorCommands,
  SpreadsheetSortCommandPort,
} from './spreadsheet-command-controller';
import {
  createSpreadsheetSortDialogSource,
  createSpreadsheetSortRangeDialogSource,
  type SpreadsheetSortDialogSource,
  type SpreadsheetSortDialogValue,
  type SpreadsheetSortOpenRequest,
  type SpreadsheetSortRangeCandidate,
  type SpreadsheetSortRangeChoice,
  type SpreadsheetSortRangeDialogSource,
  type SpreadsheetSortRequest,
} from './spreadsheet-sort';
import { SpreadsheetSortDialog } from './spreadsheet-sort-dialog';
import {
  createSpreadsheetSortCustomList,
  MAX_SPREADSHEET_SORT_SESSION_CUSTOM_LISTS,
  spreadsheetSortCustomListsEqual,
  type SpreadsheetSortCustomList,
} from './spreadsheet-sort-custom-list';
import { SpreadsheetSortRangeDialog } from './spreadsheet-sort-range-dialog';

type SpreadsheetSortSurface =
  | {
      kind: 'custom';
      selectedRange: SpreadsheetCellRange;
      source: SpreadsheetSortDialogSource;
    }
  | {
      kind: 'range';
      request: SpreadsheetSortOpenRequest;
      source: SpreadsheetSortRangeDialogSource;
    };

interface SpreadsheetSortReadRequest {
  range: SpreadsheetSortRangeCandidate['range'];
  sheetId: string;
}

export function useSpreadsheetSort({
  commandsRef,
  contentRef,
  focusGrid,
  getGridFocusTarget,
  getRows,
  preview,
}: {
  commandsRef: { current: SpreadsheetEditorCommands | null };
  contentRef: { current: WorkSpreadsheetContent };
  focusGrid: (focusOrigin: Element | null) => void;
  getGridFocusTarget: () => HTMLElement | null;
  getRows: (request: SpreadsheetSortReadRequest) => (Cell | null)[][] | null;
  preview: boolean;
}) {
  const [surface, setSurface] = useState<SpreadsheetSortSurface | null>(null);
  const [customLists, setCustomLists] = useState<
    readonly SpreadsheetSortCustomList[]
  >([]);
  const authorizedRequestRef = useRef<{
    request: SpreadsheetSortRequest;
    selectedRange: SpreadsheetCellRange;
  } | null>(null);
  const invokerRef = useRef<HTMLElement | null>(null);

  useEffect(() => {
    if (!preview) return;
    setSurface(null);
    authorizedRequestRef.current = null;
    invokerRef.current = null;
  }, [preview]);

  const sourceForCandidate = useCallback(
    (
      request: SpreadsheetSortOpenRequest,
      candidate: SpreadsheetSortRangeCandidate,
    ): SpreadsheetSortDialogSource | null => {
      if (!candidate.available) return null;
      const sheet = contentRef.current.sheets.find(
        (item) => item.id === request.sheetId,
      );
      const rows = getRows({
        sheetId: request.sheetId,
        range: candidate.range,
      });
      if (!sheet || !rows) return null;
      return createSpreadsheetSortDialogSource(
        request.sheetId,
        sheet.name,
        { range: candidate.range, activeColumn: request.activeColumn },
        rows,
        customLists,
      );
    },
    [contentRef, customLists, getRows],
  );

  const rememberCustomList = useCallback(
    (candidate: SpreadsheetSortCustomList) => {
      if (candidate.source !== 'session') return;
      const list = createSpreadsheetSortCustomList(
        candidate.entries,
        'session',
      );
      if (!list) return;
      setCustomLists((current) => {
        if (
          current.some((item) =>
            spreadsheetSortCustomListsEqual(item.entries, list.entries),
          ) ||
          current.length >= MAX_SPREADSHEET_SORT_SESSION_CUSTOM_LISTS
        ) {
          return current;
        }
        return Object.freeze([...current, list]);
      });
    },
    [],
  );

  const applyAuthorizedRequest = useCallback(
    (
      request: SpreadsheetSortRequest,
      selectedRange: SpreadsheetCellRange,
    ): boolean => {
      authorizedRequestRef.current = { request, selectedRange };
      try {
        const handled = commandsRef.current?.applyCustomSort(request) ?? false;
        if (!handled) showToast('无法应用当前排序设置。', 'error');
        return handled;
      } finally {
        authorizedRequestRef.current = null;
      }
    },
    [commandsRef],
  );

  const applyQuickCandidate = useCallback(
    (
      request: SpreadsheetSortOpenRequest,
      candidate: SpreadsheetSortRangeCandidate,
    ): boolean => {
      if (request.intent.type !== 'quick') return false;
      const source = sourceForCandidate(request, candidate);
      if (!source) return false;
      return applyAuthorizedRequest(
        {
          sheetId: source.sheetId,
          range: source.range,
          hasHeader: source.value.hasHeader,
          keys: [
            {
              column: request.activeColumn,
              direction: request.intent.direction,
            },
          ],
        },
        request.selected.range,
      );
    },
    [applyAuthorizedRequest, sourceForCandidate],
  );

  const open = useCallback(
    (request: SpreadsheetSortOpenRequest): boolean => {
      if (preview || surface) return false;
      const sheet = contentRef.current.sheets.find(
        (candidate) => candidate.id === request.sheetId,
      );
      if (!sheet) return false;
      invokerRef.current =
        document.activeElement instanceof HTMLElement
          ? document.activeElement
          : getGridFocusTarget();
      if (request.expanded) {
        const source = createSpreadsheetSortRangeDialogSource(
          sheet.name,
          request,
        );
        if (!source) return false;
        setSurface({ kind: 'range', request, source });
        return true;
      }
      const source = sourceForCandidate(request, request.selected);
      if (!source) return false;
      if (request.intent.type === 'custom') {
        setSurface({
          kind: 'custom',
          selectedRange: request.selected.range,
          source,
        });
        return true;
      }
      const handled = applyQuickCandidate(request, request.selected);
      if (handled) {
        const invoker = invokerRef.current;
        requestAnimationFrame(() => focusGrid(invoker));
        invokerRef.current = null;
      }
      return handled;
    },
    [
      applyQuickCandidate,
      contentRef,
      focusGrid,
      getGridFocusTarget,
      preview,
      sourceForCandidate,
      surface,
    ],
  );

  const close = useCallback(() => {
    const invoker = invokerRef.current;
    const grid = getGridFocusTarget();
    const openedFromGrid = invoker === grid;
    setSurface(null);
    requestAnimationFrame(() => {
      if (openedFromGrid || !invoker?.isConnected) focusGrid(invoker);
      if (invokerRef.current === invoker) invokerRef.current = null;
    });
  }, [focusGrid, getGridFocusTarget]);

  const canApply = useCallback(
    (
      request: SpreadsheetSortRequest,
      liveRange: { column: readonly number[]; row: readonly number[] },
    ): boolean => {
      const authorized = authorizedRequestRef.current;
      if (
        authorized?.request.sheetId === request.sheetId &&
        spreadsheetCellRangesEqual(authorized.request.range, request.range) &&
        spreadsheetCellRangesEqual(authorized.selectedRange, liveRange)
      ) {
        return true;
      }
      return Boolean(
        surface?.kind === 'custom' &&
          surface.source.sheetId === request.sheetId &&
          spreadsheetCellRangesEqual(surface.source.range, request.range) &&
          spreadsheetCellRangesEqual(surface.selectedRange, liveRange),
      );
    },
    [surface],
  );

  const commandPort = useMemo<SpreadsheetSortCommandPort>(
    () => ({ canApply, canOpen: !preview && surface === null, open }),
    [canApply, open, preview, surface],
  );

  const apply = useCallback(
    (value: SpreadsheetSortDialogValue): boolean => {
      if (surface?.kind !== 'custom') return false;
      return applyAuthorizedRequest(
        {
          sheetId: surface.source.sheetId,
          range: surface.source.range,
          hasHeader: value.hasHeader,
          keys: value.keys,
        },
        surface.selectedRange,
      );
    },
    [applyAuthorizedRequest, surface],
  );

  const chooseRange = useCallback(
    (choice: SpreadsheetSortRangeChoice): boolean => {
      if (surface?.kind !== 'range') return false;
      const candidate =
        choice === 'expand'
          ? surface.request.expanded
          : surface.request.selected;
      if (!candidate?.available) return false;
      if (surface.request.intent.type === 'quick') {
        return applyQuickCandidate(surface.request, candidate);
      }
      const source = sourceForCandidate(surface.request, candidate);
      if (!source) {
        showToast('无法读取当前排序区域。', 'error');
        return false;
      }
      setSurface({
        kind: 'custom',
        selectedRange: surface.request.selected.range,
        source,
      });
      return false;
    },
    [applyQuickCandidate, sourceForCandidate, surface],
  );

  const restoreFocusTarget = () =>
    invokerRef.current?.isConnected ? invokerRef.current : getGridFocusTarget();

  return {
    commandPort,
    dialog:
      surface?.kind === 'custom' ? (
        <SpreadsheetSortDialog
          source={surface.source}
          restoreFocusTarget={restoreFocusTarget}
          onApply={apply}
          onRememberCustomList={rememberCustomList}
          onClose={close}
        />
      ) : surface?.kind === 'range' ? (
        <SpreadsheetSortRangeDialog
          source={surface.source}
          restoreFocusTarget={restoreFocusTarget}
          onApply={chooseRange}
          onClose={close}
        />
      ) : null,
  };
}
