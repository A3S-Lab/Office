import type { Cell } from '@fortune-sheet/core';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { showToast } from '../../../state/app-state';
import type { WorkSpreadsheetContent } from '../work-types';
import type {
  SpreadsheetEditorCommands,
  SpreadsheetSortCommandPort,
} from './spreadsheet-command-controller';
import {
  createSpreadsheetSortDialogSource,
  type SpreadsheetSortDialogSource,
  type SpreadsheetSortDialogValue,
  type SpreadsheetSortOpenRequest,
} from './spreadsheet-sort';
import { SpreadsheetSortDialog } from './spreadsheet-sort-dialog';

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
  getRows: (request: SpreadsheetSortOpenRequest) => (Cell | null)[][] | null;
  preview: boolean;
}) {
  const [source, setSource] = useState<SpreadsheetSortDialogSource | null>(
    null,
  );
  const invokerRef = useRef<HTMLElement | null>(null);

  useEffect(() => {
    if (!preview) return;
    setSource(null);
    invokerRef.current = null;
  }, [preview]);

  const open = useCallback(
    (request: SpreadsheetSortOpenRequest): boolean => {
      if (preview || source) return false;
      const sheet = contentRef.current.sheets.find(
        (candidate) => candidate.id === request.sheetId,
      );
      const rows = getRows(request);
      if (!sheet || !rows) return false;
      const nextSource = createSpreadsheetSortDialogSource(
        request.sheetId,
        sheet.name,
        request,
        rows,
      );
      if (!nextSource) return false;
      invokerRef.current =
        document.activeElement instanceof HTMLElement
          ? document.activeElement
          : getGridFocusTarget();
      setSource(nextSource);
      return true;
    },
    [contentRef, getGridFocusTarget, getRows, preview, source],
  );

  const close = useCallback(() => {
    const invoker = invokerRef.current;
    const grid = getGridFocusTarget();
    const openedFromGrid = invoker === grid;
    setSource(null);
    requestAnimationFrame(() => {
      if (openedFromGrid || !invoker?.isConnected) focusGrid(invoker);
      else invoker.focus({ preventScroll: true });
      if (invokerRef.current === invoker) invokerRef.current = null;
    });
  }, [focusGrid, getGridFocusTarget]);

  const commandPort = useMemo<SpreadsheetSortCommandPort>(
    () => ({ canOpen: !preview && source === null, open }),
    [open, preview, source],
  );

  const apply = useCallback(
    (value: SpreadsheetSortDialogValue): boolean => {
      if (!source) return false;
      const handled =
        commandsRef.current?.applyCustomSort({
          sheetId: source.sheetId,
          range: source.range,
          hasHeader: value.hasHeader,
          keys: value.keys,
        }) ?? false;
      if (!handled) showToast('无法应用当前排序设置。', 'error');
      return handled;
    },
    [commandsRef, source],
  );

  return {
    commandPort,
    dialog: source ? (
      <SpreadsheetSortDialog
        source={source}
        restoreFocusTarget={() =>
          invokerRef.current?.isConnected
            ? invokerRef.current
            : getGridFocusTarget()
        }
        onApply={apply}
        onClose={close}
      />
    ) : null,
  };
}
