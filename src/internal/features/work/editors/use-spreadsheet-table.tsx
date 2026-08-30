import type { Selection } from '@fortune-sheet/core';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { showToast } from '../../../state/app-state';
import type { WorkSpreadsheetContent } from '../work-types';
import type {
  SpreadsheetEditorCommands,
  SpreadsheetTableCommandPort,
} from './spreadsheet-command-controller';
import { finiteSpreadsheetSelection } from './spreadsheet-editor-support';
import {
  createSpreadsheetTableDialogSource,
  type SpreadsheetTableDialogSource,
  type SpreadsheetTableDialogValue,
  type SpreadsheetTableTarget,
  spreadsheetTableFailureMessage,
  spreadsheetTableRangeFromText,
} from './spreadsheet-table';
import { SpreadsheetTableDialog } from './spreadsheet-table-dialog';

export interface SpreadsheetTableSelectionState {
  selection: Selection;
  sheetId: string;
}

export function useSpreadsheetTable({
  commandsRef,
  contentRef,
  focusGrid,
  getGridFocusTarget,
  getLiveSelection,
  preview,
}: {
  commandsRef: { current: SpreadsheetEditorCommands | null };
  contentRef: { current: WorkSpreadsheetContent };
  focusGrid: (focusOrigin: Element | null) => void;
  getGridFocusTarget: () => HTMLElement | null;
  getLiveSelection: () => Selection | undefined;
  preview: boolean;
}) {
  const [source, setSource] = useState<SpreadsheetTableDialogSource | null>(
    null,
  );
  const invokerRef = useRef<HTMLElement | null>(null);
  const selectionRef = useRef<SpreadsheetTableSelectionState | null>(null);
  const applyingRef = useRef(false);
  const appliedRef = useRef(false);

  useEffect(() => {
    if (!preview) return;
    setSource(null);
    invokerRef.current = null;
    selectionRef.current = null;
    applyingRef.current = false;
    appliedRef.current = false;
  }, [preview]);

  const open = useCallback(
    (target: SpreadsheetTableTarget): boolean => {
      if (preview || source) return false;
      const nextSource = createSpreadsheetTableDialogSource(
        contentRef.current,
        target,
      );
      if (!nextSource) {
        showToast('请选择包含标题和至少一行数据的连续、未合并区域。', 'error');
        return false;
      }
      const grid = getGridFocusTarget();
      invokerRef.current =
        document.activeElement instanceof HTMLElement
          ? document.activeElement
          : grid;
      selectionRef.current = {
        sheetId: target.sheetId,
        selection: finiteSpreadsheetSelection(
          getLiveSelection() ?? target.selection,
        ),
      };
      appliedRef.current = false;
      setSource(nextSource);
      return true;
    },
    [contentRef, getGridFocusTarget, getLiveSelection, preview, source],
  );

  const close = useCallback(() => {
    const invoker = invokerRef.current;
    const grid = getGridFocusTarget();
    const restoreGrid = appliedRef.current || invoker === grid;
    setSource(null);
    requestAnimationFrame(() => {
      if (restoreGrid || !invoker?.isConnected) focusGrid(invoker);
      else invoker.focus({ preventScroll: true });
      if (invokerRef.current === invoker) invokerRef.current = null;
      selectionRef.current = null;
      appliedRef.current = false;
    });
  }, [focusGrid, getGridFocusTarget]);

  const commandPort = useMemo<SpreadsheetTableCommandPort>(
    () => ({ canOpen: !preview && source === null, open }),
    [open, preview, source],
  );
  const selectionForChange = useCallback(
    () => (applyingRef.current ? selectionRef.current : null),
    [],
  );

  return {
    commandPort,
    selectionForChange,
    dialog: source ? (
      <SpreadsheetTableDialog
        source={source}
        restoreFocusTarget={() =>
          appliedRef.current
            ? getGridFocusTarget()
            : invokerRef.current?.isConnected
              ? invokerRef.current
              : getGridFocusTarget()
        }
        onApply={(value) => {
          const range = spreadsheetTableRangeFromText(value.rangeReference);
          if (!range) return false;
          applyingRef.current = true;
          let handled = false;
          try {
            handled =
              commandsRef.current?.applyTable({
                headerRow: value.headerRow,
                name: source.name,
                range,
                sheetId: source.sheetId,
                totalsRow: value.totalsRow === true,
              }) ?? false;
            appliedRef.current = handled;
          } finally {
            applyingRef.current = false;
          }
          if (!handled) {
            showToast(
              spreadsheetTableValidationMessage(
                contentRef.current,
                source,
                value,
              ) ?? '无法创建表格。',
              'error',
            );
          }
          return handled;
        }}
        onClose={close}
        onValidate={(value) =>
          spreadsheetTableValidationMessage(contentRef.current, source, value)
        }
      />
    ) : null,
  };
}

function spreadsheetTableValidationMessage(
  content: WorkSpreadsheetContent,
  source: SpreadsheetTableDialogSource,
  value: SpreadsheetTableDialogValue,
): string | null {
  const range = spreadsheetTableRangeFromText(value.rangeReference);
  if (!range) return '请输入一个连续区域。';
  return spreadsheetTableFailureMessage(content, {
    headerRow: value.headerRow,
    name: source.name,
    range,
    sheetId: source.sheetId,
    totalsRow: value.totalsRow === true,
  });
}
