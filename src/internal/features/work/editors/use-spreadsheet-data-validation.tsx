import type { Selection } from '@fortune-sheet/core';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { showToast } from '../../../state/app-state';
import type { WorkSpreadsheetContent } from '../work-types';
import type {
  SpreadsheetDataValidationCommandPort,
  SpreadsheetEditorCommands,
} from './spreadsheet-command-controller';
import { SpreadsheetDataValidationDialog } from './spreadsheet-data-validation-dialog';
import {
  createSpreadsheetDataValidationDialogSource,
  type SpreadsheetDataValidationDialogSource,
  type SpreadsheetDataValidationDialogValue,
  type SpreadsheetDataValidationTarget,
  spreadsheetDataValidationFailureMessage,
} from './spreadsheet-data-validation';
import { finiteSpreadsheetSelection } from './spreadsheet-editor-support';

export interface SpreadsheetDataValidationSelectionState {
  selections: Selection[];
  sheetId: string;
}

export function useSpreadsheetDataValidation({
  commandsRef,
  contentRef,
  focusGrid,
  getGridFocusTarget,
  getLiveSelections,
  preview,
}: {
  commandsRef: { current: SpreadsheetEditorCommands | null };
  contentRef: { current: WorkSpreadsheetContent };
  focusGrid: (focusOrigin: Element | null) => void;
  getGridFocusTarget: () => HTMLElement | null;
  getLiveSelections: () => Selection[] | undefined;
  preview: boolean;
}) {
  const [source, setSource] =
    useState<SpreadsheetDataValidationDialogSource | null>(null);
  const invokerRef = useRef<HTMLElement | null>(null);
  const selectionRef = useRef<SpreadsheetDataValidationSelectionState | null>(
    null,
  );
  const applyingRef = useRef(false);

  useEffect(() => {
    if (!preview) return;
    setSource(null);
    invokerRef.current = null;
    selectionRef.current = null;
    applyingRef.current = false;
  }, [preview]);

  const open = useCallback(
    (target: SpreadsheetDataValidationTarget): boolean => {
      if (preview || source) return false;
      const nextSource = createSpreadsheetDataValidationDialogSource(
        contentRef.current,
        target,
      );
      if (!nextSource) return false;
      const liveSelections = getLiveSelections();
      const selections = liveSelections?.length
        ? liveSelections.map(finiteSpreadsheetSelection)
        : target.ranges.map((range) => ({
            row: [...range.row],
            column: [...range.column],
            row_focus: range.row[0],
            column_focus: range.column[0],
          }));
      invokerRef.current =
        document.activeElement instanceof HTMLElement
          ? document.activeElement
          : getGridFocusTarget();
      selectionRef.current = { sheetId: target.sheetId, selections };
      setSource(nextSource);
      return true;
    },
    [contentRef, getGridFocusTarget, getLiveSelections, preview, source],
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
      selectionRef.current = null;
    });
  }, [focusGrid, getGridFocusTarget]);

  const commandPort = useMemo<SpreadsheetDataValidationCommandPort>(
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
      <SpreadsheetDataValidationDialog
        source={source}
        restoreFocusTarget={() =>
          invokerRef.current?.isConnected
            ? invokerRef.current
            : getGridFocusTarget()
        }
        onApply={(value) => {
          applyingRef.current = true;
          let handled = false;
          try {
            handled =
              commandsRef.current?.applyDataValidation({
                sheetId: source.sheetId,
                ranges: source.ranges,
                activeCell: source.activeCell,
                value,
              }) ?? false;
          } finally {
            applyingRef.current = false;
          }
          if (!handled) {
            showToast(
              spreadsheetDataValidationError(
                contentRef.current,
                source,
                value,
                '无法应用数据验证。',
              ) ?? '无法应用数据验证。',
              'error',
            );
          }
          return handled;
        }}
        onClose={close}
        onRemove={() => {
          applyingRef.current = true;
          let handled = false;
          try {
            handled =
              commandsRef.current?.removeDataValidation({
                sheetId: source.sheetId,
                ranges: source.ranges,
                activeCell: source.activeCell,
              }) ?? false;
          } finally {
            applyingRef.current = false;
          }
          if (!handled) showToast('无法清除数据验证。', 'error');
          return handled;
        }}
        onValidate={(value) =>
          spreadsheetDataValidationError(
            contentRef.current,
            source,
            value,
            null,
          )
        }
      />
    ) : null,
  };
}

function spreadsheetDataValidationError(
  content: WorkSpreadsheetContent,
  source: SpreadsheetDataValidationDialogSource,
  value: SpreadsheetDataValidationDialogValue,
  fallback: string | null,
): string | null {
  return (
    spreadsheetDataValidationFailureMessage(content, {
      sheetId: source.sheetId,
      ranges: source.ranges,
      activeCell: source.activeCell,
      value,
    }) ?? fallback
  );
}
