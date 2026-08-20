import type { Selection } from '@fortune-sheet/core';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { showToast } from '../../../state/app-state';
import type { WorkSpreadsheetContent } from '../work-types';
import type {
  SpreadsheetEditorCommands,
  SpreadsheetHyperlinkCommandPort,
} from './spreadsheet-command-controller';
import { finiteSpreadsheetSelection } from './spreadsheet-editor-support';
import {
  createSpreadsheetHyperlinkDialogSource,
  type SpreadsheetHyperlinkCell,
  type SpreadsheetHyperlinkDialogSource,
  type SpreadsheetHyperlinkDialogValue,
  validateSpreadsheetHyperlinkRequest,
} from './spreadsheet-hyperlink';
import { SpreadsheetHyperlinkDialog } from './spreadsheet-hyperlink-dialog';

export interface SpreadsheetHyperlinkSelectionState {
  sheetId: string;
  selection: Selection;
}

export function useSpreadsheetHyperlink({
  commandsRef,
  contentRef,
  focusGrid,
  getGridFocusTarget,
  getLiveSelection,
  preview,
  selectionState,
}: {
  commandsRef: { current: SpreadsheetEditorCommands | null };
  contentRef: { current: WorkSpreadsheetContent };
  focusGrid: (focusOrigin: Element | null) => void;
  getGridFocusTarget: () => HTMLElement | null;
  getLiveSelection: () => Selection | undefined;
  preview: boolean;
  selectionState: SpreadsheetHyperlinkSelectionState | null;
}) {
  const [source, setSource] = useState<SpreadsheetHyperlinkDialogSource | null>(
    null,
  );
  const invokerRef = useRef<HTMLElement | null>(null);
  const selectionRef = useRef<SpreadsheetHyperlinkSelectionState | null>(null);
  const applyingRef = useRef(false);

  useEffect(() => {
    if (!preview) return;
    setSource(null);
    invokerRef.current = null;
    selectionRef.current = null;
    applyingRef.current = false;
  }, [preview]);

  const open = useCallback(
    (request: SpreadsheetHyperlinkCell): boolean => {
      if (preview || source) return false;
      const nextSource = createSpreadsheetHyperlinkDialogSource(
        contentRef.current,
        request,
      );
      if (!nextSource) return false;
      const sheet = contentRef.current.sheets.find(
        (candidate) => candidate.id === request.sheetId,
      );
      const savedSelection =
        selectionState?.sheetId === request.sheetId
          ? selectionState.selection
          : sheet?.luckysheet_select_save?.at(-1);
      const selection = getLiveSelection() ??
        savedSelection ?? {
          row: [request.row, request.row],
          column: [request.column, request.column],
          row_focus: request.row,
          column_focus: request.column,
        };
      invokerRef.current =
        document.activeElement instanceof HTMLElement
          ? document.activeElement
          : getGridFocusTarget();
      selectionRef.current = {
        sheetId: request.sheetId,
        selection: finiteSpreadsheetSelection(selection),
      };
      setSource(nextSource);
      return true;
    },
    [
      contentRef,
      getGridFocusTarget,
      getLiveSelection,
      preview,
      selectionState,
      source,
    ],
  );

  const close = useCallback(() => {
    const invoker = invokerRef.current;
    const grid = getGridFocusTarget();
    const openedFromGrid = invoker === grid;
    setSource(null);
    requestAnimationFrame(() => {
      if (openedFromGrid || !invoker?.isConnected) {
        focusGrid(invoker);
      } else {
        invoker.focus({ preventScroll: true });
      }
      if (invokerRef.current === invoker) invokerRef.current = null;
      selectionRef.current = null;
    });
  }, [focusGrid, getGridFocusTarget]);

  const commandPort = useMemo<SpreadsheetHyperlinkCommandPort>(
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
      <SpreadsheetHyperlinkDialog
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
              commandsRef.current?.applyHyperlink({
                sheetId: source.sheetId,
                row: source.row,
                column: source.column,
                ...value,
              }) ?? false;
          } finally {
            applyingRef.current = false;
          }
          if (!handled) {
            showToast(
              spreadsheetHyperlinkFailureMessage(
                contentRef.current,
                source,
                value,
                '无法应用超链接。',
              ),
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
              commandsRef.current?.removeHyperlink({
                sheetId: source.sheetId,
                row: source.row,
                column: source.column,
              }) ?? false;
          } finally {
            applyingRef.current = false;
          }
          if (!handled) showToast('无法移除超链接。', 'error');
          return handled;
        }}
        onValidate={(value) =>
          spreadsheetHyperlinkValidationMessage(
            contentRef.current,
            source,
            value,
          )
        }
      />
    ) : null,
  };
}

function spreadsheetHyperlinkValidationMessage(
  content: WorkSpreadsheetContent,
  source: SpreadsheetHyperlinkDialogSource,
  value: SpreadsheetHyperlinkDialogValue,
): string | null {
  const validation = validateSpreadsheetHyperlinkRequest(content, {
    sheetId: source.sheetId,
    row: source.row,
    column: source.column,
    ...value,
  });
  return validation.ok ? null : validation.message;
}

function spreadsheetHyperlinkFailureMessage(
  content: WorkSpreadsheetContent,
  source: SpreadsheetHyperlinkDialogSource,
  value: SpreadsheetHyperlinkDialogValue,
  fallback: string,
): string {
  return (
    spreadsheetHyperlinkValidationMessage(content, source, value) ?? fallback
  );
}
