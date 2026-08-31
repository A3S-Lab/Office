import type { MutableRefObject } from 'react';
import { useCallback, useEffect, useMemo, useState } from 'react';
import { showToast } from '../../../state/app-state';
import type { WorkSpreadsheetContent } from '../work-types';
import {
  browserSpreadsheetClipboard,
  clearRichSpreadsheetClipboard,
  copySpreadsheetSelection,
  spreadsheetClipboardSnapshotForText,
  storeRichSpreadsheetClipboard,
  type SpreadsheetClipboardPort,
} from './spreadsheet-clipboard';
import type { SpreadsheetClipboardCommandPort } from './spreadsheet-command-controller';
import type { SpreadsheetCellRange } from './spreadsheet-cell-range';
import {
  applySpreadsheetPasteSpecial,
  captureSpreadsheetClipboardSnapshot,
  spreadsheetPasteSpecialValidationError,
  type SpreadsheetClipboardSnapshot,
  type SpreadsheetPasteContent,
  type SpreadsheetPasteSpecialOptions,
} from './spreadsheet-paste-special';

export interface SpreadsheetClipboardSelectionSource {
  sheetId: string;
  range: SpreadsheetCellRange;
  plainText: string;
}

export interface SpreadsheetPasteSpecialDialogSource {
  snapshot: SpreadsheetClipboardSnapshot;
  targetSheetId: string;
  targetRange: SpreadsheetCellRange;
  invoker: HTMLElement | null;
}

export function useSpreadsheetClipboard({
  clipboard = browserSpreadsheetClipboard,
  canAccessSelection,
  clearSelection,
  commit,
  contentRef,
  editable,
  fallbackFocusTarget,
  getSelection,
}: {
  clipboard?: SpreadsheetClipboardPort;
  canAccessSelection: boolean;
  clearSelection: () => boolean;
  commit: (content: WorkSpreadsheetContent) => boolean;
  contentRef: MutableRefObject<WorkSpreadsheetContent>;
  editable: boolean;
  fallbackFocusTarget: () => HTMLElement | null;
  getSelection: () => SpreadsheetClipboardSelectionSource | null;
}) {
  const [dialogSource, setDialogSource] =
    useState<SpreadsheetPasteSpecialDialogSource | null>(null);
  useEffect(() => {
    if (!editable) setDialogSource(null);
  }, [editable]);

  const currentSource = useCallback(() => {
    if (!editable) return null;
    const selection = getSelection();
    if (!selection) return null;
    const snapshot = captureSpreadsheetClipboardSnapshot(
      contentRef.current,
      selection.sheetId,
      selection.range,
      selection.plainText,
    );
    return { selection, snapshot };
  }, [contentRef, editable, getSelection]);

  const copy = useCallback(
    (cut: boolean): boolean => {
      const source = currentSource();
      if (!source) return false;
      if (source.snapshot) storeRichSpreadsheetClipboard(source.snapshot);
      else clearRichSpreadsheetClipboard();
      void copySpreadsheetSelection(clipboard, source.selection.plainText, cut);
      if (cut && !clearSelection()) {
        showToast('选区已复制，但无法清除原内容。', 'error');
        return false;
      }
      return true;
    },
    [clearSelection, clipboard, currentSource],
  );

  const applySnapshot = useCallback(
    (
      snapshot: SpreadsheetClipboardSnapshot,
      targetSheetId: string,
      targetRange: SpreadsheetCellRange,
      options: SpreadsheetPasteSpecialOptions,
      notify = true,
    ): boolean => {
      const request = {
        snapshot,
        targetSheetId,
        targetSelection: targetRange,
        options,
      };
      const error = spreadsheetPasteSpecialValidationError(
        contentRef.current,
        request,
      );
      if (error) {
        if (notify) showToast(error, 'error');
        return false;
      }
      const result = applySpreadsheetPasteSpecial(contentRef.current, request);
      if (!result || !commit(result.content)) {
        if (notify) showToast('当前选区无法粘贴这些内容。', 'error');
        return false;
      }
      contentRef.current = result.content;
      if (notify) showToast('已粘贴到当前选区', 'success');
      return true;
    },
    [commit, contentRef],
  );

  const pasteText = useCallback(
    (
      plainText: string,
      content: SpreadsheetPasteContent = 'all',
      notify = true,
    ): boolean => {
      if (!editable) return false;
      const target = getSelection();
      const snapshot = spreadsheetClipboardSnapshotForText(plainText);
      if (!target || !snapshot) {
        if (notify) showToast('剪贴板中没有可粘贴的表格内容。', 'error');
        return false;
      }
      return applySnapshot(
        snapshot,
        target.sheetId,
        target.range,
        defaultSpreadsheetPasteSpecialOptions(content),
        notify,
      );
    },
    [applySnapshot, editable, getSelection],
  );

  const readAndPaste = useCallback(
    (content: SpreadsheetPasteContent): boolean => {
      if (!editable) return false;
      const target = getSelection();
      if (!target) return false;
      void clipboard
        .readText()
        .then((plainText) => {
          const snapshot = spreadsheetClipboardSnapshotForText(plainText);
          if (!snapshot) {
            showToast('剪贴板中没有可粘贴的表格内容。', 'error');
            return;
          }
          applySnapshot(
            snapshot,
            target.sheetId,
            target.range,
            defaultSpreadsheetPasteSpecialOptions(content),
          );
        })
        .catch(() =>
          showToast('无法读取剪贴板，请使用系统粘贴快捷键。', 'error'),
        );
      return true;
    },
    [applySnapshot, clipboard, editable, getSelection],
  );

  const openPasteSpecial = useCallback((): boolean => {
    if (!editable) return false;
    const target = getSelection();
    if (!target) return false;
    const invoker =
      typeof document !== 'undefined' &&
      document.activeElement instanceof HTMLElement
        ? document.activeElement
        : fallbackFocusTarget();
    void clipboard
      .readText()
      .then((plainText) => {
        const snapshot = spreadsheetClipboardSnapshotForText(plainText);
        if (!snapshot) {
          showToast('剪贴板中没有可粘贴的表格内容。', 'error');
          return;
        }
        setDialogSource({
          snapshot,
          targetSheetId: target.sheetId,
          targetRange: target.range,
          invoker,
        });
      })
      .catch(() =>
        showToast('无法读取剪贴板，请使用系统粘贴快捷键。', 'error'),
      );
    return true;
  }, [clipboard, editable, fallbackFocusTarget, getSelection]);

  const commandPort = useMemo<SpreadsheetClipboardCommandPort>(
    () => ({
      canCopySelection: editable && canAccessSelection,
      canCutSelection: editable && canAccessSelection,
      canOpenPasteSpecial: editable && canAccessSelection,
      canPasteSelection: editable && canAccessSelection,
      canPasteSpecial: () => editable && canAccessSelection,
      copySelection: () => copy(false),
      cutSelection: () => copy(true),
      openPasteSpecial,
      pasteSelection: () => readAndPaste('all'),
      pasteSpecial: readAndPaste,
    }),
    [canAccessSelection, copy, editable, openPasteSpecial, readAndPaste],
  );

  const validateDialog = useCallback(
    (options: SpreadsheetPasteSpecialOptions): string | null => {
      if (!dialogSource) return '选择性粘贴对话框已关闭。';
      return spreadsheetPasteSpecialValidationError(contentRef.current, {
        snapshot: dialogSource.snapshot,
        targetSheetId: dialogSource.targetSheetId,
        targetSelection: dialogSource.targetRange,
        options,
      });
    },
    [contentRef, dialogSource],
  );

  const applyDialog = useCallback(
    (options: SpreadsheetPasteSpecialOptions): boolean => {
      const applied = Boolean(
        dialogSource &&
          applySnapshot(
            dialogSource.snapshot,
            dialogSource.targetSheetId,
            dialogSource.targetRange,
            options,
          ),
      );
      if (!applied) return false;

      const restoreGridFocus = () =>
        fallbackFocusTarget()?.focus({ preventScroll: true });
      if (
        typeof window !== 'undefined' &&
        typeof window.requestAnimationFrame === 'function'
      ) {
        window.requestAnimationFrame(restoreGridFocus);
      } else {
        restoreGridFocus();
      }
      return true;
    },
    [applySnapshot, dialogSource, fallbackFocusTarget],
  );

  const closeDialog = useCallback(() => setDialogSource(null), []);
  const restoreDialogFocusTarget = useCallback(
    () =>
      dialogSource?.invoker?.isConnected
        ? dialogSource.invoker
        : fallbackFocusTarget(),
    [dialogSource, fallbackFocusTarget],
  );

  const copyToDataTransfer = useCallback(
    (data: DataTransfer, cut: boolean): boolean => {
      const source = currentSource();
      if (!source) return false;
      if (source.snapshot) storeRichSpreadsheetClipboard(source.snapshot);
      else clearRichSpreadsheetClipboard();
      data.setData('text/plain', source.selection.plainText);
      if (cut && !clearSelection()) return false;
      return true;
    },
    [clearSelection, currentSource],
  );

  return {
    applyDialog,
    closeDialog,
    commandPort,
    copyToDataTransfer,
    dialogSource,
    pasteText,
    restoreDialogFocusTarget,
    validateDialog,
  };
}

function defaultSpreadsheetPasteSpecialOptions(
  content: SpreadsheetPasteContent,
): SpreadsheetPasteSpecialOptions {
  return {
    content,
    operation: 'none',
    skipBlanks: false,
    transpose: false,
  };
}
