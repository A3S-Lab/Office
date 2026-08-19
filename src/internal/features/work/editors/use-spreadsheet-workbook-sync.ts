import type { Op } from '@fortune-sheet/core';
import { useCallback, useLayoutEffect, useRef, useState } from 'react';
import type { WorkSpreadsheetContent } from '../work-types';

const MAXIMUM_BUFFERED_SPREADSHEET_OPERATIONS = 10_001;

export interface SpreadsheetWorkbookSyncController {
  acceptContent: (content: WorkSpreadsheetContent) => void;
  ignoreChangeDuringExternalSync: (matchesExternalContent: boolean) => boolean;
  mountRevision: number;
  recordOperations: (operations: readonly Op[]) => void;
  takeOperations: () => Op[];
}

export type SpreadsheetExternalWorkbookSynchronizer = (
  previous: WorkSpreadsheetContent,
  next: WorkSpreadsheetContent,
) => boolean;

export function useSpreadsheetWorkbookSync(
  content: WorkSpreadsheetContent,
  synchronizeExternalWorkbook?: SpreadsheetExternalWorkbookSynchronizer,
): SpreadsheetWorkbookSyncController {
  const mountedContentRef = useRef(content);
  const externalSyncPendingRef = useRef(false);
  const pendingOperationsRef = useRef<Op[]>([]);
  const [mountRevision, setMountRevision] = useState(0);
  if (mountedContentRef.current !== content) {
    externalSyncPendingRef.current = true;
  }

  useLayoutEffect(() => {
    if (mountedContentRef.current === content) return;
    const previous = mountedContentRef.current;
    mountedContentRef.current = content;
    pendingOperationsRef.current = [];
    if (!synchronizeExternalWorkbook?.(previous, content)) {
      setMountRevision((revision) => revision + 1);
    }
  });

  const acceptContent = useCallback((nextContent: WorkSpreadsheetContent) => {
    mountedContentRef.current = nextContent;
  }, []);

  const ignoreChangeDuringExternalSync = useCallback(
    (matchesExternalContent: boolean): boolean => {
      if (!externalSyncPendingRef.current) return false;
      if (matchesExternalContent) externalSyncPendingRef.current = false;
      return true;
    },
    [],
  );

  const recordOperations = useCallback((operations: readonly Op[]) => {
    if (!operations.length) return;
    const pending = pendingOperationsRef.current;
    const available = MAXIMUM_BUFFERED_SPREADSHEET_OPERATIONS - pending.length;
    if (available <= 0) return;
    pending.push(...operations.slice(0, available));
  }, []);

  const takeOperations = useCallback((): Op[] => {
    return pendingOperationsRef.current.splice(0);
  }, []);

  return {
    acceptContent,
    ignoreChangeDuringExternalSync,
    mountRevision,
    recordOperations,
    takeOperations,
  };
}
