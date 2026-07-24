import type { Op } from '@fortune-sheet/core';
import type { WorkbookInstance } from '@fortune-sheet/react';
import { useCallback, useEffect, useMemo, useRef, type RefObject } from 'react';
import {
  createOfficeKernelClient,
  type OfficeKernelClient,
} from '../../../kernel/office-kernel-client';
import type {
  OfficeKernelSpreadsheetCalculationIssue,
  OfficeKernelSpreadsheetSessionCellChange,
} from '../../../kernel/office-kernel-protocol';
import { effectiveSpreadsheetCalculationSettings } from '../work-spreadsheet-formulas';
import type { WorkSpreadsheetContent } from '../work-types';
import {
  spreadsheetCalculationFallbackCells,
  spreadsheetCalculationOps,
  spreadsheetCalculationSessionUpdate,
  spreadsheetCalculationTargets,
} from './spreadsheet-calculation-model';
import {
  createSpreadsheetKernelWorkbook,
  projectSpreadsheetKernelWorkbookOperations,
  refreshSpreadsheetKernelWorkbook,
  spreadsheetOperationsMayChangeCalculation,
  type SpreadsheetKernelOperationProjection,
  type SpreadsheetKernelWorkbook,
} from './spreadsheet-calculation-projection';
import type {
  SpreadsheetCalculationCommand,
  SpreadsheetCalculationCommandPort,
} from './spreadsheet-command-controller';

interface UseSpreadsheetCalculationOptions {
  content: WorkSpreadsheetContent;
  kernelWasmUrl?: string;
  workbookRef: RefObject<WorkbookInstance | null>;
}

interface ActiveCalculation {
  controller: AbortController;
  documentRevision: number;
  revision: number;
}

interface KernelSessionSnapshot {
  documentRevision: number;
  workbook: SpreadsheetKernelWorkbook;
}

interface PreparedSessionPatch {
  baseSession: KernelSessionSnapshot;
  changes: OfficeKernelSpreadsheetSessionCellChange[];
}

interface SynchronizeSnapshotOptions {
  baseWorkbook?: SpreadsheetKernelWorkbook;
  forceCalculation?: boolean;
  projection?: SpreadsheetKernelOperationProjection;
  sourceChanged?: boolean;
}

export interface SpreadsheetCalculationController
  extends SpreadsheetCalculationCommandPort {
  hasPendingResultPatches: () => boolean;
  notifyWorkbookOperations: (operations: readonly Op[]) => void;
  synchronizeWorkbook: (
    content: WorkSpreadsheetContent,
    operations: readonly Op[],
  ) => void;
}

export function useSpreadsheetCalculation({
  content,
  kernelWasmUrl,
  workbookRef,
}: UseSpreadsheetCalculationOptions): SpreadsheetCalculationController {
  const latestContentRef = useRef(content);
  const snapshotRef = useRef<SpreadsheetKernelWorkbook | null>(null);
  const clientRef = useRef<OfficeKernelClient | null>(null);
  const sessionRef = useRef<KernelSessionSnapshot | null>(null);
  const activeRef = useRef<ActiveCalculation | null>(null);
  const acceptedContentRef = useRef<WorkSpreadsheetContent | null>(null);
  const resultOperationsRef = useRef<Op[]>([]);
  const pendingOperationCancellationRef = useRef(false);
  const automaticTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const fallbackTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const revisionRef = useRef(0);
  const documentRevisionRef = useRef(0);
  const calculationSettingsKeyRef = useRef<string | null>(null);
  latestContentRef.current = content;

  useEffect(() => {
    const client = createOfficeKernelClient(kernelWasmUrl);
    clientRef.current = client;
    sessionRef.current = null;
    snapshotRef.current = null;
    acceptedContentRef.current = null;
    resultOperationsRef.current = [];
    pendingOperationCancellationRef.current = false;
    calculationSettingsKeyRef.current = null;
    return () => {
      activeRef.current?.controller.abort();
      activeRef.current = null;
      clearTimer(automaticTimerRef);
      clearTimer(fallbackTimerRef);
      clientRef.current = null;
      sessionRef.current = null;
      snapshotRef.current = null;
      client.dispose();
    };
  }, [kernelWasmUrl]);

  const runCalculation = useCallback(
    (
      command: SpreadsheetCalculationCommand,
      snapshot: SpreadsheetKernelWorkbook,
      documentRevision: number,
      includeDataTables = true,
      automatic = false,
      preparedPatch?: PreparedSessionPatch,
    ): void => {
      const client = clientRef.current;
      if (!client) {
        fallbackWithFortune(workbookRef.current, command);
        return;
      }
      activeRef.current?.controller.abort();
      activeRef.current = null;
      clearTimer(fallbackTimerRef);
      let calculationSnapshot = snapshot;
      let forceSessionReplace = false;
      const compatibilityCells = spreadsheetCalculationFallbackCells(
        snapshot,
        command,
        includeDataTables,
      );
      if (compatibilityCells.length) {
        const workbook = workbookRef.current;
        if (!workbook) return;
        for (const cell of compatibilityCells) {
          workbook.calculateFormula(cell.sheetId, {
            row: [cell.row, cell.row],
            column: [cell.column, cell.column],
          });
        }
        const refreshed = refreshSpreadsheetKernelWorkbook(
          workbook.getAllSheets(),
          snapshot.fallbackCells,
        );
        if (!refreshed) {
          fallbackWithFortune(workbook, command);
          return;
        }
        calculationSnapshot = refreshed;
        snapshotRef.current = refreshed;
        forceSessionReplace = true;
      }
      const targets = spreadsheetCalculationTargets(
        calculationSnapshot,
        command,
      );
      const hasKernelTargets =
        targets === undefined
          ? calculationSnapshot.sheets.some((sheet) =>
              sheet.cells.some((cell) => cell.formula),
            )
          : targets.length > 0;
      if (!hasKernelTargets) return;

      const active: ActiveCalculation = {
        controller: new AbortController(),
        documentRevision,
        revision: ++revisionRef.current,
      };
      activeRef.current = active;
      const previousSession = sessionRef.current;
      const update =
        !forceSessionReplace &&
        previousSession &&
        preparedPatch?.baseSession === previousSession
          ? {
              kind: 'patch' as const,
              baseDocumentRevision: previousSession.documentRevision,
              changes: preparedPatch.changes,
            }
          : spreadsheetCalculationSessionUpdate(
              previousSession?.workbook ?? null,
              calculationSnapshot,
              previousSession?.documentRevision ?? 0,
              forceSessionReplace,
            );
      const calculation =
        command.scope === 'selection'
          ? { kind: 'targets' as const, targets: targets ?? [] }
          : automatic && update.kind === 'patch'
            ? { kind: 'dirty' as const }
            : { kind: 'workbook' as const };
      sessionRef.current = {
        documentRevision: active.documentRevision,
        workbook: calculationSnapshot,
      };
      void client
        .spreadsheetSessionCalculation(
          {
            revision: active.revision,
            documentRevision: active.documentRevision,
            update,
            calculation,
            fallbackSheets: calculationSnapshot.sheets,
          },
          active.controller.signal,
        )
        .then((result) => {
          if (
            !isCurrentCalculation(activeRef.current, active) ||
            result.revision !== active.revision ||
            result.documentRevision !== active.documentRevision
          ) {
            return;
          }
          const workbook = workbookRef.current;
          if (!workbook) return;
          const ops = spreadsheetCalculationOps(
            workbook.getAllSheets(),
            result.cells,
          );
          if (ops.length) {
            resultOperationsRef.current.push(...ops);
            workbook.applyOp(ops);
          }
          scheduleIssueFallback(
            workbookRef,
            result.issues,
            active,
            revisionRef,
            fallbackTimerRef,
          );
        })
        .catch((error: unknown) => {
          if (
            isAbortError(error) ||
            !isCurrentCalculation(activeRef.current, active)
          ) {
            return;
          }
          if (
            sessionRef.current?.documentRevision === active.documentRevision
          ) {
            sessionRef.current = null;
          }
          fallbackWithFortune(workbookRef.current, command);
        })
        .finally(() => {
          if (isCurrentCalculation(activeRef.current, active)) {
            activeRef.current = null;
          }
        });
    },
    [workbookRef],
  );

  const synchronizeSnapshot = useCallback(
    (
      nextContent: WorkSpreadsheetContent,
      nextSnapshot: SpreadsheetKernelWorkbook | null,
      options: SynchronizeSnapshotOptions = {},
    ): void => {
      const settings = effectiveSpreadsheetCalculationSettings(
        nextContent.calculation,
      );
      const settingsKey = settings.mode;
      const previousSnapshot = snapshotRef.current;
      const sourceChanged =
        options.sourceChanged ??
        previousSnapshot?.sourceKey !== nextSnapshot?.sourceKey;
      const settingsChanged = calculationSettingsKeyRef.current !== settingsKey;
      const calculationSnapshot =
        !sourceChanged && previousSnapshot ? previousSnapshot : nextSnapshot;
      calculationSettingsKeyRef.current = settingsKey;
      snapshotRef.current = calculationSnapshot;
      if (!sourceChanged && !settingsChanged && !options.forceCalculation) {
        return;
      }

      activeRef.current?.controller.abort();
      activeRef.current = null;
      clearTimer(automaticTimerRef);
      clearTimer(fallbackTimerRef);
      const documentRevision = ++documentRevisionRef.current;
      if (settings.mode === 'manual') return;

      const baseSession = sessionRef.current;
      const preparedPatch =
        options.projection?.sourceChanged &&
        options.baseWorkbook &&
        baseSession?.workbook === options.baseWorkbook
          ? {
              baseSession,
              changes: options.projection.changes,
            }
          : undefined;
      automaticTimerRef.current = setTimeout(() => {
        automaticTimerRef.current = null;
        if (documentRevisionRef.current !== documentRevision) return;
        if (!calculationSnapshot) {
          fallbackWithFortune(workbookRef.current, { scope: 'workbook' });
          return;
        }
        if (
          !calculationSnapshot.fallbackCells.length &&
          !calculationSnapshot.sheets.some((sheet) =>
            sheet.cells.some((cell) => cell.formula),
          )
        ) {
          return;
        }
        runCalculation(
          { scope: 'workbook' },
          calculationSnapshot,
          documentRevision,
          settings.mode !== 'automatic-except-data-tables',
          true,
          preparedPatch,
        );
      }, 0);
    },
    [runCalculation, workbookRef],
  );

  const synchronizeWorkbook = useCallback(
    (nextContent: WorkSpreadsheetContent, operations: readonly Op[]): void => {
      acceptedContentRef.current = nextContent;
      const resultOperations = resultOperationsRef.current.splice(0);
      const combinedOperations = [...resultOperations, ...operations];
      const baseWorkbook = snapshotRef.current;
      const forceCalculation = pendingOperationCancellationRef.current;
      pendingOperationCancellationRef.current = false;
      const canProject =
        Boolean(baseWorkbook) &&
        combinedOperations.length > 0 &&
        !nextContent.sheets.some((sheet) => sheet.pivotTables?.length);
      const projection =
        canProject && baseWorkbook
          ? projectSpreadsheetKernelWorkbookOperations(
              baseWorkbook,
              nextContent.sheets,
              combinedOperations,
            )
          : null;
      if (projection) {
        synchronizeSnapshot(nextContent, projection.workbook, {
          baseWorkbook: baseWorkbook ?? undefined,
          forceCalculation,
          projection,
          sourceChanged: projection.sourceChanged,
        });
        return;
      }
      synchronizeSnapshot(
        nextContent,
        createSpreadsheetKernelWorkbook(nextContent),
        { forceCalculation },
      );
    },
    [synchronizeSnapshot],
  );

  const notifyWorkbookOperations = useCallback(
    (operations: readonly Op[]): void => {
      if (
        !spreadsheetOperationsMayChangeCalculation(
          operations,
          snapshotRef.current,
        )
      ) {
        return;
      }
      pendingOperationCancellationRef.current = true;
      activeRef.current?.controller.abort();
      activeRef.current = null;
      clearTimer(automaticTimerRef);
      clearTimer(fallbackTimerRef);
    },
    [],
  );

  useEffect(() => {
    if (acceptedContentRef.current === content) {
      acceptedContentRef.current = null;
      return;
    }
    acceptedContentRef.current = null;
    resultOperationsRef.current = [];
    pendingOperationCancellationRef.current = false;
    synchronizeSnapshot(content, createSpreadsheetKernelWorkbook(content));
  }, [content, kernelWasmUrl, synchronizeSnapshot]);

  const recalculate = useCallback(
    (command: SpreadsheetCalculationCommand): void => {
      clearTimer(automaticTimerRef);
      let snapshot = snapshotRef.current;
      if (!snapshot) {
        snapshot = createSpreadsheetKernelWorkbook(latestContentRef.current);
        snapshotRef.current = snapshot;
      }
      if (!snapshot) {
        fallbackWithFortune(workbookRef.current, command);
        return;
      }
      runCalculation(command, snapshot, documentRevisionRef.current);
    },
    [runCalculation, workbookRef],
  );

  const hasPendingResultPatches = useCallback(
    (): boolean => resultOperationsRef.current.length > 0,
    [],
  );

  return useMemo(
    () => ({
      hasPendingResultPatches,
      notifyWorkbookOperations,
      recalculate,
      synchronizeWorkbook,
    }),
    [
      hasPendingResultPatches,
      notifyWorkbookOperations,
      recalculate,
      synchronizeWorkbook,
    ],
  );
}

function scheduleIssueFallback(
  workbookRef: RefObject<WorkbookInstance | null>,
  issues: readonly OfficeKernelSpreadsheetCalculationIssue[],
  active: ActiveCalculation,
  revisionRef: RefObject<number>,
  timerRef: RefObject<ReturnType<typeof setTimeout> | null>,
): void {
  if (!issues.length) return;
  clearTimer(timerRef);
  timerRef.current = setTimeout(() => {
    timerRef.current = null;
    if (
      revisionRef.current !== active.revision ||
      active.controller.signal.aborted
    ) {
      return;
    }
    const workbook = workbookRef.current;
    if (!workbook) return;
    for (const issue of issues) {
      workbook.calculateFormula(issue.cell.sheetId, {
        row: [issue.cell.row, issue.cell.row],
        column: [issue.cell.column, issue.cell.column],
      });
    }
  }, 0);
}

function fallbackWithFortune(
  workbook: WorkbookInstance | null,
  command: SpreadsheetCalculationCommand,
): void {
  if (!workbook) return;
  if (command.scope === 'workbook') {
    workbook.calculateFormula();
  } else {
    workbook.calculateFormula(command.sheetId, command.range);
  }
}

function clearTimer(
  timerRef: RefObject<ReturnType<typeof setTimeout> | null>,
): void {
  if (timerRef.current !== null) clearTimeout(timerRef.current);
  timerRef.current = null;
}

function isCurrentCalculation(
  current: ActiveCalculation | null,
  expected: ActiveCalculation,
): boolean {
  return (
    current?.revision === expected.revision &&
    current.documentRevision === expected.documentRevision &&
    !expected.controller.signal.aborted
  );
}

function isAbortError(error: unknown): boolean {
  return error instanceof DOMException && error.name === 'AbortError';
}
