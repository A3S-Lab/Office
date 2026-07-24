import type { WorkbookInstance } from '@fortune-sheet/react';
import { useCallback, useEffect, useMemo, useRef, type RefObject } from 'react';
import {
  createOfficeKernelClient,
  type OfficeKernelClient,
} from '../../../kernel/office-kernel-client';
import type { OfficeKernelSpreadsheetCalculationIssue } from '../../../kernel/office-kernel-protocol';
import { effectiveSpreadsheetCalculationSettings } from '../work-spreadsheet-formulas';
import type { WorkSpreadsheetContent } from '../work-types';
import {
  createSpreadsheetKernelWorkbook,
  refreshSpreadsheetKernelWorkbook,
  spreadsheetCalculationFallbackCells,
  spreadsheetCalculationOps,
  spreadsheetCalculationSourceKey,
  spreadsheetCalculationTargets,
  type SpreadsheetKernelWorkbook,
} from './spreadsheet-calculation-model';
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
  sourceKey: string;
}

export function useSpreadsheetCalculation({
  content,
  kernelWasmUrl,
  workbookRef,
}: UseSpreadsheetCalculationOptions): SpreadsheetCalculationCommandPort {
  const compiled = useMemo(
    () => createSpreadsheetKernelWorkbook(content),
    [content],
  );
  const compiledRef = useRef(compiled);
  const clientRef = useRef<OfficeKernelClient | null>(null);
  const activeRef = useRef<ActiveCalculation | null>(null);
  const fallbackTimerRef = useRef<ReturnType<typeof setTimeout> | null>(null);
  const revisionRef = useRef(0);
  const documentRevisionRef = useRef(0);
  compiledRef.current = compiled;

  useEffect(() => {
    const client = createOfficeKernelClient(kernelWasmUrl);
    clientRef.current = client;
    return () => {
      activeRef.current?.controller.abort();
      activeRef.current = null;
      clearFallbackTimer(fallbackTimerRef);
      clientRef.current = null;
      client.dispose();
    };
  }, [kernelWasmUrl]);

  const runCalculation = useCallback(
    (
      command: SpreadsheetCalculationCommand,
      snapshot: SpreadsheetKernelWorkbook,
      documentRevision: number,
      includeDataTables = true,
    ): void => {
      const client = clientRef.current;
      if (!client) {
        fallbackWithFortune(workbookRef.current, command);
        return;
      }
      activeRef.current?.controller.abort();
      activeRef.current = null;
      clearFallbackTimer(fallbackTimerRef);
      let calculationSnapshot = snapshot;
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
        sourceKey: calculationSnapshot.sourceKey,
      };
      activeRef.current = active;
      void client
        .spreadsheetCalculation(
          {
            revision: active.revision,
            documentRevision: active.documentRevision,
            sheets: calculationSnapshot.sheets,
            targets,
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
          if (
            !workbook ||
            spreadsheetCalculationSourceKey(workbook.getAllSheets()) !==
              active.sourceKey
          ) {
            return;
          }
          const ops = spreadsheetCalculationOps(
            workbook.getAllSheets(),
            result.cells,
          );
          if (ops.length) workbook.applyOp(ops);
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
          const workbook = workbookRef.current;
          if (
            workbook &&
            spreadsheetCalculationSourceKey(workbook.getAllSheets()) ===
              active.sourceKey
          ) {
            fallbackWithFortune(workbook, command);
          }
        })
        .finally(() => {
          if (isCurrentCalculation(activeRef.current, active)) {
            activeRef.current = null;
          }
        });
    },
    [workbookRef],
  );

  useEffect(() => {
    const documentRevision = ++documentRevisionRef.current;
    activeRef.current?.controller.abort();
    activeRef.current = null;
    clearFallbackTimer(fallbackTimerRef);
    const settings = effectiveSpreadsheetCalculationSettings(
      content.calculation,
    );
    if (settings.mode === 'manual') return;
    const timer = setTimeout(() => {
      if (!compiled) {
        fallbackWithFortune(workbookRef.current, { scope: 'workbook' });
        return;
      }
      if (
        !compiled.fallbackCells.length &&
        !compiled.sheets.some((sheet) =>
          sheet.cells.some((cell) => cell.formula),
        )
      ) {
        return;
      }
      runCalculation(
        { scope: 'workbook' },
        compiled,
        documentRevision,
        settings.mode !== 'automatic-except-data-tables',
      );
    }, 0);
    return () => clearTimeout(timer);
  }, [
    compiled?.sourceKey,
    content.calculation?.mode,
    runCalculation,
    workbookRef,
  ]);

  const recalculate = useCallback(
    (command: SpreadsheetCalculationCommand): void => {
      const snapshot = compiledRef.current;
      if (!snapshot) {
        fallbackWithFortune(workbookRef.current, command);
        return;
      }
      runCalculation(command, snapshot, documentRevisionRef.current);
    },
    [runCalculation, workbookRef],
  );

  return useMemo(() => ({ recalculate }), [recalculate]);
}

function scheduleIssueFallback(
  workbookRef: RefObject<WorkbookInstance | null>,
  issues: readonly OfficeKernelSpreadsheetCalculationIssue[],
  active: ActiveCalculation,
  revisionRef: RefObject<number>,
  timerRef: RefObject<ReturnType<typeof setTimeout> | null>,
): void {
  if (!issues.length) return;
  clearFallbackTimer(timerRef);
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

function clearFallbackTimer(
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
    current.sourceKey === expected.sourceKey &&
    !expected.controller.signal.aborted
  );
}

function isAbortError(error: unknown): boolean {
  return error instanceof DOMException && error.name === 'AbortError';
}
