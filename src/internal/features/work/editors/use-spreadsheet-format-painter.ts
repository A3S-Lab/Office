import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import type { WorkSpreadsheetContent } from '../work-types';
import type {
  SpreadsheetCommandRange,
  SpreadsheetCommandSelection,
  SpreadsheetFormatPainterCommandPort,
  SpreadsheetWorkbookCommandPort,
} from './spreadsheet-command-controller';
import {
  captureSpreadsheetFormatPattern,
  type SpreadsheetFormatPainterMode,
  type SpreadsheetFormatPattern,
  spreadsheetFormatPainterBatches,
  spreadsheetFormatPainterCellCount,
  spreadsheetFormatPainterTargetRange,
} from './spreadsheet-format-painter';

export type { SpreadsheetFormatPainterMode } from './spreadsheet-format-painter';

export const spreadsheetFormatPainterMaximumCells = 50_000;

interface SpreadsheetFormatPainterSession {
  lastAppliedTargetKey?: string;
  mode: SpreadsheetFormatPainterMode;
  pattern: SpreadsheetFormatPattern;
}

export interface UseSpreadsheetFormatPainterOptions {
  content: WorkSpreadsheetContent;
  editable: boolean;
  onError: (message: string) => void;
  sourceRange: SpreadsheetCommandRange;
  sourceSheetId: string;
  workbook: SpreadsheetWorkbookCommandPort | null;
}

export interface SpreadsheetFormatPainterController {
  commandPort: SpreadsheetFormatPainterCommandPort;
  mode: SpreadsheetFormatPainterMode | null;
}

export function useSpreadsheetFormatPainter({
  content,
  editable,
  onError,
  sourceRange,
  sourceSheetId,
  workbook,
}: UseSpreadsheetFormatPainterOptions): SpreadsheetFormatPainterController {
  const contentRef = useRef(content);
  const editableRef = useRef(editable);
  const onErrorRef = useRef(onError);
  const sessionRef = useRef<SpreadsheetFormatPainterSession | null>(null);
  const sourceRef = useRef({
    range: copySpreadsheetFormatPainterRange(sourceRange),
    sheetId: sourceSheetId,
  });
  const workbookRef = useRef(workbook);
  const [mode, setMode] = useState<SpreadsheetFormatPainterMode | null>(null);
  contentRef.current = content;
  editableRef.current = editable;
  onErrorRef.current = onError;
  sourceRef.current = {
    range: copySpreadsheetFormatPainterRange(sourceRange),
    sheetId: sourceSheetId,
  };
  workbookRef.current = workbook;

  const clearSession = useCallback(() => {
    sessionRef.current = null;
    setMode(null);
  }, []);

  useEffect(() => {
    if (!editable || !workbook) clearSession();
  }, [clearSession, editable, workbook]);

  const activate = useCallback(
    (nextMode: SpreadsheetFormatPainterMode): boolean => {
      const activeWorkbook = workbookRef.current;
      const source = sourceRef.current;
      if (!editableRef.current || !activeWorkbook || !source.sheetId) {
        return false;
      }
      if (
        spreadsheetFormatPainterCellCount(source.range) >
        spreadsheetFormatPainterMaximumCells
      ) {
        onErrorRef.current(
          `格式刷一次最多复制 ${spreadsheetFormatPainterMaximumCells.toLocaleString()} 个单元格，请缩小源区域后重试。`,
        );
        return false;
      }
      try {
        const pattern = captureSpreadsheetFormatPattern(
          activeWorkbook.getCellsByRange(source.range, {
            id: source.sheetId,
          }),
        );
        if (
          !pattern ||
          !spreadsheetFormatPatternMatchesRange(pattern, source.range)
        ) {
          onErrorRef.current('无法读取完整的源格式，请重新选择后重试。');
          return false;
        }
        sessionRef.current = { mode: nextMode, pattern };
        setMode(nextMode);
        return true;
      } catch {
        onErrorRef.current('无法读取源格式，请重新选择后重试。');
        return false;
      }
    },
    [],
  );

  const applySelection = useCallback(
    (target: SpreadsheetCommandSelection): boolean => {
      const activeWorkbook = workbookRef.current;
      const session = sessionRef.current;
      const sheet = contentRef.current.sheets.find(
        (candidate) => candidate.id === target.sheetId,
      );
      if (!editableRef.current || !activeWorkbook || !session || !sheet) {
        return false;
      }
      const rawTarget = spreadsheetSelectionRange(target);
      const targetRange = spreadsheetFormatPainterTargetRange(
        rawTarget,
        session.pattern,
        spreadsheetFormatPainterSheetBounds(sheet, rawTarget, session.pattern),
      );
      if (!targetRange) return false;
      const targetKey = spreadsheetFormatPainterTargetKey(
        target.sheetId,
        targetRange,
      );
      if (session.lastAppliedTargetKey === targetKey) return false;
      if (
        spreadsheetFormatPainterCellCount(targetRange) >
        spreadsheetFormatPainterMaximumCells
      ) {
        onErrorRef.current(
          `格式刷一次最多应用到 ${spreadsheetFormatPainterMaximumCells.toLocaleString()} 个单元格，请缩小目标区域后重试。`,
        );
        return false;
      }
      const batches = spreadsheetFormatPainterBatches(
        session.pattern,
        targetRange,
      );
      if (!batches.length) return false;
      session.lastAppliedTargetKey = targetKey;
      try {
        activeWorkbook.batchCallApis(
          batches.map((batch) => ({
            name: 'setCellFormatByRange',
            args: [
              batch.attribute,
              batch.value,
              batch.ranges,
              { id: target.sheetId },
            ],
          })),
        );
      } catch {
        session.lastAppliedTargetKey = undefined;
        onErrorRef.current('无法应用格式，请缩小目标区域或重试。');
        return false;
      }
      if (session.mode === 'once') clearSession();
      return true;
    },
    [clearSession],
  );

  const cancel = useCallback((): boolean => {
    if (!sessionRef.current) return false;
    clearSession();
    return true;
  }, [clearSession]);

  const commandPort = useMemo<SpreadsheetFormatPainterCommandPort>(
    () => ({
      active: mode !== null,
      activate,
      applySelection,
      canActivate: editable && Boolean(workbook && sourceSheetId),
      cancel,
      mode,
    }),
    [activate, applySelection, cancel, editable, mode, sourceSheetId, workbook],
  );

  return { commandPort, mode };
}

function spreadsheetFormatPatternMatchesRange(
  pattern: SpreadsheetFormatPattern,
  range: SpreadsheetCommandRange,
): boolean {
  return (
    pattern.rowCount === Math.abs(range.row[1] - range.row[0]) + 1 &&
    pattern.columnCount === Math.abs(range.column[1] - range.column[0]) + 1
  );
}

function spreadsheetSelectionRange(
  target: SpreadsheetCommandSelection,
): SpreadsheetCommandRange {
  return {
    row: [...target.selection.row],
    column: [...target.selection.column],
  };
}

function copySpreadsheetFormatPainterRange(
  range: SpreadsheetCommandRange,
): SpreadsheetCommandRange {
  return {
    row: [...range.row],
    column: [...range.column],
  };
}

function spreadsheetFormatPainterTargetKey(
  sheetId: string,
  range: SpreadsheetCommandRange,
): string {
  return `${sheetId}:${range.row[0]}:${range.row[1]}:${range.column[0]}:${range.column[1]}`;
}

function spreadsheetFormatPainterSheetBounds(
  sheet: WorkSpreadsheetContent['sheets'][number],
  target: SpreadsheetCommandRange,
  pattern: SpreadsheetFormatPattern,
): { columnCount: number; rowCount: number } {
  const requestedRowCount =
    Math.max(target.row[0], target.row[1]) + pattern.rowCount;
  const requestedColumnCount =
    Math.max(target.column[0], target.column[1]) + pattern.columnCount;
  return {
    rowCount:
      positiveInteger(sheet.row) ??
      Math.max(sheet.data?.length ?? 0, requestedRowCount),
    columnCount:
      positiveInteger(sheet.column) ??
      Math.max(
        ...(sheet.data?.map((row) => row?.length ?? 0) ?? [0]),
        requestedColumnCount,
      ),
  };
}

function positiveInteger(value: unknown): number | null {
  return typeof value === 'number' && Number.isFinite(value) && value > 0
    ? Math.floor(value)
    : null;
}
