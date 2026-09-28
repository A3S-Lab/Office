import type { Cell, CellMatrix } from '@fortune-sheet/core';
import type { WorkbookInstance } from '@fortune-sheet/react';
import {
  registerDerivedSpreadsheetMatrix,
  spreadsheetMatrixProfile,
} from '../work-spreadsheet-matrix-profile';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetSheet,
} from '../work-types';
import { sparseArrayIndexes } from '../spreadsheet-sparse';

type SpreadsheetWorkbookUpdateTarget = Pick<
  WorkbookInstance,
  'updateSheet' | 'setCellValue'
>;

type SpreadsheetCellTextEdit = {
  column: number;
  row: number;
  sheetId: string;
  value: string;
};

/**
 * Replaces sheet data on an already-mounted workbook when only cell values
 * change. Fortune keeps selection / editing chrome outside `luckysheetfile`,
 * so structural sheet features must stay identical or we remount.
 *
 * Bot leaf typing paints one glyph per frame: remounting on each frame is what
 * made the spreadsheet flicker. Prefer `setCellValue` for text-only cell edits
 * so Fortune keeps the sheet container and remote presence overlays; fall back
 * to `updateSheet` for richer matrix swaps.
 */
export function synchronizeSpreadsheetWorkbookInPlace(
  workbook: SpreadsheetWorkbookUpdateTarget | null,
  previous: WorkSpreadsheetContent,
  next: WorkSpreadsheetContent,
  projectedSheets: WorkSpreadsheetContent['sheets'],
  preview: boolean,
): boolean {
  if (preview || !workbook || !sameSpreadsheetSheetIdentity(previous.sheets, next.sheets)) {
    return false;
  }

  for (let index = 0; index < next.sheets.length; index += 1) {
    if (
      spreadsheetSheetStructureChanged(previous.sheets[index]!, next.sheets[index]!)
    ) {
      return false;
    }
  }

  const cellEdits = spreadsheetCellTextEdits(previous, next);
  if (cellEdits) {
    certifyDerivedSpreadsheetMatrices(previous, next);
    try {
      for (const edit of cellEdits) {
        workbook.setCellValue(edit.row, edit.column, edit.value, {
          id: edit.sheetId,
        });
      }
      return true;
    } catch {
      return false;
    }
  }

  if (!sameSpreadsheetSheetIdentity(next.sheets, projectedSheets)) {
    return false;
  }

  certifyDerivedSpreadsheetMatrices(previous, next);

  for (let index = 0; index < next.sheets.length; index += 1) {
    const profile = spreadsheetMatrixProfile(projectedSheets[index]!.data);
    if (profile?.protectionCellKey) {
      return false;
    }
  }

  try {
    workbook.updateSheet(projectedSheets);
    return true;
  } catch {
    return false;
  }
}

/**
 * Leaf-typing frames clone row arrays and lose the WeakMap matrix profile.
 * Carry fortuneReady forward from the previous matrix so projection stays cheap
 * and in-place sync does not fall back to a full remount.
 */
export function certifyDerivedSpreadsheetMatrices(
  previous: WorkSpreadsheetContent,
  next: WorkSpreadsheetContent,
): void {
  for (let index = 0; index < next.sheets.length; index += 1) {
    const before = previous.sheets[index];
    const after = next.sheets[index];
    if (!before?.data || !after?.data || before.id !== after.id) continue;
    if (spreadsheetMatrixProfile(after.data)?.fortuneReady) continue;
    if (!spreadsheetMatrixProfile(before.data)?.fortuneReady) continue;
    registerDerivedSpreadsheetMatrix(
      after.data,
      before.data,
      spreadsheetMatrixCellChanges(before.data, after.data),
    );
  }
}

/**
 * Text-only cell deltas safe for Fortune `setCellValue`. Returns `null` when
 * any non-text cell field changes (formula, style, merge, …) so callers fall
 * back to `updateSheet` / remount.
 */
export function spreadsheetCellTextEdits(
  previous: WorkSpreadsheetContent,
  next: WorkSpreadsheetContent,
): SpreadsheetCellTextEdit[] | null {
  if (!sameSpreadsheetSheetIdentity(previous.sheets, next.sheets)) return null;
  const edits: SpreadsheetCellTextEdit[] = [];
  for (let index = 0; index < next.sheets.length; index += 1) {
    const previousSheet = previous.sheets[index]!;
    const nextSheet = next.sheets[index]!;
    const sheetId = nextSheet.id;
    if (!sheetId) return null;
    const previousData = previousSheet.data ?? [];
    const nextData = nextSheet.data ?? [];
    const rowIndexes = new Set<number>([
      ...sparseArrayIndexes(previousData),
      ...sparseArrayIndexes(nextData),
    ]);
    for (let row = 0; row < Math.max(previousData.length, nextData.length); row += 1) {
      rowIndexes.add(row);
    }
    for (const row of rowIndexes) {
      const previousRow = previousData[row];
      const nextRow = nextData[row];
      if (previousRow === nextRow) continue;
      const columnIndexes = new Set<number>([
        ...sparseArrayIndexes(previousRow ?? []),
        ...sparseArrayIndexes(nextRow ?? []),
      ]);
      for (
        let column = 0;
        column < Math.max(previousRow?.length ?? 0, nextRow?.length ?? 0);
        column += 1
      ) {
        columnIndexes.add(column);
      }
      for (const column of columnIndexes) {
        const before = previousRow?.[column];
        const after = nextRow?.[column];
        if (before === after) continue;
        if (!spreadsheetCellTextOnlyChange(before, after)) return null;
        const text = spreadsheetCellDisplayText(after);
        const previousText = spreadsheetCellDisplayText(before);
        if (text === previousText) continue;
        edits.push({ sheetId, row, column, value: text });
      }
    }
  }
  return edits;
}

function spreadsheetCellDisplayText(cell: Cell | null | undefined): string {
  if (!cell) return '';
  if (typeof cell.m === 'string') return cell.m;
  if (typeof cell.m === 'number') return String(cell.m);
  if (typeof cell.v === 'string') return cell.v;
  if (typeof cell.v === 'number' || typeof cell.v === 'boolean') {
    return String(cell.v);
  }
  return '';
}

/** True when two cells differ only in display/value text (v / m). */
function spreadsheetCellTextOnlyChange(
  previous: Cell | null | undefined,
  next: Cell | null | undefined,
): boolean {
  if (previous == null && next == null) return true;
  if (next == null) {
    // Clearing a cell is still a text edit when the previous cell was text-only.
    return spreadsheetCellIsPlainText(previous);
  }
  if (previous == null) return spreadsheetCellIsPlainText(next);
  if (!spreadsheetCellIsPlainText(previous) || !spreadsheetCellIsPlainText(next)) {
    return false;
  }
  const previousKeys = Object.keys(previous).filter((key) => key !== 'v' && key !== 'm');
  const nextKeys = Object.keys(next).filter((key) => key !== 'v' && key !== 'm');
  if (previousKeys.length !== nextKeys.length) return false;
  return previousKeys.every(
    (key) =>
      Object.hasOwn(next, key) &&
      JSON.stringify((previous as Record<string, unknown>)[key]) ===
        JSON.stringify((next as Record<string, unknown>)[key]),
  );
}

function spreadsheetCellIsPlainText(cell: Cell | null | undefined): boolean {
  if (!cell) return true;
  if (cell.f != null && cell.f !== '') return false;
  if (cell.ct && cell.ct.t && cell.ct.t !== 'g' && cell.ct.t !== 's') return false;
  return true;
}

function sameSpreadsheetSheetIdentity(
  left: readonly WorkSpreadsheetSheet[],
  right: readonly WorkSpreadsheetSheet[],
): boolean {
  return (
    left.length > 0 &&
    left.length === right.length &&
    left.every(
      (sheet, index) => Boolean(sheet.id) && sheet.id === right[index]?.id,
    )
  );
}

function spreadsheetSheetStructureChanged(
  left: WorkSpreadsheetSheet,
  right: WorkSpreadsheetSheet,
): boolean {
  return spreadsheetSheetStructureKey(left) !== spreadsheetSheetStructureKey(right);
}

/**
 * Everything Fortune cannot reliably refresh via `updateSheet` alone. Cell
 * matrices (`data` / `celldata`) are intentionally excluded — those are what
 * leaf typing and collab cell edits change every frame.
 */
function spreadsheetSheetStructureKey(sheet: WorkSpreadsheetSheet): string {
  return JSON.stringify([
    sheet.config ?? null,
    sheet.images ?? null,
    sheet.charts ?? null,
    sheet.pivotTables ?? null,
    sheet.pivotTable ?? null,
    sheet.isPivotTable ?? null,
    sheet.frozen ?? null,
    sheet.filter ?? null,
    sheet.filter_select ?? null,
    sheet.luckysheet_conditionformat_save ?? null,
    sheet.luckysheet_alternateformat_save ?? null,
    sheet.dataVerification ?? null,
    sheet.dataValidationRanges ?? null,
    sheet.hyperlink ?? null,
    sheet.dynamicArray_compute ?? null,
    sheet.dynamicArray ?? null,
    sheet.formulaMetadata ?? null,
    sheet.zoomRatio ?? null,
    sheet.defaultRowHeight ?? null,
    sheet.defaultColWidth ?? null,
    sheet.showGridLines ?? null,
    sheet.tables ?? null,
    sheet.row ?? null,
    sheet.column ?? null,
    sheet.name ?? null,
    sheet.order ?? null,
    sheet.status ?? null,
  ]);
}

function spreadsheetMatrixCellChanges(
  previous: CellMatrix,
  next: CellMatrix,
): Array<{
  column: number;
  current: Cell | null | undefined;
  previous: Cell | null | undefined;
  row: number;
}> {
  const changes: Array<{
    column: number;
    current: Cell | null | undefined;
    previous: Cell | null | undefined;
    row: number;
  }> = [];
  const rowIndexes = new Set<number>([
    ...sparseArrayIndexes(previous),
    ...sparseArrayIndexes(next),
  ]);
  for (let row = 0; row < Math.max(previous.length, next.length); row += 1) {
    rowIndexes.add(row);
  }
  for (const row of rowIndexes) {
    const previousRow = previous[row];
    const nextRow = next[row];
    if (previousRow === nextRow) continue;
    const columnIndexes = new Set<number>([
      ...sparseArrayIndexes(previousRow ?? []),
      ...sparseArrayIndexes(nextRow ?? []),
    ]);
    for (
      let column = 0;
      column < Math.max(previousRow?.length ?? 0, nextRow?.length ?? 0);
      column += 1
    ) {
      columnIndexes.add(column);
    }
    for (const column of columnIndexes) {
      const before = previousRow?.[column];
      const after = nextRow?.[column];
      if (before === after) continue;
      changes.push({ row, column, previous: before, current: after });
    }
  }
  return changes;
}
