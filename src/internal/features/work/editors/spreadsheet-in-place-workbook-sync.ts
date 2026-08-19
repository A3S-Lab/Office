import type { WorkbookInstance } from '@fortune-sheet/react';
import { spreadsheetMatrixProfile } from '../work-spreadsheet-matrix-profile';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetSheet,
} from '../work-types';

type SpreadsheetWorkbookUpdateTarget = Pick<WorkbookInstance, 'updateSheet'>;

/**
 * Replaces only the data of an already-mounted, structurally simple workbook.
 * Fortune keeps additional active-sheet state outside `luckysheetfile`, so
 * richer workbook structures must continue through the remount path.
 */
export function synchronizeSpreadsheetWorkbookInPlace(
  workbook: SpreadsheetWorkbookUpdateTarget | null,
  previous: WorkSpreadsheetContent,
  next: WorkSpreadsheetContent,
  projectedSheets: WorkSpreadsheetContent['sheets'],
  preview: boolean,
): boolean {
  if (
    preview ||
    !workbook ||
    !sameSpreadsheetSheetIdentity(previous.sheets, next.sheets) ||
    !sameSpreadsheetSheetIdentity(next.sheets, projectedSheets)
  ) {
    return false;
  }

  for (const sheet of projectedSheets) {
    const profile = spreadsheetMatrixProfile(sheet.data);
    if (
      !profile?.fortuneReady ||
      profile.protectionCellKey ||
      spreadsheetSheetRequiresRemount(sheet)
    ) {
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

function spreadsheetSheetRequiresRemount(sheet: WorkSpreadsheetSheet): boolean {
  return Boolean(
    nonEmptyRecord(sheet.config) ||
      nonEmptyArray(sheet.images) ||
      nonEmptyArray(sheet.charts) ||
      nonEmptyArray(sheet.pivotTables) ||
      sheet.pivotTable ||
      sheet.isPivotTable ||
      sheet.frozen ||
      nonEmptyRecord(sheet.filter) ||
      sheet.filter_select ||
      nonEmptyArray(sheet.luckysheet_conditionformat_save) ||
      nonEmptyArray(sheet.luckysheet_alternateformat_save) ||
      nonEmptyRecord(sheet.dataVerification) ||
      nonEmptyArray(sheet.dataValidationRanges) ||
      nonEmptyRecord(sheet.hyperlink) ||
      sheet.dynamicArray_compute ||
      nonEmptyArray(sheet.dynamicArray) ||
      sheet.formulaMetadata ||
      sheet.zoomRatio !== undefined ||
      sheet.defaultRowHeight !== undefined ||
      sheet.defaultColWidth !== undefined ||
      sheet.showGridLines !== undefined,
  );
}

function nonEmptyArray(value: readonly unknown[] | undefined): boolean {
  return Boolean(value?.length);
}

function nonEmptyRecord(value: unknown): boolean {
  return Boolean(
    value &&
      typeof value === 'object' &&
      !Array.isArray(value) &&
      Object.keys(value).length,
  );
}
