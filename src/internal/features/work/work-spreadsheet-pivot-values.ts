import type { Cell } from '@fortune-sheet/core';
import type {
  WorkSpreadsheetPivotFilterValue,
  WorkSpreadsheetPivotReportFilter,
} from './work-types';

export type SpreadsheetPivotReportFilterSelection =
  | { kind: 'all' }
  | { kind: 'none' }
  | { kind: 'items'; items: WorkSpreadsheetPivotFilterValue[] };

export function spreadsheetPivotReportFilterSelection(
  filter: WorkSpreadsheetPivotReportFilter,
): SpreadsheetPivotReportFilterSelection {
  if (filter.selectedItems !== undefined) {
    if (!filter.selectedItems.length) return { kind: 'none' };
    return { kind: 'items', items: filter.selectedItems };
  }
  if (filter.selectedItem !== undefined) {
    return { kind: 'items', items: [filter.selectedItem] };
  }
  return { kind: 'all' };
}

export function displaySpreadsheetPivotReportFilterSelection(
  filter: WorkSpreadsheetPivotReportFilter,
): string {
  const selection = spreadsheetPivotReportFilterSelection(filter);
  if (selection.kind === 'all') return '(全部)';
  if (selection.kind === 'none') return '(无)';
  return selection.items.map(displaySpreadsheetPivotValue).join(', ');
}

/** True when native XLSX pageFields can represent the filter without multi-select XML. */
export function spreadsheetPivotReportFilterIsSingleSelectExportable(
  filter: WorkSpreadsheetPivotReportFilter,
): boolean {
  if (filter.selectedItems === undefined) return true;
  return filter.selectedItems.length === 1;
}

export function spreadsheetPivotCellValue(
  cell: Cell | null | undefined,
): unknown {
  return cell?.v ?? cell?.m ?? null;
}

export function normalizeSpreadsheetPivotFilterValue(
  value: unknown,
): WorkSpreadsheetPivotFilterValue {
  if (value === null || value === undefined || value === '') return null;
  if (value instanceof Date) return value.toISOString();
  if (typeof value === 'number')
    return Number.isFinite(value) ? value : String(value);
  if (typeof value === 'boolean') return value;
  return String(value);
}

export function spreadsheetPivotFilterValueKey(
  value: WorkSpreadsheetPivotFilterValue,
): string {
  if (value === null) return 'blank:';
  return `${typeof value}:${String(value)}`;
}

export function displaySpreadsheetPivotValue(value: unknown): string {
  if (value === null || value === undefined || value === '') return '(空白)';
  if (value instanceof Date) return value.toISOString().slice(0, 10);
  if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
  return String(value);
}

export function finiteSpreadsheetPivotNumber(value: unknown): number | null {
  if (typeof value === 'number') return Number.isFinite(value) ? value : null;
  if (value instanceof Date) return value.getTime();
  return null;
}
