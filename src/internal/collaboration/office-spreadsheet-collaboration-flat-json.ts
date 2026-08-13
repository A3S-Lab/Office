import {
  patchWorkOfficeCollaborationFlatJsonMap,
  readWorkOfficeCollaborationFlatJsonMap,
  type WorkOfficeCollaborationFlatJsonMap,
} from './office-collaboration-flat-json';
import { invalidWorkOfficeSpreadsheetShared as invalidSharedSpreadsheet } from './office-spreadsheet-collaboration-validation';

/**
 * Recursively flattens object fields into one stable Y.Map. Arrays remain atomic
 * because their positional meaning varies across OOXML features. This avoids a
 * concurrent first-write race for nested Y.Map instances while still letting
 * independent object leaves converge.
 */
export function patchSpreadsheetFlatJsonMap(
  target: SpreadsheetFlatJsonMap,
  previous: Record<string, unknown> | undefined,
  next: Record<string, unknown> | undefined,
  label: string,
): void {
  patchWorkOfficeCollaborationFlatJsonMap(target, previous, next, label);
}

export function readSpreadsheetFlatJsonMap(
  target: SpreadsheetFlatJsonMap,
  label: string,
): Record<string, unknown> {
  return readWorkOfficeCollaborationFlatJsonMap(
    target,
    label,
    invalidSharedSpreadsheet,
  );
}

export type SpreadsheetFlatJsonMap = WorkOfficeCollaborationFlatJsonMap;
