import { WorkOfficeCollaborationError } from './office-collaboration';
import {
  cloneWorkOfficeCollaborationJson as cloneJsonValue,
  isWorkOfficeCollaborationRecord as isRecord,
} from './office-collaboration-json';

export function requiredCoordinate(
  value: unknown,
  maximum: number,
  label: string,
  sheetId: string,
): number {
  if (
    !Number.isSafeInteger(value) ||
    (value as number) < 0 ||
    (value as number) >= maximum
  ) {
    invalidWorkOfficeSpreadsheetInput(
      `a valid ${label} coordinate in sheet '${sheetId}'`,
    );
  }
  return value as number;
}

export function requiredInputRecord(
  value: unknown,
  label: string,
): Record<string, unknown> {
  if (!isRecord(value)) {
    invalidWorkOfficeSpreadsheetInput(`a valid ${label} record`);
  }
  return value as Record<string, unknown>;
}

export function requiredIdentifier(value: unknown, label: string): string {
  const result = requiredNonEmptyString(value, `${label} ID`);
  if (result !== result.trim() || result.length > 256) {
    invalidWorkOfficeSpreadsheetInput(
      `a ${label} ID containing 1 to 256 characters`,
    );
  }
  return result;
}

export function requiredNonEmptyString(value: unknown, label: string): string {
  if (typeof value !== 'string' || !value.trim() || value.length > 256) {
    invalidWorkOfficeSpreadsheetInput(
      `a non-empty string of at most 256 characters for ${label}`,
    );
  }
  return value as string;
}

export function validateJsonRecord(
  value: Record<string, unknown>,
  label: string,
): Record<string, unknown> {
  try {
    return cloneJsonValue(value) as Record<string, unknown>;
  } catch (error) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.content_invalid',
      `Spreadsheet collaboration requires a JSON-compatible ${label}: ${error instanceof Error ? error.message : String(error)}`,
    );
  }
}

export function invalidWorkOfficeSpreadsheetInput(expected: string): never {
  throw new WorkOfficeCollaborationError(
    'office.collaboration.content_invalid',
    `Spreadsheet collaboration requires ${expected}.`,
  );
}
