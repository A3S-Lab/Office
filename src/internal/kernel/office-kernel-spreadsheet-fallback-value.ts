import type {
  OfficeKernelSpreadsheetError,
  OfficeKernelSpreadsheetValue,
} from './office-kernel-spreadsheet-protocol';
import {
  isOfficeKernelSpreadsheetError,
  OFFICE_KERNEL_SPREADSHEET_MAX_TEXT_BYTES,
} from './office-kernel-spreadsheet-protocol';

const textEncoder = new TextEncoder();

export function spreadsheetValueFromParser(
  value: unknown,
): OfficeKernelSpreadsheetValue {
  if (Array.isArray(value)) {
    const first = value[0];
    return spreadsheetValueFromParser(Array.isArray(first) ? first[0] : first);
  }
  if (value === null || value === undefined || value === '') {
    return { kind: 'blank' };
  }
  if (typeof value === 'number') {
    return Number.isFinite(value)
      ? { kind: 'number', value: Object.is(value, -0) ? 0 : value }
      : { kind: 'error', value: '#NUM!' };
  }
  if (typeof value === 'boolean') return { kind: 'boolean', value };
  if (value instanceof Error) {
    return { kind: 'error', value: spreadsheetError(value.message) };
  }
  if (typeof value === 'string' && !value.startsWith('#')) {
    const error = recognizedSpreadsheetError(value);
    if (error) return { kind: 'error', value: error };
  }
  const text = String(value);
  return textEncoder.encode(text).byteLength <=
    OFFICE_KERNEL_SPREADSHEET_MAX_TEXT_BYTES
    ? { kind: 'text', value: text }
    : { kind: 'error', value: '#VALUE!' };
}

export function spreadsheetValueForParser(
  value: OfficeKernelSpreadsheetValue,
): unknown {
  switch (value.kind) {
    case 'blank':
      return null;
    case 'number':
    case 'text':
    case 'boolean':
      return value.value;
    case 'error':
      return new Error(value.value);
  }
}

export function spreadsheetError(value: string): OfficeKernelSpreadsheetError {
  return recognizedSpreadsheetError(value) ?? '#VALUE!';
}

export function recognizedSpreadsheetError(
  value: string,
): OfficeKernelSpreadsheetError | null {
  const normalized = value.startsWith('#') ? value : `#${value}`;
  const withSuffix = normalized === '#DIV/0' ? '#DIV/0!' : normalized;
  return isOfficeKernelSpreadsheetError(withSuffix) ? withSuffix : null;
}

export function evaluateParserIfError(parameters: unknown[]): unknown {
  const [value, fallback = null] = parameters;
  return isParserErrorValue(value) ? fallback : value;
}

function isParserErrorValue(value: unknown): boolean {
  if (value instanceof Error) return true;
  if (typeof value !== 'string' || value.startsWith('#')) return false;
  const normalized = `#${value}`;
  const withSuffix = normalized === '#DIV/0' ? '#DIV/0!' : normalized;
  return isOfficeKernelSpreadsheetError(withSuffix);
}
