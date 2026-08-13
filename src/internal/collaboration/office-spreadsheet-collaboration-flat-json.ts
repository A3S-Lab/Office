import * as Y from 'yjs';
import { WorkOfficeCollaborationError } from './office-collaboration';
import {
  cloneWorkOfficeCollaborationJson as cloneJsonValue,
  isWorkOfficeCollaborationRecord as isRecord,
  workOfficeCollaborationJsonEqual as jsonEqual,
} from './office-collaboration-json';
import { invalidWorkOfficeSpreadsheetShared as invalidSharedSpreadsheet } from './office-spreadsheet-collaboration-validation';

type FlatJsonEntryKind = 'object' | 'value';

interface DecodedFlatJsonEntry {
  kind: FlatJsonEntryKind;
  path: string[];
  value: unknown;
}

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
  const before = flattenRecord(previous ?? {});
  const after = flattenRecord(next ?? {});
  const keys = new Set([...before.keys(), ...after.keys()]);
  for (const key of keys) {
    const beforeValue = before.get(key);
    const afterValue = after.get(key);
    const hadBefore = before.has(key);
    const hasAfter = after.has(key);
    if (hadBefore && hasAfter && jsonEqual(beforeValue, afterValue)) continue;
    if (!hasAfter) {
      if (hadBefore && !jsonEqual(target.get(key), beforeValue)) {
        throw new WorkOfficeCollaborationError(
          'office.collaboration.content_invalid',
          `The ${label} changed before one of its fields was removed. Refresh the shared snapshot before retrying.`,
        );
      }
      target.delete(key);
    } else {
      target.set(key, cloneJsonValue(afterValue));
    }
  }
}

export function readSpreadsheetFlatJsonMap(
  target: SpreadsheetFlatJsonMap,
  label: string,
): Record<string, unknown> {
  const entries: DecodedFlatJsonEntry[] = [];
  for (const [encoded, value] of target.entries()) {
    if (value instanceof Y.AbstractType) {
      invalidSharedSpreadsheet(`${label} field '${encoded}'`);
    }
    entries.push(decodedEntry(encoded, value, label));
  }
  entries.sort((left, right) => {
    const depth = left.path.length - right.path.length;
    if (depth !== 0) return depth;
    if (left.kind === right.kind) return 0;
    return left.kind === 'object' ? -1 : 1;
  });
  const result: Record<string, unknown> = {};
  const objectPaths = new Set<string>();
  for (const entry of entries) {
    const identity = JSON.stringify(entry.path);
    if (entry.kind === 'object') {
      if (entry.value !== true || entry.path.length === 0) {
        invalidSharedSpreadsheet(`${label} object marker`);
      }
      const existing = valueAtPath(result, entry.path);
      if (existing !== undefined && !isRecord(existing)) {
        invalidSharedSpreadsheet(`${label} object/value overlap`);
      }
      ensureObjectPath(result, entry.path, label);
      objectPaths.add(identity);
      continue;
    }
    if (entry.path.length === 0)
      invalidSharedSpreadsheet(`${label} root value`);
    setValueAtPath(result, entry.path, cloneJsonValue(entry.value), label);
  }
  for (const entry of entries) {
    for (let depth = 1; depth < entry.path.length; depth += 1) {
      if (!objectPaths.has(JSON.stringify(entry.path.slice(0, depth)))) {
        invalidSharedSpreadsheet(`${label} missing object marker`);
      }
    }
  }
  return result;
}

export interface SpreadsheetFlatJsonMap {
  delete(key: string): unknown;
  entries(): IterableIterator<[string, unknown]>;
  get(key: string): unknown;
  set(key: string, value: unknown): unknown;
}

function flattenRecord(value: Record<string, unknown>): Map<string, unknown> {
  const result = new Map<string, unknown>();
  for (const [key, child] of Object.entries(value)) {
    flattenValue(result, [key], child);
  }
  return result;
}

function flattenValue(
  result: Map<string, unknown>,
  path: string[],
  value: unknown,
): void {
  assertSafePath(path, 'Spreadsheet collaboration field');
  if (isRecord(value)) {
    result.set(encodedEntry('object', path), true);
    for (const [key, child] of Object.entries(value)) {
      flattenValue(result, [...path, key], child);
    }
  } else {
    result.set(encodedEntry('value', path), cloneJsonValue(value));
  }
}

function encodedEntry(kind: FlatJsonEntryKind, path: string[]): string {
  return JSON.stringify([kind, ...path]);
}

function decodedEntry(
  encoded: string,
  value: unknown,
  label: string,
): DecodedFlatJsonEntry {
  let identity: unknown;
  try {
    identity = JSON.parse(encoded);
  } catch {
    invalidSharedSpreadsheet(`${label} field identity`);
  }
  if (
    !Array.isArray(identity) ||
    identity.length < 2 ||
    (identity[0] !== 'object' && identity[0] !== 'value') ||
    identity.slice(1).some((part) => typeof part !== 'string') ||
    encodedEntry(
      identity[0] as FlatJsonEntryKind,
      identity.slice(1) as string[],
    ) !== encoded
  ) {
    invalidSharedSpreadsheet(`${label} field identity`);
  }
  assertSafePath(identity.slice(1) as string[], label);
  return {
    kind: identity[0] as FlatJsonEntryKind,
    path: identity.slice(1) as string[],
    value,
  };
}

function assertSafePath(path: string[], label: string): void {
  if (
    path.some(
      (part) =>
        !part ||
        part === '__proto__' ||
        part === 'constructor' ||
        part === 'prototype',
    )
  ) {
    invalidSharedSpreadsheet(`${label} field path`);
  }
}

function valueAtPath(root: Record<string, unknown>, path: string[]): unknown {
  let current: unknown = root;
  for (const part of path) {
    if (!isRecord(current) || !Object.hasOwn(current, part)) return undefined;
    current = current[part];
  }
  return current;
}

function ensureObjectPath(
  root: Record<string, unknown>,
  path: string[],
  label: string,
): void {
  let current = root;
  for (const part of path) {
    const existing = current[part];
    if (existing === undefined) {
      defineJsonProperty(current, part, {});
      current = current[part] as Record<string, unknown>;
    } else if (isRecord(existing)) {
      current = existing;
    } else {
      invalidSharedSpreadsheet(`${label} object/value overlap`);
    }
  }
}

function setValueAtPath(
  root: Record<string, unknown>,
  path: string[],
  value: unknown,
  label: string,
): void {
  let current = root;
  for (const part of path.slice(0, -1)) {
    const existing = current[part];
    if (!isRecord(existing)) {
      invalidSharedSpreadsheet(`${label} value without object parent`);
    }
    current = existing;
  }
  const key = path.at(-1) as string;
  if (Object.hasOwn(current, key)) {
    invalidSharedSpreadsheet(`${label} duplicate object/value field`);
  }
  defineJsonProperty(current, key, value);
}

function defineJsonProperty(
  target: Record<string, unknown>,
  key: string,
  value: unknown,
): void {
  Object.defineProperty(target, key, {
    configurable: true,
    enumerable: true,
    value,
    writable: true,
  });
}
