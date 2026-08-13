import { WorkOfficeCollaborationError } from './office-collaboration';

export function cloneWorkOfficeCollaborationJson(
  value: unknown,
  seen = new WeakSet<object>(),
): unknown {
  if (
    value === null ||
    typeof value === 'string' ||
    typeof value === 'boolean'
  ) {
    return value;
  }
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (Array.isArray(value)) {
    if (seen.has(value)) invalidCollaborationJson('acyclic JSON values');
    seen.add(value);
    const clone = value.map((item) =>
      cloneWorkOfficeCollaborationJson(item, seen),
    );
    seen.delete(value);
    return clone;
  }
  if (isWorkOfficeCollaborationRecord(value)) {
    if (!isPlainJsonRecord(value)) {
      invalidCollaborationJson('plain JSON objects');
    }
    if (seen.has(value)) invalidCollaborationJson('acyclic JSON values');
    seen.add(value);
    const clone: Record<string, unknown> = {};
    for (const [key, item] of Object.entries(value)) {
      if (isUnsafeJsonKey(key)) {
        invalidCollaborationJson('objects without prototype keys');
      }
      if (item !== undefined) {
        Object.defineProperty(clone, key, {
          configurable: true,
          enumerable: true,
          value: cloneWorkOfficeCollaborationJson(item, seen),
          writable: true,
        });
      }
    }
    seen.delete(value);
    return clone;
  }
  invalidCollaborationJson('JSON-compatible values');
}

export function workOfficeCollaborationJsonEqual(
  left: unknown,
  right: unknown,
): boolean {
  if (Object.is(left, right)) return true;
  if (Array.isArray(left) && Array.isArray(right)) {
    return (
      left.length === right.length &&
      left.every((value, index) =>
        workOfficeCollaborationJsonEqual(value, right[index]),
      )
    );
  }
  if (
    isWorkOfficeCollaborationRecord(left) &&
    isWorkOfficeCollaborationRecord(right)
  ) {
    if (!isPlainJsonRecord(left) || !isPlainJsonRecord(right)) return false;
    const leftKeys = Object.keys(left).filter((key) => left[key] !== undefined);
    const rightKeys = Object.keys(right).filter(
      (key) => right[key] !== undefined,
    );
    return (
      leftKeys.length === rightKeys.length &&
      leftKeys.every(
        (key) =>
          Object.hasOwn(right, key) &&
          workOfficeCollaborationJsonEqual(left[key], right[key]),
      )
    );
  }
  return false;
}

export function canonicalWorkOfficeCollaborationJson(value: unknown): string {
  return JSON.stringify(
    canonicalJsonValue(cloneWorkOfficeCollaborationJson(value)),
  );
}

export function isWorkOfficeCollaborationRecord(
  value: unknown,
): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function canonicalJsonValue(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(canonicalJsonValue);
  if (!isWorkOfficeCollaborationRecord(value)) return value;
  const result: Record<string, unknown> = {};
  for (const key of Object.keys(value).sort()) {
    Object.defineProperty(result, key, {
      configurable: true,
      enumerable: true,
      value: canonicalJsonValue(value[key]),
      writable: true,
    });
  }
  return result;
}

function isPlainJsonRecord(value: Record<string, unknown>): boolean {
  const prototype = Object.getPrototypeOf(value);
  return prototype === Object.prototype || prototype === null;
}

function isUnsafeJsonKey(key: string): boolean {
  return key === '__proto__' || key === 'constructor' || key === 'prototype';
}

function invalidCollaborationJson(expected: string): never {
  throw new WorkOfficeCollaborationError(
    'office.collaboration.content_invalid',
    `Office collaboration requires ${expected}.`,
  );
}
