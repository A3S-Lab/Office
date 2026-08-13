import * as Y from 'yjs';
import { WorkOfficeCollaborationError } from './office-collaboration';

export function validatedOrder(
  order: Y.Array<string>,
  records: Y.Map<unknown>,
  label: string,
): string[] {
  const values = order.toArray();
  const seen = new Set<string>();
  const uniqueValues: string[] = [];
  for (const value of values) {
    if (typeof value !== 'string' || !value.trim()) {
      invalidSharedSidecars(`${label} order`);
    }
    // Two disconnected clients can retry the same stable-ID insert. Y.Array
    // preserves both entries after synchronization, while the record map has
    // one idempotent value. Collapse only the repeated order entry on read.
    if (seen.has(value)) continue;
    seen.add(value);
    uniqueValues.push(value);
  }
  if (
    seen.size !== records.size ||
    Array.from(records.keys()).some((id) => !seen.has(id))
  ) {
    invalidSharedSidecars(`${label} order and record set`);
  }
  return uniqueValues;
}

export function requiredSharedMap(
  parent: Y.Map<unknown>,
  key: string,
  label: string,
): Y.Map<unknown> {
  const value = parent.get(key);
  if (!(value instanceof Y.Map)) invalidSharedSidecars(label);
  return value as Y.Map<unknown>;
}

export function requiredNestedMap(
  parent: Y.Map<unknown>,
  key: string,
  label: string,
): Y.Map<unknown> {
  return requiredSharedMap(parent, key, label);
}

export function requiredNestedArray(
  parent: Y.Map<unknown>,
  key: string,
  label: string,
): Y.Array<string> {
  const value = parent.get(key);
  if (!(value instanceof Y.Array)) invalidSharedSidecars(label);
  return value as Y.Array<string>;
}

export function patchRequiredScalar(
  target: Y.Map<unknown>,
  key: string,
  previous: unknown,
  next: unknown,
): void {
  if (!jsonEqual(previous, next)) target.set(key, cloneJsonValue(next));
}

export function patchOptionalScalar(
  target: Y.Map<unknown>,
  key: string,
  previous: unknown,
  next: unknown,
): void {
  if (jsonEqual(previous, next)) return;
  if (next === undefined) target.delete(key);
  else target.set(key, cloneJsonValue(next));
}

export function removeFromOrder(order: Y.Array<string>, id: string): void {
  for (let index = order.length - 1; index >= 0; index -= 1) {
    if (order.get(index) === id) order.delete(index, 1);
  }
}

export function insertIntoOrder(
  order: Y.Array<string>,
  desired: string[],
  id: string,
): void {
  if (order.toArray().includes(id)) return;
  const desiredIndex = desired.indexOf(id);
  for (let index = desiredIndex - 1; index >= 0; index -= 1) {
    const currentIndex = order.toArray().indexOf(desired[index]);
    if (currentIndex >= 0) {
      order.insert(currentIndex + 1, [id]);
      return;
    }
  }
  for (let index = desiredIndex + 1; index < desired.length; index += 1) {
    const currentIndex = order.toArray().indexOf(desired[index]);
    if (currentIndex >= 0) {
      order.insert(currentIndex, [id]);
      return;
    }
  }
  order.push([id]);
}

export function optionalSharedString(
  value: unknown,
  label: string,
): string | undefined {
  if (value === undefined) return undefined;
  return requiredSharedString(value, label);
}

export function requiredIdentifier(value: unknown, label: string): string {
  const result = requiredString(value, `${label} ID`);
  if (result !== result.trim() || result.length > 256) {
    invalidInputSidecars(`a ${label} ID containing 1 to 256 characters`);
  }
  return result;
}

export function requiredSharedIdentifier(
  value: unknown,
  label: string,
): string {
  const result = requiredSharedString(value, `${label} ID`);
  if (result !== result.trim() || result.length > 256) {
    invalidSharedSidecars(`${label} ID`);
  }
  return result;
}

export function requiredString(value: unknown, label: string): string {
  if (typeof value !== 'string') invalidInputSidecars(`a string ${label}`);
  return value as string;
}

export function requiredSharedString(value: unknown, label: string): string {
  if (typeof value !== 'string') invalidSharedSidecars(label);
  return value as string;
}

export function requiredBoolean(value: unknown, label: string): boolean {
  if (typeof value !== 'boolean') invalidInputSidecars(`a boolean ${label}`);
  return value as boolean;
}

export function requiredSharedBoolean(value: unknown, label: string): boolean {
  if (typeof value !== 'boolean') invalidSharedSidecars(label);
  return value as boolean;
}

export function cloneJsonRecord(
  value: Record<string, unknown>,
): Record<string, unknown> {
  return cloneJsonValue(value) as Record<string, unknown>;
}

export function cloneJsonValue(
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
    if (seen.has(value)) invalidInputSidecars('acyclic JSON sidecar values');
    seen.add(value);
    const clone = value.map((item) => cloneJsonValue(item, seen));
    seen.delete(value);
    return clone;
  }
  if (isRecord(value)) {
    if (!isPlainJsonRecord(value)) {
      invalidInputSidecars('plain JSON sidecar objects');
    }
    if (seen.has(value)) invalidInputSidecars('acyclic JSON sidecar values');
    seen.add(value);
    const clone: Record<string, unknown> = {};
    for (const [key, item] of Object.entries(value)) {
      if (isUnsafeJsonKey(key)) {
        invalidInputSidecars('sidecar objects without prototype keys');
      }
      if (item !== undefined) {
        Object.defineProperty(clone, key, {
          configurable: true,
          enumerable: true,
          value: cloneJsonValue(item, seen),
          writable: true,
        });
      }
    }
    seen.delete(value);
    return clone;
  }
  invalidInputSidecars('JSON-compatible sidecar values');
}

export function jsonEqual(left: unknown, right: unknown): boolean {
  if (Object.is(left, right)) return true;
  if (Array.isArray(left) && Array.isArray(right)) {
    return (
      left.length === right.length &&
      left.every((value, index) => jsonEqual(value, right[index]))
    );
  }
  if (isRecord(left) && isRecord(right)) {
    if (!isPlainJsonRecord(left) || !isPlainJsonRecord(right)) return false;
    const leftKeys = Object.keys(left).filter((key) => left[key] !== undefined);
    const rightKeys = Object.keys(right).filter(
      (key) => right[key] !== undefined,
    );
    return (
      leftKeys.length === rightKeys.length &&
      leftKeys.every(
        (key) => Object.hasOwn(right, key) && jsonEqual(left[key], right[key]),
      )
    );
  }
  return false;
}

export function canonicalJson(value: unknown): string {
  return JSON.stringify(canonicalJsonValue(cloneJsonValue(value)));
}

export function assertNoAddedCollision<T extends { id: string }>(
  previous: T[],
  next: T[],
  shared: T[],
  label: string,
): void {
  const beforeIds = new Set(previous.map((value) => value.id));
  const sharedById = new Map(shared.map((value) => [value.id, value]));
  for (const value of next) {
    if (beforeIds.has(value.id)) continue;
    const existing = sharedById.get(value.id);
    if (existing && !jsonEqual(existing, value)) {
      throw new WorkOfficeCollaborationError(
        'office.collaboration.content_invalid',
        `The ${label} ID '${value.id}' was concurrently assigned to different records.`,
      );
    }
  }
}

export function changedRecordMissing(label: string, id: string): never {
  throw new WorkOfficeCollaborationError(
    'office.collaboration.content_invalid',
    `The ${label} '${id}' was removed before this change could be applied. Refresh the shared snapshot before retrying.`,
  );
}

export function invalidInputSidecars(expected: string): never {
  throw new WorkOfficeCollaborationError(
    'office.collaboration.content_invalid',
    `Document collaboration requires ${expected}.`,
  );
}

export function invalidSharedSidecars(label: string): never {
  throw new WorkOfficeCollaborationError(
    'office.collaboration.content_invalid',
    `The shared Document collaboration ${label} is invalid.`,
  );
}

export function isRecord(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function isPlainJsonRecord(value: Record<string, unknown>): boolean {
  const prototype = Object.getPrototypeOf(value);
  return prototype === Object.prototype || prototype === null;
}

function canonicalJsonValue(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(canonicalJsonValue);
  if (!isRecord(value)) return value;
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

function isUnsafeJsonKey(key: string): boolean {
  return key === '__proto__' || key === 'constructor' || key === 'prototype';
}
