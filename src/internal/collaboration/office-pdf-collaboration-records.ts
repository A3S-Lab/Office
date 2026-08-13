import type * as Y from 'yjs';
import {
  patchWorkOfficeCollaborationFlatJsonMap as patchFlatJsonMap,
  readWorkOfficeCollaborationFlatJsonMap as readFlatJsonMap,
  type WorkOfficeCollaborationFlatJsonMap,
} from './office-collaboration-flat-json';
import { workOfficeCollaborationJsonEqual as jsonEqual } from './office-collaboration-json';
import { invalidWorkOfficePdfShared as invalidSharedPdf } from './office-pdf-collaboration-validation';

export interface WorkOfficePdfRecordCollectionRoots {
  presence: Y.Map<unknown>;
  fields: Y.Map<unknown>;
  order: Y.Array<string>;
}

export function workOfficePdfRecordCollectionRoots(
  document: Y.Doc,
  rootName: (suffix: string) => string,
  name: string,
): WorkOfficePdfRecordCollectionRoots {
  return {
    presence: document.getMap(rootName(`pdf.${name}.presence`)),
    fields: document.getMap(rootName(`pdf.${name}.fields`)),
    order: document.getArray(rootName(`pdf.${name}.order`)),
  };
}

export function patchWorkOfficePdfRecords<T extends { id: string }>(
  roots: WorkOfficePdfRecordCollectionRoots,
  previous: readonly T[],
  next: readonly T[],
  label: string,
): void {
  const beforeById = new Map(previous.map((value) => [value.id, value]));
  const afterById = new Map(next.map((value) => [value.id, value]));
  for (const before of previous) {
    if (afterById.has(before.id)) continue;
    const presence = roots.presence.get(before.id);
    if (presence === undefined) {
      removeAll(roots.order, before.id);
      continue;
    }
    if (presence !== true) invalidSharedPdf(`${label} record presence`);
    roots.presence.delete(before.id);
    patchFlatJsonMap(
      new PdfRecordFlatJsonMap(roots.fields, before.id, label),
      before as Record<string, unknown>,
      undefined,
      `${label} '${before.id}'`,
    );
    removeAll(roots.order, before.id);
  }
  for (const after of next) {
    const before = beforeById.get(after.id);
    if (before && jsonEqual(before, after)) continue;
    roots.presence.set(after.id, true);
    patchFlatJsonMap(
      new PdfRecordFlatJsonMap(roots.fields, after.id, label),
      before as Record<string, unknown> | undefined,
      after as Record<string, unknown>,
      `${label} '${after.id}'`,
    );
  }
  patchWorkOfficePdfOrder(
    roots.order,
    previous.map(({ id }) => id),
    next.map(({ id }) => id),
  );
}

export function readWorkOfficePdfRecords<
  T extends Record<string, unknown> & { id: string },
>(roots: WorkOfficePdfRecordCollectionRoots, label: string): T[] {
  const ids = validatedPdfOrder(roots.order, roots.presence, label);
  const fieldsById = pdfFieldsById(roots, label);
  return ids.map((id) => {
    const value = readFlatJsonMap(
      fieldsById.get(id) ?? new Map<string, unknown>(),
      `${label} '${id}'`,
      invalidSharedPdf,
    ) as T;
    if (value.id !== id) invalidSharedPdf(`${label} identity`);
    return value;
  });
}

export function assertWorkOfficePdfRecordCollectionEmpty(
  roots: WorkOfficePdfRecordCollectionRoots,
): boolean {
  return (
    roots.presence.size === 0 &&
    roots.fields.size === 0 &&
    roots.order.length === 0
  );
}

export function workOfficePdfRecordCollectionUndoScope(
  roots: WorkOfficePdfRecordCollectionRoots,
): Array<Y.Map<unknown> | Y.Array<string>> {
  return [roots.presence, roots.fields, roots.order];
}

/**
 * Reports whether a transaction changed one recursively addressed record
 * field. This lets format bindings keep an otherwise mutable collection on
 * the local undo stack while treating selected fields as irreversible.
 */
export function workOfficePdfRecordFieldChanged(
  transaction: Y.Transaction,
  roots: WorkOfficePdfRecordCollectionRoots,
  path: readonly string[],
): boolean {
  const changedKeys = transaction.changed.get(
    roots.fields as unknown as Y.AbstractType<
      Y.YEvent<Y.AbstractType<unknown>>
    >,
  );
  if (!changedKeys) return false;
  const encodedField = JSON.stringify(['value', ...path]);
  for (const key of changedKeys) {
    if (key === null) return true;
    const identity = tryDecodeRecordFieldIdentity(key);
    if (identity?.[1] === encodedField) return true;
  }
  return false;
}

function validatedPdfOrder(
  order: Y.Array<string>,
  presence: Y.Map<unknown>,
  label: string,
): string[] {
  const seen = new Set<string>();
  const result: string[] = [];
  for (const id of order.toArray()) {
    if (typeof id !== 'string' || !id.trim())
      invalidSharedPdf(`${label} order`);
    if (seen.has(id)) continue;
    seen.add(id);
    result.push(id);
  }
  for (const [id, value] of presence.entries()) {
    if (value !== true || !id.trim() || !seen.has(id)) {
      invalidSharedPdf(`${label} order and record set`);
    }
  }
  if (seen.size !== presence.size) {
    invalidSharedPdf(`${label} order and record set`);
  }
  return result;
}

function pdfFieldsById(
  roots: WorkOfficePdfRecordCollectionRoots,
  label: string,
): Map<string, Map<string, unknown>> {
  const result = new Map<string, Map<string, unknown>>();
  for (const [encoded, value] of roots.fields.entries()) {
    const [id, field] = decodedRecordFieldIdentity(encoded, label);
    if (roots.presence.get(id) !== true) {
      invalidSharedPdf(`${label} orphan field`);
    }
    let fields = result.get(id);
    if (!fields) {
      fields = new Map<string, unknown>();
      result.set(id, fields);
    }
    fields.set(field, value);
  }
  return result;
}

function patchWorkOfficePdfOrder(
  order: Y.Array<string>,
  previous: readonly string[],
  next: readonly string[],
): void {
  if (jsonEqual(previous, next)) return;
  const nextIds = new Set(next);
  for (const id of previous) {
    if (!nextIds.has(id)) removeAll(order, id);
  }
  const previousIds = new Set(previous);
  for (const id of next) {
    if (previousIds.has(id) || order.toArray().includes(id)) continue;
    insertRelative(order, next, id);
  }
  if (!sameRelativeOrder(previous, next)) {
    for (const id of next.filter((id) => previousIds.has(id)))
      removeAll(order, id);
    for (const id of next) {
      if (!order.toArray().includes(id)) insertRelative(order, next, id);
    }
  }
}

class PdfRecordFlatJsonMap implements WorkOfficeCollaborationFlatJsonMap {
  constructor(
    private readonly fields: Y.Map<unknown>,
    private readonly id: string,
    private readonly label: string,
  ) {}

  delete(key: string): boolean {
    const encoded = encodedRecordFieldIdentity(this.id, key);
    const hadValue = this.fields.has(encoded);
    this.fields.delete(encoded);
    return hadValue;
  }

  *entries(): IterableIterator<[string, unknown]> {
    for (const [encoded, value] of this.fields.entries()) {
      const [id, field] = decodedRecordFieldIdentity(encoded, this.label);
      if (id === this.id) yield [field, value];
    }
  }

  get(key: string): unknown {
    return this.fields.get(encodedRecordFieldIdentity(this.id, key));
  }

  set(key: string, value: unknown): Y.Map<unknown> {
    this.fields.set(encodedRecordFieldIdentity(this.id, key), value);
    return this.fields;
  }
}

function encodedRecordFieldIdentity(id: string, field: string): string {
  return JSON.stringify([id, field]);
}

function decodedRecordFieldIdentity(
  encoded: string,
  label: string,
): [string, string] {
  const identity = tryDecodeRecordFieldIdentity(encoded);
  if (!identity?.[0].trim()) {
    invalidSharedPdf(`${label} field identity`);
  }
  if (encodedRecordFieldIdentity(identity[0], identity[1]) !== encoded) {
    invalidSharedPdf(`${label} field identity`);
  }
  return identity;
}

function tryDecodeRecordFieldIdentity(
  encoded: string,
): [string, string] | undefined {
  let identity: unknown;
  try {
    identity = JSON.parse(encoded);
  } catch {
    return undefined;
  }
  if (
    !Array.isArray(identity) ||
    identity.length !== 2 ||
    identity.some((value) => typeof value !== 'string')
  ) {
    return undefined;
  }
  return identity as [string, string];
}

function insertRelative(
  order: Y.Array<string>,
  desired: readonly string[],
  id: string,
): void {
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

function sameRelativeOrder(
  previous: readonly string[],
  next: readonly string[],
): boolean {
  const nextIds = new Set(next);
  const previousIds = new Set(previous);
  return jsonEqual(
    previous.filter((id) => nextIds.has(id)),
    next.filter((id) => previousIds.has(id)),
  );
}

function removeAll(order: Y.Array<string>, id: string): void {
  for (let index = order.length - 1; index >= 0; index -= 1) {
    if (order.get(index) === id) order.delete(index, 1);
  }
}
