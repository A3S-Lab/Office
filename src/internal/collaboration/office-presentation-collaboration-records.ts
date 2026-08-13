import * as Y from 'yjs';
import type { WorkSlideElement } from '../features/work/work-types';
import {
  cloneWorkOfficeCollaborationJson as cloneJsonValue,
  workOfficeCollaborationJsonEqual as jsonEqual,
} from './office-collaboration-json';
import { invalidWorkOfficePresentationShared as invalidSharedPresentation } from './office-presentation-collaboration-validation';

export const PRESENTATION_RECORD_ELEMENTS = 'elements';
export const PRESENTATION_RECORD_ELEMENT_ORDER = 'elementOrder';
export const PRESENTATION_RECORD_COMMENTS = 'comments';
export const PRESENTATION_RECORD_COMMENT_ORDER = 'commentOrder';
export const PRESENTATION_RECORD_COMMENTS_PRESENT = 'commentsPresent';

export type WorkOfficePresentationRecord = {
  id: string;
  elements: WorkSlideElement[];
};

export function nestedPresentationMap(
  parent: Y.Map<unknown>,
  key: string,
  label: string,
): Y.Map<unknown> {
  const existing = parent.get(key);
  if (existing instanceof Y.Map) return existing;
  if (existing !== undefined) invalidSharedPresentation(label);
  const value = new Y.Map<unknown>();
  parent.set(key, value);
  return value;
}

export function nestedPresentationArray(
  parent: Y.Map<unknown>,
  key: string,
  label: string,
): Y.Array<string> {
  const existing = parent.get(key);
  if (existing instanceof Y.Array) return existing as Y.Array<string>;
  if (existing !== undefined) invalidSharedPresentation(label);
  const value = new Y.Array<string>();
  parent.set(key, value);
  return value;
}

export function requiredPresentationMap(
  parent: Y.Map<unknown>,
  key: string,
  label: string,
): Y.Map<unknown> {
  const value = parent.get(key);
  if (!(value instanceof Y.Map)) invalidSharedPresentation(label);
  return value as Y.Map<unknown>;
}

export function requiredPresentationArray(
  parent: Y.Map<unknown>,
  key: string,
  label: string,
): Y.Array<string> {
  const value = parent.get(key);
  if (!(value instanceof Y.Array)) invalidSharedPresentation(label);
  return value as Y.Array<string>;
}

export function validatedPresentationOrder(
  order: Y.Array<string>,
  records: Y.Map<unknown>,
  label: string,
): string[] {
  const seen = new Set<string>();
  const values: string[] = [];
  for (const value of order.toArray()) {
    if (typeof value !== 'string' || !value.trim()) {
      invalidSharedPresentation(`${label} order`);
    }
    // Concurrent idempotent inserts may duplicate only the order entry.
    if (seen.has(value)) continue;
    seen.add(value);
    values.push(value);
  }
  if (
    seen.size !== records.size ||
    Array.from(records.keys()).some((id) => !seen.has(id))
  ) {
    invalidSharedPresentation(`${label} order and record set`);
  }
  return values;
}

export function readPresentationCollection(
  records: Y.Map<unknown>,
  order: Y.Array<string>,
  label: string,
): WorkOfficePresentationRecord[] {
  return validatedPresentationOrder(order, records, label).map((id) => {
    const record = requiredPresentationMap(records, id, label);
    assertCommentRootsMatchRecordKind(record, label);
    const value: Record<string, unknown> = {};
    for (const [key, item] of record.entries()) {
      if (key === PRESENTATION_RECORD_COMMENTS_PRESENT) {
        if (label !== 'slide' || item !== true) {
          invalidSharedPresentation(`${label} comment presence`);
        }
        continue;
      }
      if (isNestedPresentationField(key)) continue;
      if (item instanceof Y.AbstractType) {
        invalidSharedPresentation(`${label} field '${key}'`);
      }
      value[key] = cloneJsonValue(item);
    }
    if (value.id !== id) invalidSharedPresentation(`${label} identity`);
    value.elements = readIdRecords(
      requiredPresentationMap(
        record,
        PRESENTATION_RECORD_ELEMENTS,
        `${label} elements`,
      ),
      requiredPresentationArray(
        record,
        PRESENTATION_RECORD_ELEMENT_ORDER,
        `${label} element order`,
      ),
      `${label} element`,
    );
    if (label === 'slide') readSlideComments(record, value);
    return value as WorkOfficePresentationRecord;
  });
}

export function patchPresentationElements(
  records: Y.Map<unknown>,
  order: Y.Array<string>,
  previous: WorkSlideElement[],
  next: WorkSlideElement[],
  parentLabel: string,
): void {
  patchPresentationIdRecords(
    records,
    order,
    previous,
    next,
    `${parentLabel} element`,
  );
}

export function patchPresentationIdRecords<T extends { id: string }>(
  records: Y.Map<unknown>,
  order: Y.Array<string>,
  previous: T[],
  next: T[],
  label: string,
): void {
  const previousById = new Map(previous.map((value) => [value.id, value]));
  const nextById = new Map(next.map((value) => [value.id, value]));
  for (const value of previous) {
    if (!nextById.has(value.id)) records.delete(value.id);
  }
  for (const value of next) {
    const before = previousById.get(value.id);
    if (before && jsonEqual(before, value)) continue;
    let record = records.get(value.id);
    if (record === undefined) {
      record = new Y.Map<unknown>();
      records.set(value.id, record);
    }
    if (!(record instanceof Y.Map)) invalidSharedPresentation(label);
    patchPresentationJsonMap(
      record as Y.Map<unknown>,
      (before ?? {}) as unknown as Record<string, unknown>,
      value as unknown as Record<string, unknown>,
    );
  }
  patchPresentationOrder(
    order,
    previous.map((value) => value.id),
    next.map((value) => value.id),
  );
}

export function patchPresentationOrder(
  order: Y.Array<string>,
  previous: string[],
  next: string[],
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
    const changedIds = next.filter((id) => previousIds.has(id));
    for (const id of changedIds) removeAll(order, id);
    for (const id of next) {
      if (!order.toArray().includes(id)) insertRelative(order, next, id);
    }
  }
}

function assertCommentRootsMatchRecordKind(
  record: Y.Map<unknown>,
  label: string,
): void {
  if (
    label !== 'slide' &&
    (record.has(PRESENTATION_RECORD_COMMENTS) ||
      record.has(PRESENTATION_RECORD_COMMENT_ORDER) ||
      record.has(PRESENTATION_RECORD_COMMENTS_PRESENT))
  ) {
    invalidSharedPresentation(`${label} comments`);
  }
}

function isNestedPresentationField(key: string): boolean {
  return (
    key === PRESENTATION_RECORD_ELEMENTS ||
    key === PRESENTATION_RECORD_ELEMENT_ORDER ||
    key === PRESENTATION_RECORD_COMMENTS ||
    key === PRESENTATION_RECORD_COMMENT_ORDER
  );
}

function readSlideComments(
  record: Y.Map<unknown>,
  value: Record<string, unknown>,
): void {
  const commentRecords = requiredPresentationMap(
    record,
    PRESENTATION_RECORD_COMMENTS,
    'slide comments',
  );
  const commentOrder = requiredPresentationArray(
    record,
    PRESENTATION_RECORD_COMMENT_ORDER,
    'slide comment order',
  );
  if (commentRecords.size > 0 || commentOrder.length > 0) {
    value.comments = readIdRecords(
      commentRecords,
      commentOrder,
      'slide comment',
    );
  } else if (record.get(PRESENTATION_RECORD_COMMENTS_PRESENT) === true) {
    value.comments = [];
  }
}

function readIdRecords(
  records: Y.Map<unknown>,
  order: Y.Array<string>,
  label: string,
): Record<string, unknown>[] {
  return validatedPresentationOrder(order, records, label).map((id) =>
    readJsonMap(requiredPresentationMap(records, id, label), id, label),
  );
}

function readJsonMap(
  record: Y.Map<unknown>,
  expectedId: string,
  label: string,
): Record<string, unknown> {
  const value: Record<string, unknown> = {};
  for (const [key, item] of record.entries()) {
    if (item instanceof Y.AbstractType) {
      invalidSharedPresentation(`${label} field '${key}'`);
    }
    value[key] = cloneJsonValue(item);
  }
  if (value.id !== expectedId) invalidSharedPresentation(`${label} identity`);
  return value;
}

function patchPresentationJsonMap(
  target: Y.Map<unknown>,
  previous: Record<string, unknown>,
  next: Record<string, unknown>,
): void {
  const keys = new Set([...Object.keys(previous), ...Object.keys(next)]);
  for (const key of keys) {
    if (jsonEqual(previous[key], next[key])) continue;
    if (next[key] === undefined) target.delete(key);
    else target.set(key, cloneJsonValue(next[key]));
  }
}

function insertRelative(
  order: Y.Array<string>,
  desired: string[],
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

function sameRelativeOrder(previous: string[], next: string[]): boolean {
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
