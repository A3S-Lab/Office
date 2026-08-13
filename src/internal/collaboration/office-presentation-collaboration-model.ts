import * as Y from 'yjs';
import type {
  WorkPresentationContent,
  WorkPresentationLayout,
  WorkPresentationMaster,
  WorkSlide,
} from '../features/work/work-types';
import { WorkOfficeCollaborationError } from './office-collaboration';
import {
  cloneWorkOfficeCollaborationJson as cloneJsonValue,
  workOfficeCollaborationJsonEqual as jsonEqual,
} from './office-collaboration-json';
import {
  invalidWorkOfficePresentationShared as invalidSharedPresentation,
  validateSharedWorkOfficePresentationContent,
} from './office-presentation-collaboration-validation';
import {
  nestedPresentationArray as nestedArray,
  nestedPresentationMap as nestedMap,
  patchPresentationElements as patchElements,
  patchPresentationIdRecords as patchIdRecords,
  patchPresentationOrder as patchOrder,
  PRESENTATION_RECORD_COMMENT_ORDER as RECORD_COMMENT_ORDER,
  PRESENTATION_RECORD_COMMENTS as RECORD_COMMENTS,
  PRESENTATION_RECORD_COMMENTS_PRESENT as RECORD_COMMENTS_PRESENT,
  PRESENTATION_RECORD_ELEMENT_ORDER as RECORD_ELEMENT_ORDER,
  PRESENTATION_RECORD_ELEMENTS as RECORD_ELEMENTS,
  readPresentationCollection as readCollection,
  type WorkOfficePresentationRecord as PresentationRecord,
} from './office-presentation-collaboration-records';

export { validateWorkOfficePresentationContent } from './office-presentation-collaboration-validation';

export const PRESENTATION_OPTIONS_ROOT = 'presentation.options';
export const PRESENTATION_SLIDES_ROOT = 'presentation.slides';
export const PRESENTATION_SLIDE_ORDER_ROOT = 'presentation.slide-order';
export const PRESENTATION_MASTERS_ROOT = 'presentation.masters';
export const PRESENTATION_MASTER_ORDER_ROOT = 'presentation.master-order';
export const PRESENTATION_LAYOUTS_ROOT = 'presentation.layouts';
export const PRESENTATION_LAYOUT_ORDER_ROOT = 'presentation.layout-order';

export interface WorkOfficePresentationRoots {
  options: Y.Map<unknown>;
  slides: Y.Map<unknown>;
  slideOrder: Y.Array<string>;
  masters: Y.Map<unknown>;
  masterOrder: Y.Array<string>;
  layouts: Y.Map<unknown>;
  layoutOrder: Y.Array<string>;
}

interface PresentationCollection<T extends PresentationRecord> {
  label: string;
  order: Y.Array<string>;
  previous: T[];
  records: Y.Map<unknown>;
  next: T[];
}

export function workOfficePresentationRoots(
  document: Y.Doc,
  rootName: (suffix: string) => string,
): WorkOfficePresentationRoots {
  return {
    options: document.getMap(rootName(PRESENTATION_OPTIONS_ROOT)),
    slides: document.getMap(rootName(PRESENTATION_SLIDES_ROOT)),
    slideOrder: document.getArray(rootName(PRESENTATION_SLIDE_ORDER_ROOT)),
    masters: document.getMap(rootName(PRESENTATION_MASTERS_ROOT)),
    masterOrder: document.getArray(rootName(PRESENTATION_MASTER_ORDER_ROOT)),
    layouts: document.getMap(rootName(PRESENTATION_LAYOUTS_ROOT)),
    layoutOrder: document.getArray(rootName(PRESENTATION_LAYOUT_ORDER_ROOT)),
  };
}

export function initializeWorkOfficePresentationRoots(
  roots: WorkOfficePresentationRoots,
  content: WorkPresentationContent,
): void {
  patchOptionalJson(roots.options, 'width', undefined, content.width);
  patchOptionalJson(roots.options, 'height', undefined, content.height);
  patchCollection({
    label: 'slide',
    order: roots.slideOrder,
    previous: [],
    records: roots.slides,
    next: content.slides,
  });
  patchCollection({
    label: 'presentation master',
    order: roots.masterOrder,
    previous: [],
    records: roots.masters,
    next: content.masters ?? [],
  });
  patchCollection({
    label: 'presentation layout',
    order: roots.layoutOrder,
    previous: [],
    records: roots.layouts,
    next: content.layouts ?? [],
  });
}

export function readWorkOfficePresentationRoots(
  roots: WorkOfficePresentationRoots,
): WorkPresentationContent {
  const width = optionalSharedNumber(
    roots.options.get('width'),
    'presentation width',
  );
  const height = optionalSharedNumber(
    roots.options.get('height'),
    'presentation height',
  );
  const result: WorkPresentationContent = {
    type: 'presentation',
    slides: readCollection(
      roots.slides,
      roots.slideOrder,
      'slide',
    ) as unknown as WorkSlide[],
  };
  if (width !== undefined) result.width = width;
  if (height !== undefined) result.height = height;
  const masters = readCollection(
    roots.masters,
    roots.masterOrder,
    'presentation master',
  ) as unknown as WorkPresentationMaster[];
  const layouts = readCollection(
    roots.layouts,
    roots.layoutOrder,
    'presentation layout',
  ) as unknown as WorkPresentationLayout[];
  if (masters.length > 0) result.masters = masters;
  if (layouts.length > 0) result.layouts = layouts;
  return validateSharedWorkOfficePresentationContent(result);
}

export function patchWorkOfficePresentationRoots(
  roots: WorkOfficePresentationRoots,
  previous: WorkPresentationContent,
  next: WorkPresentationContent,
): void {
  const shared = readWorkOfficePresentationRoots(roots);
  assertChangedRecordsStillExist(previous, next, shared);
  patchOptionalJson(roots.options, 'width', previous.width, next.width);
  patchOptionalJson(roots.options, 'height', previous.height, next.height);
  patchCollection({
    label: 'slide',
    order: roots.slideOrder,
    previous: previous.slides,
    records: roots.slides,
    next: next.slides,
  });
  patchCollection({
    label: 'presentation master',
    order: roots.masterOrder,
    previous: previous.masters ?? [],
    records: roots.masters,
    next: next.masters ?? [],
  });
  patchCollection({
    label: 'presentation layout',
    order: roots.layoutOrder,
    previous: previous.layouts ?? [],
    records: roots.layouts,
    next: next.layouts ?? [],
  });
}

function assertChangedRecordsStillExist(
  previous: WorkPresentationContent,
  next: WorkPresentationContent,
  shared: WorkPresentationContent,
): void {
  assertChangedCollectionExists(
    previous.slides,
    next.slides,
    shared.slides,
    'slide',
  );
  assertChangedCollectionExists(
    previous.masters ?? [],
    next.masters ?? [],
    shared.masters ?? [],
    'presentation master',
  );
  assertChangedCollectionExists(
    previous.layouts ?? [],
    next.layouts ?? [],
    shared.layouts ?? [],
    'presentation layout',
  );
}

function assertChangedCollectionExists<T extends PresentationRecord>(
  previous: T[],
  next: T[],
  shared: T[],
  label: string,
): void {
  const previousById = new Map(previous.map((value) => [value.id, value]));
  const sharedById = new Map(shared.map((value) => [value.id, value]));
  for (const value of next) {
    const before = previousById.get(value.id);
    const current = sharedById.get(value.id);
    if (before && !jsonEqual(before, value) && !current) {
      throw new WorkOfficeCollaborationError(
        'office.collaboration.content_invalid',
        `The ${label} '${value.id}' was removed before this change could be applied. Refresh the shared snapshot before retrying.`,
      );
    }
    if (!before || !current) continue;
    assertChangedChildRecordsExist(
      before.elements,
      value.elements,
      current.elements,
      `element in ${label} '${value.id}'`,
    );
    if (label === 'slide') {
      assertChangedChildRecordsExist(
        (before as unknown as WorkSlide).comments ?? [],
        (value as unknown as WorkSlide).comments ?? [],
        (current as unknown as WorkSlide).comments ?? [],
        `comment in slide '${value.id}'`,
      );
    }
  }
}

function assertChangedChildRecordsExist<T extends { id: string }>(
  previous: T[],
  next: T[],
  shared: T[],
  label: string,
): void {
  const previousById = new Map(previous.map((value) => [value.id, value]));
  const sharedIds = new Set(shared.map((value) => value.id));
  for (const value of next) {
    const before = previousById.get(value.id);
    if (before && !jsonEqual(before, value) && !sharedIds.has(value.id)) {
      throw new WorkOfficeCollaborationError(
        'office.collaboration.content_invalid',
        `The ${label} '${value.id}' was removed before this change could be applied. Refresh the shared snapshot before retrying.`,
      );
    }
  }
}

export function assertWorkOfficePresentationRootsEmpty(
  roots: WorkOfficePresentationRoots,
): void {
  if (
    roots.options.size > 0 ||
    roots.slides.size > 0 ||
    roots.slideOrder.length > 0 ||
    roots.masters.size > 0 ||
    roots.masterOrder.length > 0 ||
    roots.layouts.size > 0 ||
    roots.layoutOrder.length > 0
  ) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.bootstrap_ambiguous',
      'The Presentation collaboration roots contain data without initialized metadata.',
    );
  }
}

export function workOfficePresentationUndoScope(
  roots: WorkOfficePresentationRoots,
): Array<Y.Map<unknown> | Y.Array<string>> {
  return [
    roots.options,
    roots.slides,
    roots.slideOrder,
    roots.masters,
    roots.masterOrder,
    roots.layouts,
    roots.layoutOrder,
  ];
}

function patchCollection<T extends PresentationRecord>({
  label,
  order,
  previous,
  records,
  next,
}: PresentationCollection<T>): void {
  const previousById = new Map(previous.map((value) => [value.id, value]));
  const nextById = new Map(next.map((value) => [value.id, value]));
  for (const value of previous) {
    if (nextById.has(value.id)) continue;
    records.delete(value.id);
    removeAll(order, value.id);
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
    patchRecord(record as Y.Map<unknown>, before, value, label);
  }
  patchOrder(
    order,
    previous.map((value) => value.id),
    next.map((value) => value.id),
  );
}

function patchRecord<T extends PresentationRecord>(
  record: Y.Map<unknown>,
  previous: T | undefined,
  next: T,
  label: string,
): void {
  const elements = nestedMap(record, RECORD_ELEMENTS, `${label} elements`);
  const elementOrder = nestedArray(
    record,
    RECORD_ELEMENT_ORDER,
    `${label} element order`,
  );
  const nextRecord = next as unknown as Record<string, unknown>;
  const previousRecord = previous as unknown as
    | Record<string, unknown>
    | undefined;
  const keys = new Set([
    ...Object.keys(previousRecord ?? {}),
    ...Object.keys(nextRecord),
  ]);
  keys.delete('elements');
  keys.delete('comments');
  for (const key of keys) {
    if (previous && jsonEqual(previousRecord?.[key], nextRecord[key])) continue;
    if (nextRecord[key] === undefined) record.delete(key);
    else record.set(key, cloneJsonValue(nextRecord[key]));
  }
  const comments =
    label === 'slide'
      ? nestedMap(record, RECORD_COMMENTS, 'slide comments')
      : undefined;
  const commentOrder =
    label === 'slide'
      ? nestedArray(record, RECORD_COMMENT_ORDER, 'slide comment order')
      : undefined;
  patchElements(
    elements,
    elementOrder,
    previous?.elements ?? [],
    next.elements,
    label,
  );
  if (comments && commentOrder) {
    const previousComments = previousRecord?.comments as
      | Array<{ id: string }>
      | undefined;
    const nextComments = nextRecord.comments as
      | Array<{ id: string }>
      | undefined;
    if (previousComments !== undefined || nextComments !== undefined) {
      if (nextComments !== undefined) record.set(RECORD_COMMENTS_PRESENT, true);
      else record.delete(RECORD_COMMENTS_PRESENT);
      patchIdRecords(
        comments,
        commentOrder,
        previousComments ?? [],
        nextComments ?? [],
        'slide comment',
      );
    }
  }
}

function optionalSharedNumber(
  value: unknown,
  label: string,
): number | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    invalidSharedPresentation(label);
  }
  return value as number;
}

function removeAll(order: Y.Array<string>, id: string): void {
  for (let index = order.length - 1; index >= 0; index -= 1) {
    if (order.get(index) === id) order.delete(index, 1);
  }
}

function patchOptionalJson(
  target: Y.Map<unknown>,
  key: string,
  previous: unknown,
  next: unknown,
): void {
  if (jsonEqual(previous, next)) return;
  if (next === undefined) target.delete(key);
  else target.set(key, cloneJsonValue(next));
}
