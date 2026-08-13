import { WorkOfficeCollaborationError } from './office-collaboration';
import {
  isWorkOfficeCollaborationRecord as isRecord,
  workOfficeCollaborationJsonEqual as jsonEqual,
} from './office-collaboration-json';
import type {
  WorkPdfCollaborationAnnotation,
  WorkPdfCollaborationContent,
} from './office-pdf-collaboration-types';

export function assertWorkOfficePdfPatchSafe(
  previous: WorkPdfCollaborationContent,
  next: WorkPdfCollaborationContent,
  shared: WorkPdfCollaborationContent,
): void {
  if (!jsonEqual(previous.source, next.source)) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.content_invalid',
      'The immutable PDF source identity cannot be changed by a collaboration update.',
    );
  }
  if (!jsonEqual(previous.source, shared.source)) {
    conflict('immutable PDF source identity');
  }
  assertAnnotationCollection(
    previous.annotations,
    next.annotations,
    shared.annotations,
  );
  assertMutableRecordCollection(
    previous.formValues,
    next.formValues,
    shared.formValues,
    'PDF form value',
  );
  assertAppendOnlyCollection(
    previous.signaturePlacements,
    next.signaturePlacements,
    shared.signaturePlacements,
    'PDF signature placement',
  );
  assertAppendOnlyCollection(
    previous.redactionProposals,
    next.redactionProposals,
    shared.redactionProposals,
    'PDF redaction proposal',
  );
  assertAppendOnlyCollection(
    previous.pageOperations,
    next.pageOperations,
    shared.pageOperations,
    'PDF page operation',
  );
  assertAppendOnlyCollection(
    previous.reviewDecisions,
    next.reviewDecisions,
    shared.reviewDecisions,
    'PDF review decision',
  );
}

function assertAnnotationCollection(
  previous: readonly WorkPdfCollaborationAnnotation[],
  next: readonly WorkPdfCollaborationAnnotation[],
  shared: readonly WorkPdfCollaborationAnnotation[],
): void {
  const nextById = new Map(next.map((value) => [value.id, value]));
  const sharedById = new Map(shared.map((value) => [value.id, value]));
  const previousIds = new Set(previous.map(({ id }) => id));
  for (const before of previous) {
    const after = nextById.get(before.id);
    const current = sharedById.get(before.id);
    if (!after) {
      throw new WorkOfficeCollaborationError(
        'office.collaboration.content_invalid',
        `PDF annotation '${before.id}' must be retained with deleted: true instead of removing its durable record.`,
      );
    }
    if (after.source !== before.source) {
      throw new WorkOfficeCollaborationError(
        'office.collaboration.content_invalid',
        `PDF annotation '${before.id}' cannot change its source identity.`,
      );
    }
    if (before.deleted && !after.deleted) {
      throw new WorkOfficeCollaborationError(
        'office.collaboration.content_invalid',
        `PDF annotation '${before.id}' has a durable deletion tombstone that cannot be cleared.`,
      );
    }
    if (jsonEqual(before, after)) continue;
    if (!current) removedConflict(`PDF annotation '${before.id}'`);
    if (current.source !== before.source)
      conflict(`PDF annotation '${before.id}' source`);
    assertCompatibleValue(
      before,
      after,
      current,
      `PDF annotation '${before.id}'`,
    );
  }
  for (const after of next) {
    if (previousIds.has(after.id)) continue;
    const current = sharedById.get(after.id);
    if (!current) continue;
    if (after.source !== current.source)
      conflict(`PDF annotation '${after.id}' source`);
    assertCompatibleValue<unknown>(
      {},
      after,
      current,
      `PDF annotation '${after.id}'`,
    );
  }
}

function assertMutableRecordCollection<T extends { id: string }>(
  previous: readonly T[],
  next: readonly T[],
  shared: readonly T[],
  label: string,
): void {
  const nextById = new Map(next.map((value) => [value.id, value]));
  const sharedById = new Map(shared.map((value) => [value.id, value]));
  const previousIds = new Set(previous.map(({ id }) => id));
  for (const before of previous) {
    const after = nextById.get(before.id);
    const current = sharedById.get(before.id);
    if (!after) {
      if (current && !jsonEqual(before, current))
        conflict(`${label} '${before.id}'`);
      continue;
    }
    if (jsonEqual(before, after)) continue;
    if (!current) removedConflict(`${label} '${before.id}'`);
    assertCompatibleValue(before, after, current, `${label} '${before.id}'`);
  }
  for (const after of next) {
    if (previousIds.has(after.id)) continue;
    const current = sharedById.get(after.id);
    if (current)
      assertCompatibleValue<unknown>(
        {},
        after,
        current,
        `${label} '${after.id}'`,
      );
  }
}

function assertAppendOnlyCollection<T extends { id: string }>(
  previous: readonly T[],
  next: readonly T[],
  shared: readonly T[],
  label: string,
): void {
  if (previous.some((value, index) => next[index]?.id !== value.id)) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.content_invalid',
      `${label} records are append-only and cannot be reordered or inserted before an existing record.`,
    );
  }
  const nextById = new Map(next.map((value) => [value.id, value]));
  const sharedById = new Map(shared.map((value) => [value.id, value]));
  const previousIds = new Set(previous.map(({ id }) => id));
  for (const before of previous) {
    const after = nextById.get(before.id);
    if (!after || !jsonEqual(before, after)) {
      throw new WorkOfficeCollaborationError(
        'office.collaboration.content_invalid',
        `${label} '${before.id}' is an append-only audit record and cannot be changed or removed.`,
      );
    }
  }
  for (const after of next) {
    if (previousIds.has(after.id)) continue;
    const current = sharedById.get(after.id);
    if (current && !jsonEqual(after, current))
      conflict(`${label} '${after.id}'`);
  }
}

function assertCompatibleValue<T>(
  previous: T,
  next: T,
  shared: T,
  label: string,
): void {
  if (jsonEqual(previous, next)) return;
  if (isRecord(previous) && isRecord(next) && isRecord(shared)) {
    const keys = new Set([...Object.keys(previous), ...Object.keys(next)]);
    for (const key of keys) {
      const hadPrevious = Object.hasOwn(previous, key);
      const hasNext = Object.hasOwn(next, key);
      const hasShared = Object.hasOwn(shared, key);
      const fieldLabel = `${label} field '${key}'`;
      if (!hasNext) {
        if (
          hadPrevious &&
          hasShared &&
          !jsonEqual(previous[key], shared[key])
        ) {
          conflict(fieldLabel);
        }
        continue;
      }
      if (!hadPrevious) {
        if (hasShared && !jsonEqual(next[key], shared[key]))
          conflict(fieldLabel);
        continue;
      }
      if (jsonEqual(previous[key], next[key])) continue;
      if (!hasShared) removedConflict(fieldLabel);
      assertCompatibleValue(previous[key], next[key], shared[key], fieldLabel);
    }
    return;
  }
  if (!jsonEqual(previous, shared) && !jsonEqual(next, shared)) conflict(label);
}

function removedConflict(label: string): never {
  throw new WorkOfficeCollaborationError(
    'office.collaboration.content_invalid',
    `The ${label} was removed before this change could be applied. Refresh the shared snapshot before retrying.`,
  );
}

function conflict(label: string): never {
  throw new WorkOfficeCollaborationError(
    'office.collaboration.content_invalid',
    `The ${label} changed concurrently. Refresh the shared snapshot before retrying.`,
  );
}
