import * as Y from 'yjs';
import type {
  WorkDocumentChangeDecision,
  WorkDocumentChangeDecisionAction,
  WorkDocumentChangeKind,
} from '../features/work/work-types';
import { WorkOfficeCollaborationError } from './office-collaboration';
import {
  assertNoAddedCollision,
  insertIntoOrder,
  invalidInputSidecars,
  invalidSharedSidecars,
  isRecord,
  jsonEqual,
  requiredIdentifier,
  requiredSharedIdentifier,
  requiredSharedMap,
  requiredSharedString,
  requiredString,
  validatedOrder,
} from './office-document-collaboration-sidecar-utils';

export function validatedWorkOfficeDocumentChangeDecisions(
  value: unknown,
): WorkDocumentChangeDecision[] {
  if (!Array.isArray(value)) {
    invalidInputSidecars('an array of tracked-change decisions');
  }
  const ids = new Set<string>();
  const changes = new Set<string>();
  return value.map((candidate) => {
    const decision = validatedDecision(candidate);
    if (ids.has(decision.id)) {
      invalidInputSidecars(
        `a unique tracked-change decision ID; '${decision.id}' is repeated`,
      );
    }
    const change = changeIdentity(decision);
    if (changes.has(change)) {
      invalidInputSidecars(
        `one final decision for tracked change '${decision.changeId}'`,
      );
    }
    ids.add(decision.id);
    changes.add(change);
    return decision;
  });
}

export function initializeWorkOfficeDocumentChangeDecisions(
  records: Y.Map<unknown>,
  order: Y.Array<string>,
  decisions: WorkDocumentChangeDecision[],
): void {
  for (const decision of decisions) {
    createDecisionRecord(records, decision);
    order.push([decision.id]);
  }
}

export function readWorkOfficeDocumentChangeDecisions(
  records: Y.Map<unknown>,
  order: Y.Array<string>,
): WorkDocumentChangeDecision[] {
  const decisions = validatedOrder(
    order,
    records,
    'tracked-change decision',
  ).map((id) =>
    readDecisionRecord(
      requiredSharedMap(records, id, 'tracked-change decision'),
      id,
    ),
  );
  const changes = new Set<string>();
  for (const decision of decisions) {
    const identity = changeIdentity(decision);
    if (changes.has(identity)) {
      invalidSharedSidecars(
        `multiple final decisions for tracked change '${decision.changeId}'`,
      );
    }
    changes.add(identity);
  }
  return decisions;
}

export function patchWorkOfficeDocumentChangeDecisions(
  records: Y.Map<unknown>,
  order: Y.Array<string>,
  previous: WorkDocumentChangeDecision[],
  next: WorkDocumentChangeDecision[],
): void {
  const beforeById = new Map(
    previous.map((decision) => [decision.id, decision]),
  );
  for (const decision of previous) {
    const candidate = next.find(({ id }) => id === decision.id);
    if (!candidate || !jsonEqual(candidate, decision)) {
      immutableDecision(decision.id);
    }
  }
  for (const decision of next) {
    if (beforeById.has(decision.id)) continue;
    if (!records.has(decision.id)) createDecisionRecord(records, decision);
    insertIntoOrder(
      order,
      next.map(({ id }) => id),
      decision.id,
    );
  }
}

export function assertWorkOfficeDocumentChangeDecisionConflicts(
  previous: WorkDocumentChangeDecision[],
  next: WorkDocumentChangeDecision[],
  shared: WorkDocumentChangeDecision[],
): void {
  assertNoAddedCollision(previous, next, shared, 'tracked-change decision');
  const nextById = new Map(next.map((decision) => [decision.id, decision]));
  for (const decision of previous) {
    if (!jsonEqual(nextById.get(decision.id), decision)) {
      immutableDecision(decision.id);
    }
  }
  const sharedByChange = new Map(
    shared.map((decision) => [changeIdentity(decision), decision]),
  );
  const previousIds = new Set(previous.map(({ id }) => id));
  for (const decision of next) {
    if (previousIds.has(decision.id)) continue;
    const current = sharedByChange.get(changeIdentity(decision));
    if (current && !jsonEqual(current, decision)) {
      throw new WorkOfficeCollaborationError(
        'office.collaboration.content_invalid',
        `Tracked change '${decision.changeId}' already has a different final decision.`,
      );
    }
  }
}

function validatedDecision(value: unknown): WorkDocumentChangeDecision {
  if (!isRecord(value)) {
    invalidInputSidecars('valid tracked-change decision records');
  }
  const decision: WorkDocumentChangeDecision = {
    id: requiredIdentifier(value.id, 'tracked-change decision'),
    changeId: requiredIdentifier(value.changeId, 'tracked change'),
    changeKind: changeKind(value.changeKind, false),
    suggestedBy: requiredString(value.suggestedBy, 'suggestion author'),
    suggestedAt: requiredString(value.suggestedAt, 'suggestion date'),
    text: requiredString(value.text, 'suggestion text'),
    decision: decisionAction(value.decision, false),
    decidedBy: requiredString(value.decidedBy, 'decision author'),
    decidedAt: requiredString(value.decidedAt, 'decision date'),
  };
  if (value.suggestedByActorId !== undefined) {
    decision.suggestedByActorId = requiredIdentifier(
      value.suggestedByActorId,
      'suggestion actor',
    );
  }
  if (value.decidedByActorId !== undefined) {
    decision.decidedByActorId = requiredIdentifier(
      value.decidedByActorId,
      'decision actor',
    );
  }
  return decision;
}

function createDecisionRecord(
  records: Y.Map<unknown>,
  decision: WorkDocumentChangeDecision,
): void {
  const record = new Y.Map<unknown>();
  records.set(decision.id, record);
  for (const [key, value] of Object.entries(decision)) record.set(key, value);
}

function readDecisionRecord(
  record: Y.Map<unknown>,
  expectedId: string,
): WorkDocumentChangeDecision {
  const id = requiredSharedIdentifier(
    record.get('id'),
    'tracked-change decision',
  );
  if (id !== expectedId)
    invalidSharedSidecars('tracked-change decision identity');
  const allowed = new Set([
    'id',
    'changeId',
    'changeKind',
    'suggestedByActorId',
    'suggestedBy',
    'suggestedAt',
    'text',
    'decision',
    'decidedByActorId',
    'decidedBy',
    'decidedAt',
  ]);
  if (Array.from(record.keys()).some((key) => !allowed.has(key))) {
    invalidSharedSidecars('tracked-change decision fields');
  }
  const decision: WorkDocumentChangeDecision = {
    id,
    changeId: requiredSharedIdentifier(
      record.get('changeId'),
      'tracked change',
    ),
    changeKind: changeKind(record.get('changeKind'), true),
    suggestedBy: requiredSharedString(
      record.get('suggestedBy'),
      'suggestion author',
    ),
    suggestedAt: requiredSharedString(
      record.get('suggestedAt'),
      'suggestion date',
    ),
    text: requiredSharedString(record.get('text'), 'suggestion text'),
    decision: decisionAction(record.get('decision'), true),
    decidedBy: requiredSharedString(record.get('decidedBy'), 'decision author'),
    decidedAt: requiredSharedString(record.get('decidedAt'), 'decision date'),
  };
  const suggestedByActorId = record.get('suggestedByActorId');
  if (suggestedByActorId !== undefined) {
    decision.suggestedByActorId = requiredSharedIdentifier(
      suggestedByActorId,
      'suggestion actor',
    );
  }
  const decidedByActorId = record.get('decidedByActorId');
  if (decidedByActorId !== undefined) {
    decision.decidedByActorId = requiredSharedIdentifier(
      decidedByActorId,
      'decision actor',
    );
  }
  return decision;
}

function changeIdentity(
  decision: Pick<WorkDocumentChangeDecision, 'changeId' | 'changeKind'>,
): string {
  return `${decision.changeKind}:${decision.changeId}`;
}

function changeKind(value: unknown, shared: boolean): WorkDocumentChangeKind {
  if (value === 'insertion' || value === 'deletion' || value === 'formatting') {
    return value;
  }
  if (shared) invalidSharedSidecars('tracked-change decision kind');
  invalidInputSidecars(
    'an insertion, deletion, or formatting tracked-change kind',
  );
}

function decisionAction(
  value: unknown,
  shared: boolean,
): WorkDocumentChangeDecisionAction {
  if (value === 'accept' || value === 'reject') return value;
  if (shared) invalidSharedSidecars('tracked-change decision action');
  invalidInputSidecars('an accept or reject tracked-change decision');
}

function immutableDecision(id: string): never {
  throw new WorkOfficeCollaborationError(
    'office.collaboration.permission_denied',
    `Tracked-change decision '${id}' is immutable and cannot be rewritten or removed.`,
  );
}
