import type {
  WorkDocumentComment,
  WorkDocumentCommentReply,
} from '../features/work/work-types';
import {
  WorkOfficeCollaborationError,
  type WorkOfficeCollaborationActor,
  type WorkOfficeCollaborationSession,
} from './office-collaboration';
import { jsonEqual } from './office-document-collaboration-sidecar-utils';

export function assertWorkOfficeDocumentCommentMutationAllowed(
  session: WorkOfficeCollaborationSession,
  previous: readonly WorkDocumentComment[],
  next: readonly WorkDocumentComment[],
): void {
  if (session.mode === 'edit') return;
  if (session.mode !== 'comment') deniedReviewMutation(session);
  const actor = requiredReviewActor(session);
  assertRetainedOrder(previous, next, 'comments');
  const nextById = new Map(next.map((comment) => [comment.id, comment]));
  const previousIds = new Set(previous.map((comment) => comment.id));

  for (const comment of previous) {
    const updated = nextById.get(comment.id);
    if (!updated) {
      assertOwnedRecord(comment.actorId, actor, 'own comment');
      continue;
    }
    assertExistingCommentMutation(actor, comment, updated);
  }
  for (const comment of next) {
    if (!previousIds.has(comment.id)) assertNewComment(actor, comment);
  }
}

function assertExistingCommentMutation(
  actor: WorkOfficeCollaborationActor,
  previous: WorkDocumentComment,
  next: WorkDocumentComment,
): void {
  if (
    previous.id !== next.id ||
    previous.actorId !== next.actorId ||
    previous.author !== next.author ||
    previous.date !== next.date ||
    previous.text !== next.text
  ) {
    deniedImmutableRecord('comment');
  }
  assertReplyMutations(
    actor,
    previous.id,
    previous.replies ?? [],
    next.replies ?? [],
  );
}

function assertReplyMutations(
  actor: WorkOfficeCollaborationActor,
  commentId: string,
  previous: readonly WorkDocumentCommentReply[],
  next: readonly WorkDocumentCommentReply[],
): void {
  assertRetainedOrder(previous, next, `replies in comment '${commentId}'`);
  const nextById = new Map(next.map((reply) => [reply.id, reply]));
  const previousIds = new Set(previous.map((reply) => reply.id));
  for (const reply of previous) {
    const updated = nextById.get(reply.id);
    if (!updated) {
      assertOwnedRecord(reply.actorId, actor, 'own comment reply');
      continue;
    }
    if (!jsonEqual(reply, updated)) deniedImmutableRecord('comment reply');
  }
  for (const reply of next) {
    if (!previousIds.has(reply.id)) assertActorRecord(actor, reply, 'reply');
  }
}

function assertNewComment(
  actor: WorkOfficeCollaborationActor,
  comment: WorkDocumentComment,
): void {
  assertActorRecord(actor, comment, 'comment');
  for (const reply of comment.replies ?? []) {
    assertActorRecord(actor, reply, 'reply');
  }
}

function assertActorRecord(
  actor: WorkOfficeCollaborationActor,
  record: Pick<WorkDocumentComment, 'actorId' | 'author'>,
  label: string,
): void {
  if (record.actorId === actor.id && record.author === actor.name) return;
  throw new WorkOfficeCollaborationError(
    'office.collaboration.permission_denied',
    `A Document ${label} created in comment mode must use collaboration actor '${actor.id}' and its current display name.`,
  );
}

function assertOwnedRecord(
  actorId: string | undefined,
  actor: WorkOfficeCollaborationActor,
  label: string,
): void {
  if (actorId === actor.id) return;
  throw new WorkOfficeCollaborationError(
    'office.collaboration.permission_denied',
    `Comment mode can delete only the collaboration actor's ${label}.`,
  );
}

function assertRetainedOrder<T extends { id: string }>(
  previous: readonly T[],
  next: readonly T[],
  label: string,
): void {
  const previousIds = new Set(previous.map((record) => record.id));
  const nextIds = new Set(next.map((record) => record.id));
  const retainedBefore = previous
    .filter((record) => nextIds.has(record.id))
    .map((record) => record.id);
  const retainedAfter = next
    .filter((record) => previousIds.has(record.id))
    .map((record) => record.id);
  let encounteredNewRecord = false;
  let existingAfterNewRecord = false;
  for (const record of next) {
    if (previousIds.has(record.id)) {
      if (encounteredNewRecord) existingAfterNewRecord = true;
    } else {
      encounteredNewRecord = true;
    }
  }
  if (jsonEqual(retainedBefore, retainedAfter) && !existingAfterNewRecord) {
    return;
  }
  throw new WorkOfficeCollaborationError(
    'office.collaboration.permission_denied',
    `Comment mode must preserve existing ${label} order and append new review records.`,
  );
}

function requiredReviewActor(
  session: WorkOfficeCollaborationSession,
): WorkOfficeCollaborationActor {
  if (session.actor) return session.actor;
  throw new WorkOfficeCollaborationError(
    'office.collaboration.permission_denied',
    'Comment mode requires a collaboration actor before it can write review records.',
  );
}

function deniedImmutableRecord(label: string): never {
  throw new WorkOfficeCollaborationError(
    'office.collaboration.permission_denied',
    `Comment mode cannot rewrite an existing ${label}; it may append replies or change resolution state.`,
  );
}

function deniedReviewMutation(session: WorkOfficeCollaborationSession): never {
  throw new WorkOfficeCollaborationError(
    'office.collaboration.permission_denied',
    `The '${session.mode}' collaboration mode cannot modify Document review records.`,
  );
}
