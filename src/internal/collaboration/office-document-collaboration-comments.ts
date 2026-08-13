import * as Y from 'yjs';
import type {
  WorkDocumentComment,
  WorkDocumentCommentReply,
} from '../features/work/work-types';
import {
  assertNoAddedCollision,
  changedRecordMissing,
  insertIntoOrder,
  invalidInputSidecars,
  invalidSharedSidecars,
  isRecord,
  jsonEqual,
  patchRequiredScalar,
  removeFromOrder,
  requiredBoolean,
  requiredIdentifier,
  requiredNestedArray,
  requiredNestedMap,
  requiredSharedBoolean,
  requiredSharedIdentifier,
  requiredSharedMap,
  requiredSharedString,
  requiredString,
  validatedOrder,
} from './office-document-collaboration-sidecar-utils';

export function validatedWorkOfficeDocumentComments(
  value: unknown,
): WorkDocumentComment[] {
  if (!Array.isArray(value)) invalidInputSidecars('an array of comments');
  const ids = new Set<string>();
  return value.map((comment) => {
    const validated = validatedComment(comment);
    if (ids.has(validated.id)) {
      invalidInputSidecars(
        `a unique comment ID; '${validated.id}' is repeated`,
      );
    }
    ids.add(validated.id);
    return validated;
  });
}

export function initializeWorkOfficeDocumentComments(
  records: Y.Map<unknown>,
  order: Y.Array<string>,
  comments: WorkDocumentComment[],
): void {
  for (const comment of comments) {
    createCommentRecord(records, comment);
    order.push([comment.id]);
  }
}

export function readWorkOfficeDocumentComments(
  records: Y.Map<unknown>,
  order: Y.Array<string>,
): WorkDocumentComment[] {
  return validatedOrder(order, records, 'comment').map((id) =>
    readCommentRecord(requiredSharedMap(records, id, 'comment'), id),
  );
}

export function patchWorkOfficeDocumentComments(
  records: Y.Map<unknown>,
  order: Y.Array<string>,
  previous: WorkDocumentComment[],
  next: WorkDocumentComment[],
): void {
  const beforeById = new Map(previous.map((comment) => [comment.id, comment]));
  const afterById = new Map(next.map((comment) => [comment.id, comment]));
  for (const comment of previous) {
    if (afterById.has(comment.id)) continue;
    records.delete(comment.id);
    removeFromOrder(order, comment.id);
  }
  for (const comment of next) {
    const before = beforeById.get(comment.id);
    if (!before) {
      if (!records.has(comment.id)) createCommentRecord(records, comment);
      insertIntoOrder(
        order,
        next.map((item) => item.id),
        comment.id,
      );
      continue;
    }
    if (jsonEqual(before, comment)) continue;
    patchCommentRecord(
      requiredSharedMap(records, comment.id, 'comment'),
      before,
      comment,
    );
  }
}

export function assertWorkOfficeDocumentCommentConflicts(
  previous: WorkDocumentComment[],
  next: WorkDocumentComment[],
  shared: WorkDocumentComment[],
): void {
  assertNoAddedCollision(previous, next, shared, 'comment');
  const beforeById = new Map(previous.map((comment) => [comment.id, comment]));
  const sharedById = new Map(shared.map((comment) => [comment.id, comment]));
  for (const comment of next) {
    const before = beforeById.get(comment.id);
    if (!before) continue;
    const current = sharedById.get(comment.id);
    if (!jsonEqual(before, comment) && !current) {
      changedRecordMissing('comment', comment.id);
    }
    assertNoAddedCollision(
      before.replies ?? [],
      comment.replies ?? [],
      current?.replies ?? [],
      `reply in comment '${comment.id}'`,
    );
    if (!current) continue;
    const beforeReplies = new Map(
      (before.replies ?? []).map((reply) => [reply.id, reply]),
    );
    const sharedReplies = new Set(
      (current.replies ?? []).map((reply) => reply.id),
    );
    for (const reply of comment.replies ?? []) {
      const beforeReply = beforeReplies.get(reply.id);
      if (
        beforeReply &&
        !jsonEqual(beforeReply, reply) &&
        !sharedReplies.has(reply.id)
      ) {
        changedRecordMissing(`reply in comment '${comment.id}'`, reply.id);
      }
    }
  }
}

function validatedComment(value: unknown): WorkDocumentComment {
  if (!isRecord(value)) invalidInputSidecars('valid comment records');
  const id = requiredIdentifier(value.id, 'comment');
  const comment: WorkDocumentComment = {
    id,
    author: requiredString(value.author, 'comment author'),
    date: requiredString(value.date, 'comment date'),
    text: requiredString(value.text, 'comment text'),
    resolved: requiredBoolean(value.resolved, 'comment resolution'),
  };
  if (value.replies !== undefined) {
    if (!Array.isArray(value.replies)) {
      invalidInputSidecars('an array of comment replies');
    }
    const ids = new Set<string>();
    comment.replies = value.replies.map((reply) => {
      const validated = validatedReply(reply);
      if (ids.has(validated.id)) {
        invalidInputSidecars(
          `a unique reply ID within comment '${id}'; '${validated.id}' is repeated`,
        );
      }
      ids.add(validated.id);
      return validated;
    });
  }
  return comment;
}

function validatedReply(value: unknown): WorkDocumentCommentReply {
  if (!isRecord(value)) invalidInputSidecars('valid comment reply records');
  return {
    id: requiredIdentifier(value.id, 'comment reply'),
    author: requiredString(value.author, 'comment reply author'),
    date: requiredString(value.date, 'comment reply date'),
    text: requiredString(value.text, 'comment reply text'),
  };
}

function createCommentRecord(
  records: Y.Map<unknown>,
  comment: WorkDocumentComment,
): void {
  const record = new Y.Map<unknown>();
  records.set(comment.id, record);
  record.set('id', comment.id);
  record.set('author', comment.author);
  record.set('date', comment.date);
  record.set('text', comment.text);
  record.set('resolved', comment.resolved);
  const replies = new Y.Map<unknown>();
  const replyOrder = new Y.Array<string>();
  record.set('replies', replies);
  record.set('replyOrder', replyOrder);
  for (const reply of comment.replies ?? []) {
    createReplyRecord(replies, reply);
    replyOrder.push([reply.id]);
  }
}

function createReplyRecord(
  records: Y.Map<unknown>,
  reply: WorkDocumentCommentReply,
): void {
  const record = new Y.Map<unknown>();
  records.set(reply.id, record);
  record.set('id', reply.id);
  record.set('author', reply.author);
  record.set('date', reply.date);
  record.set('text', reply.text);
}

function readCommentRecord(
  record: Y.Map<unknown>,
  expectedId: string,
): WorkDocumentComment {
  const id = requiredSharedIdentifier(record.get('id'), 'comment');
  if (id !== expectedId) invalidSharedSidecars('comment identity');
  const replies = requiredNestedMap(record, 'replies', 'comment replies');
  const replyOrder = requiredNestedArray(
    record,
    'replyOrder',
    'comment reply order',
  );
  const orderedReplies = validatedOrder(
    replyOrder,
    replies,
    `reply in comment '${id}'`,
  ).map((replyId) =>
    readReplyRecord(
      requiredSharedMap(replies, replyId, 'comment reply'),
      replyId,
    ),
  );
  const comment: WorkDocumentComment = {
    id,
    author: requiredSharedString(record.get('author'), 'comment author'),
    date: requiredSharedString(record.get('date'), 'comment date'),
    text: requiredSharedString(record.get('text'), 'comment text'),
    resolved: requiredSharedBoolean(
      record.get('resolved'),
      'comment resolution',
    ),
  };
  if (orderedReplies.length > 0) comment.replies = orderedReplies;
  return comment;
}

function readReplyRecord(
  record: Y.Map<unknown>,
  expectedId: string,
): WorkDocumentCommentReply {
  const id = requiredSharedIdentifier(record.get('id'), 'comment reply');
  if (id !== expectedId) invalidSharedSidecars('comment reply identity');
  return {
    id,
    author: requiredSharedString(record.get('author'), 'comment reply author'),
    date: requiredSharedString(record.get('date'), 'comment reply date'),
    text: requiredSharedString(record.get('text'), 'comment reply text'),
  };
}

function patchCommentRecord(
  record: Y.Map<unknown>,
  previous: WorkDocumentComment,
  next: WorkDocumentComment,
): void {
  patchRequiredScalar(record, 'author', previous.author, next.author);
  patchRequiredScalar(record, 'date', previous.date, next.date);
  patchRequiredScalar(record, 'text', previous.text, next.text);
  patchRequiredScalar(record, 'resolved', previous.resolved, next.resolved);
  patchReplies(
    requiredNestedMap(record, 'replies', 'comment replies'),
    requiredNestedArray(record, 'replyOrder', 'comment reply order'),
    previous.replies ?? [],
    next.replies ?? [],
  );
}

function patchReplies(
  records: Y.Map<unknown>,
  order: Y.Array<string>,
  previous: WorkDocumentCommentReply[],
  next: WorkDocumentCommentReply[],
): void {
  const beforeById = new Map(previous.map((reply) => [reply.id, reply]));
  const afterById = new Map(next.map((reply) => [reply.id, reply]));
  for (const reply of previous) {
    if (afterById.has(reply.id)) continue;
    records.delete(reply.id);
    removeFromOrder(order, reply.id);
  }
  for (const reply of next) {
    const before = beforeById.get(reply.id);
    if (!before) {
      if (!records.has(reply.id)) createReplyRecord(records, reply);
      insertIntoOrder(
        order,
        next.map((item) => item.id),
        reply.id,
      );
      continue;
    }
    if (jsonEqual(before, reply)) continue;
    const record = requiredSharedMap(records, reply.id, 'comment reply');
    patchRequiredScalar(record, 'author', before.author, reply.author);
    patchRequiredScalar(record, 'date', before.date, reply.date);
    patchRequiredScalar(record, 'text', before.text, reply.text);
  }
}
