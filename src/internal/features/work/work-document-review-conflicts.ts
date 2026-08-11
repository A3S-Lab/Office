import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import {
  collectDocumentChanges,
  type WorkDocumentChange,
  type WorkDocumentChangeKind,
} from './work-document-changes';
import {
  collectDocumentCommentAnchors,
  type WorkDocumentCommentAnchor,
} from './work-document-comments';
import type { WorkDocumentComment } from './work-types';

export type WorkDocumentReviewKind = 'comment' | WorkDocumentChangeKind;

export type WorkDocumentReviewConflictReason =
  | 'kind-changed'
  | 'record-removed'
  | 'removed'
  | 'text-changed';

export interface WorkDocumentReviewConflict {
  readonly id: string;
  readonly kind: WorkDocumentReviewKind;
  readonly reason: WorkDocumentReviewConflictReason;
  readonly previousText: string;
  readonly nextKind?: WorkDocumentChangeKind;
  readonly nextText?: string;
}

export interface WorkDocumentReviewConflictEvent {
  readonly artifactId?: string;
  readonly conflicts: readonly WorkDocumentReviewConflict[];
}

export interface WorkDocumentReviewSnapshot {
  comments: readonly WorkDocumentCommentAnchor[];
  changes: readonly WorkDocumentChange[];
}

export interface ReconcileDocumentReviewConflictsOptions {
  active: readonly WorkDocumentReviewConflict[];
  before: WorkDocumentReviewSnapshot;
  after: WorkDocumentReviewSnapshot;
  comments: readonly WorkDocumentComment[];
}

export interface ReconciledDocumentReviewConflicts {
  active: WorkDocumentReviewConflict[];
  detected: WorkDocumentReviewConflict[];
}

export function collectDocumentReviewSnapshot(
  document: ProseMirrorNode,
): WorkDocumentReviewSnapshot {
  return {
    comments: collectDocumentCommentAnchors(document),
    changes: collectDocumentChanges(document),
  };
}

export function reconcileDocumentReviewConflicts({
  active,
  before,
  after,
  comments,
}: ReconcileDocumentReviewConflictsOptions): ReconciledDocumentReviewConflicts {
  const commentIds = new Set(
    comments.map((comment) => comment.id.trim()).filter(Boolean),
  );
  const afterComments = new Map(
    after.comments.map((comment) => [comment.id, comment] as const),
  );
  const afterChanges = new Map(
    after.changes.map((change) => [changeKey(change), change] as const),
  );
  const surviving = active.filter((conflict) =>
    conflictStillApplies(conflict, commentIds, afterComments, afterChanges),
  );
  const detected = detectDocumentReviewConflicts(
    before,
    after,
    commentIds,
    afterComments,
    afterChanges,
  );
  const activeKeys = new Set(active.map(documentReviewConflictKey));
  const newlyDetected = detected.filter(
    (conflict) => !activeKeys.has(documentReviewConflictKey(conflict)),
  );
  const byKey = new Map(
    surviving.map(
      (conflict) => [documentReviewConflictKey(conflict), conflict] as const,
    ),
  );
  for (const conflict of detected) {
    const key = documentReviewConflictKey(conflict);
    if (!activeKeys.has(key)) byKey.set(key, conflict);
  }

  return {
    active: Array.from(byKey.values()),
    detected: newlyDetected,
  };
}

function detectDocumentReviewConflicts(
  before: WorkDocumentReviewSnapshot,
  after: WorkDocumentReviewSnapshot,
  commentIds: ReadonlySet<string>,
  afterComments: ReadonlyMap<string, WorkDocumentCommentAnchor>,
  afterChanges: ReadonlyMap<string, WorkDocumentChange>,
): WorkDocumentReviewConflict[] {
  const conflicts: WorkDocumentReviewConflict[] = [];
  for (const previous of before.comments) {
    const next = afterComments.get(previous.id);
    if (!commentIds.has(previous.id)) {
      if (next) {
        conflicts.push({
          id: previous.id,
          kind: 'comment',
          reason: 'record-removed',
          previousText: previous.anchorText,
          nextText: next.anchorText,
        });
      }
      continue;
    }
    if (!next) {
      conflicts.push({
        id: previous.id,
        kind: 'comment',
        reason: 'removed',
        previousText: previous.anchorText,
      });
    } else if (next.anchorText !== previous.anchorText) {
      conflicts.push({
        id: previous.id,
        kind: 'comment',
        reason: 'text-changed',
        previousText: previous.anchorText,
        nextText: next.anchorText,
      });
    }
  }

  const afterChangesById = new Map<string, WorkDocumentChange[]>();
  for (const change of after.changes) {
    const candidates = afterChangesById.get(change.id) ?? [];
    candidates.push(change);
    afterChangesById.set(change.id, candidates);
  }
  for (const previous of before.changes) {
    const next = afterChanges.get(changeKey(previous));
    if (next) {
      if (next.text !== previous.text) {
        conflicts.push({
          id: previous.id,
          kind: previous.kind,
          reason: 'text-changed',
          previousText: previous.text,
          nextText: next.text,
        });
      }
      continue;
    }
    const changedKind = afterChangesById.get(previous.id)?.[0];
    conflicts.push(
      changedKind
        ? {
            id: previous.id,
            kind: previous.kind,
            reason: 'kind-changed',
            previousText: previous.text,
            nextKind: changedKind.kind,
            nextText: changedKind.text,
          }
        : {
            id: previous.id,
            kind: previous.kind,
            reason: 'removed',
            previousText: previous.text,
          },
    );
  }
  return conflicts;
}

function conflictStillApplies(
  conflict: WorkDocumentReviewConflict,
  commentIds: ReadonlySet<string>,
  comments: ReadonlyMap<string, WorkDocumentCommentAnchor>,
  changes: ReadonlyMap<string, WorkDocumentChange>,
): boolean {
  if (conflict.kind === 'comment') {
    if (!commentIds.has(conflict.id)) return comments.has(conflict.id);
    return comments.get(conflict.id)?.anchorText !== conflict.previousText;
  }
  return (
    changes.get(`${conflict.kind}:${conflict.id}`)?.text !==
    conflict.previousText
  );
}

function changeKey(change: Pick<WorkDocumentChange, 'id' | 'kind'>): string {
  return `${change.kind}:${change.id}`;
}

export function documentReviewConflictKey(
  conflict: WorkDocumentReviewConflict,
): string {
  return conflict.kind === 'comment'
    ? `comment:${conflict.id}`
    : `change:${conflict.id}`;
}
