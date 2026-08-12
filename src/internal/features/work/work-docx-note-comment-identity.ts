import {
  documentNoteKey,
  type WorkDocumentNote,
  type WorkDocumentNoteKind,
} from './work-document-notes';
import type {
  WorkDocumentComment,
  WorkDocumentCommentReply,
} from './work-types';

const MAX_NATIVE_ID = 2_147_483_647;
const MIN_NATIVE_COMMENT_ID = -2_147_483_648;

export interface AssignedDocxCommentReply {
  id: number;
  reply: WorkDocumentCommentReply;
}

export interface AssignedDocxCommentThread {
  comment: WorkDocumentComment;
  id: number;
  replies: AssignedDocxCommentReply[];
}

interface CommentCandidate {
  nativeId: number | null;
}

export function importedDocumentNoteId(
  kind: WorkDocumentNoteKind,
  target: string,
): string | null {
  const match = new RegExp(`^${kind}-([1-9][0-9]{0,9})$`).exec(target);
  const nativeId = match ? decimalId(match[1], false) : null;
  return nativeId === null ? null : `docx-${kind}-${nativeId}`;
}

export function assignDocxNoteIds(
  notes: readonly WorkDocumentNote[],
): Map<string, number> {
  const candidates = notes.map((note) => ({
    key: documentNoteKey(note.kind, note.id),
    kind: note.kind,
    nativeId: documentNoteNativeId(note.kind, note.id),
  }));
  const counts = occurrenceCounts(
    candidates.flatMap(({ kind, nativeId }) =>
      nativeId === null ? [] : [`${kind}:${nativeId}`],
    ),
  );
  const used: Record<WorkDocumentNoteKind, Set<number>> = {
    footnote: new Set(),
    endnote: new Set(),
  };
  for (const candidate of candidates) {
    if (
      candidate.nativeId !== null &&
      counts.get(`${candidate.kind}:${candidate.nativeId}`) === 1
    ) {
      used[candidate.kind].add(candidate.nativeId);
    }
  }
  const next = { footnote: 1, endnote: 1 };
  return new Map(
    candidates.map((candidate) => {
      const stable =
        candidate.nativeId !== null &&
        counts.get(`${candidate.kind}:${candidate.nativeId}`) === 1
          ? candidate.nativeId
          : null;
      const id =
        stable ?? nextAvailableId(used[candidate.kind], next, candidate.kind);
      used[candidate.kind].add(id);
      return [candidate.key, id] as const;
    }),
  );
}

export function assignDocxCommentThreads(
  comments: readonly WorkDocumentComment[],
): AssignedDocxCommentThread[] {
  const candidates: CommentCandidate[] = [];
  for (const comment of comments) {
    candidates.push({
      nativeId: documentCommentNativeId(comment.id, false),
    });
    for (const reply of comment.replies ?? []) {
      candidates.push({
        nativeId: documentCommentNativeId(reply.id, true),
      });
    }
  }
  const counts = occurrenceCounts(
    candidates.flatMap(({ nativeId }) =>
      nativeId === null ? [] : [String(nativeId)],
    ),
  );
  const used = new Set<number>();
  for (const candidate of candidates) {
    if (
      candidate.nativeId !== null &&
      counts.get(String(candidate.nativeId)) === 1
    ) {
      used.add(candidate.nativeId);
    }
  }
  let nextId = 0;
  const assigned = new Map<CommentCandidate, number>();
  for (const candidate of candidates) {
    const stable =
      candidate.nativeId !== null &&
      counts.get(String(candidate.nativeId)) === 1
        ? candidate.nativeId
        : null;
    if (stable !== null) {
      assigned.set(candidate, stable);
      continue;
    }
    while (used.has(nextId)) nextId += 1;
    assigned.set(candidate, nextId);
    used.add(nextId);
    nextId += 1;
  }
  let index = 0;
  return comments.map((comment) => {
    const root = candidates[index];
    index += 1;
    const replies = (comment.replies ?? []).map((reply) => {
      const candidate = candidates[index];
      index += 1;
      return { id: assigned.get(candidate) as number, reply };
    });
    return { comment, id: assigned.get(root) as number, replies };
  });
}

function documentNoteNativeId(
  kind: WorkDocumentNoteKind,
  id: string,
): number | null {
  const match = /^docx-(footnote|endnote)-([1-9][0-9]{0,9})$/.exec(id);
  return match?.[1] === kind ? decimalId(match[2], false) : null;
}

function documentCommentNativeId(id: string, reply: boolean): number | null {
  const pattern = reply
    ? /^docx-comment-reply-(0|-?[1-9][0-9]{0,9})$/
    : /^docx-comment-(0|-?[1-9][0-9]{0,9})$/;
  return normalizeDocxCommentId(pattern.exec(id)?.[1]);
}

export function normalizeDocxCommentId(
  value: string | null | undefined,
): number | null {
  if (!value || !/^(?:0|-?[1-9][0-9]{0,9})$/.test(value)) return null;
  const number = Number(value);
  return Number.isSafeInteger(number) &&
    number >= MIN_NATIVE_COMMENT_ID &&
    number <= MAX_NATIVE_ID
    ? number
    : null;
}

function decimalId(
  value: string | undefined,
  allowZero: boolean,
): number | null {
  if (!value) return null;
  const number = Number(value);
  return Number.isSafeInteger(number) &&
    number <= MAX_NATIVE_ID &&
    (allowZero ? number >= 0 : number > 0)
    ? number
    : null;
}

function occurrenceCounts(values: readonly string[]): Map<string, number> {
  const counts = new Map<string, number>();
  for (const value of values) counts.set(value, (counts.get(value) ?? 0) + 1);
  return counts;
}

function nextAvailableId(
  used: ReadonlySet<number>,
  next: Record<WorkDocumentNoteKind, number>,
  kind: WorkDocumentNoteKind,
): number {
  while (used.has(next[kind])) next[kind] += 1;
  const id = next[kind];
  next[kind] += 1;
  return id;
}
