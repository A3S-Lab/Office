import type * as Y from 'yjs';
import type {
  WorkDocumentBibliography,
  WorkDocumentCitationSource,
  WorkDocumentComment,
  WorkDocumentCommentReply,
} from '../features/work/work-types';
import { WorkOfficeCollaborationError } from './office-collaboration';
import {
  canonicalJson,
  invalidSharedSidecars,
  isRecord,
  requiredSharedIdentifier,
  requiredSharedString,
} from './office-document-collaboration-sidecar-utils';

type DocumentRecordClaimKind =
  | 'bibliography-source'
  | 'comment'
  | 'comment-reply';

interface ClaimableDocumentSidecars {
  comments?: WorkDocumentComment[];
  bibliography?: WorkDocumentBibliography;
}

interface DocumentRecordClaim {
  fingerprint: string;
  id: string;
  kind: DocumentRecordClaimKind;
  parentId?: string;
}

export function appendWorkOfficeDocumentRecordClaims(
  claims: Y.Array<string>,
  previous: ClaimableDocumentSidecars,
  next: ClaimableDocumentSidecars,
): void {
  const previousComments = new Map(
    (previous.comments ?? []).map((comment) => [comment.id, comment]),
  );
  for (const comment of next.comments ?? []) {
    const previousComment = previousComments.get(comment.id);
    if (!previousComment) appendClaim(claims, 'comment', comment.id, comment);
    const previousReplies = new Set(
      (previousComment?.replies ?? []).map((reply) => reply.id),
    );
    for (const reply of comment.replies ?? []) {
      if (!previousReplies.has(reply.id)) {
        appendClaim(claims, 'comment-reply', reply.id, reply, comment.id);
      }
    }
  }

  const previousSources = new Set(
    (previous.bibliography?.sources ?? []).map((source) => source.id),
  );
  for (const source of next.bibliography?.sources ?? []) {
    if (!previousSources.has(source.id)) {
      appendClaim(claims, 'bibliography-source', source.id, source);
    }
  }
}

export function assertWorkOfficeDocumentRecordClaims(
  claims: Y.Array<string>,
  sidecars: ClaimableDocumentSidecars,
): void {
  const fingerprints = new Map<string, string>();
  for (const rawClaim of claims.toArray()) {
    const claim = parsedClaim(rawClaim);
    const identity = claimIdentity(claim.kind, claim.id, claim.parentId);
    const existing = fingerprints.get(identity);
    if (existing !== undefined && existing !== claim.fingerprint) {
      throw new WorkOfficeCollaborationError(
        'office.collaboration.content_invalid',
        `The ${claimLabel(claim)} ID '${claim.id}' was concurrently assigned to different records.`,
      );
    }
    fingerprints.set(identity, claim.fingerprint);
  }

  for (const comment of sidecars.comments ?? []) {
    assertClaimExists(fingerprints, 'comment', comment.id);
    for (const reply of comment.replies ?? []) {
      assertClaimExists(fingerprints, 'comment-reply', reply.id, comment.id);
    }
  }
  for (const source of sidecars.bibliography?.sources ?? []) {
    assertClaimExists(fingerprints, 'bibliography-source', source.id);
  }
}

function appendClaim(
  claims: Y.Array<string>,
  kind: DocumentRecordClaimKind,
  id: string,
  value:
    | WorkDocumentCitationSource
    | WorkDocumentComment
    | WorkDocumentCommentReply,
  parentId?: string,
): void {
  const serialized = canonicalJson({
    fingerprint: canonicalJson(value),
    id,
    kind,
    ...(parentId === undefined ? {} : { parentId }),
  } satisfies DocumentRecordClaim);
  if (!claims.toArray().includes(serialized)) claims.push([serialized]);
}

function parsedClaim(rawClaim: unknown): DocumentRecordClaim {
  if (typeof rawClaim !== 'string') invalidSharedSidecars('record claim');
  let value: unknown;
  try {
    value = JSON.parse(rawClaim as string);
  } catch {
    invalidSharedSidecars('record claim');
  }
  if (!isRecord(value) || canonicalJson(value) !== rawClaim) {
    invalidSharedSidecars('record claim');
  }
  const keys = Object.keys(value);
  if (
    keys.some(
      (key) =>
        key !== 'fingerprint' &&
        key !== 'id' &&
        key !== 'kind' &&
        key !== 'parentId',
    )
  ) {
    invalidSharedSidecars('record claim');
  }
  const kind = value.kind;
  if (
    kind !== 'bibliography-source' &&
    kind !== 'comment' &&
    kind !== 'comment-reply'
  ) {
    invalidSharedSidecars('record claim kind');
  }
  const id = requiredSharedIdentifier(value.id, 'record claim');
  const fingerprint = requiredSharedString(
    value.fingerprint,
    'record claim fingerprint',
  );
  const parentId =
    value.parentId === undefined
      ? undefined
      : requiredSharedIdentifier(value.parentId, 'record claim parent');
  if ((kind === 'comment-reply') !== (parentId !== undefined)) {
    invalidSharedSidecars('record claim parent');
  }
  try {
    const fingerprintValue = JSON.parse(fingerprint);
    if (
      !isRecord(fingerprintValue) ||
      canonicalJson(fingerprintValue) !== fingerprint
    ) {
      invalidSharedSidecars('record claim fingerprint');
    }
  } catch {
    invalidSharedSidecars('record claim fingerprint');
  }
  return { fingerprint, id, kind, parentId };
}

function assertClaimExists(
  claims: Map<string, string>,
  kind: DocumentRecordClaimKind,
  id: string,
  parentId?: string,
): void {
  if (claims.has(claimIdentity(kind, id, parentId))) return;
  invalidSharedSidecars(`${claimLabel({ kind, parentId })} record claim`);
}

function claimIdentity(
  kind: DocumentRecordClaimKind,
  id: string,
  parentId?: string,
): string {
  return canonicalJson([kind, parentId ?? null, id]);
}

function claimLabel(
  claim: Pick<DocumentRecordClaim, 'kind' | 'parentId'>,
): string {
  if (claim.kind === 'comment') return 'comment';
  if (claim.kind === 'bibliography-source') return 'bibliography source';
  return `reply in comment '${claim.parentId}'`;
}
