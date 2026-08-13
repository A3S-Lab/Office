import type * as Y from 'yjs';
import {
  canonicalWorkOfficeCollaborationJson as canonicalJson,
  isWorkOfficeCollaborationRecord as isRecord,
} from './office-collaboration-json';
import type { WorkPdfCollaborationContent } from './office-pdf-collaboration-types';
import { invalidWorkOfficePdfShared as invalidSharedPdf } from './office-pdf-collaboration-validation';
import { WorkOfficeCollaborationError } from './office-collaboration';

type PdfRecordClaimKind =
  | 'annotation'
  | 'page-operation'
  | 'redaction'
  | 'review-decision'
  | 'signature-placement';

interface PdfRecordClaim {
  fingerprint: string;
  id: string;
  kind: PdfRecordClaimKind;
}

export function appendWorkOfficePdfRecordClaims(
  claims: Y.Array<string>,
  previous: WorkPdfCollaborationContent | undefined,
  next: WorkPdfCollaborationContent,
): void {
  appendClaims(
    claims,
    'annotation',
    previous?.annotations ?? [],
    next.annotations,
    annotationFingerprint,
  );
  appendClaims(
    claims,
    'signature-placement',
    previous?.signaturePlacements ?? [],
    next.signaturePlacements,
  );
  appendClaims(
    claims,
    'redaction',
    previous?.redactionProposals ?? [],
    next.redactionProposals,
  );
  appendClaims(
    claims,
    'page-operation',
    previous?.pageOperations ?? [],
    next.pageOperations,
  );
  appendClaims(
    claims,
    'review-decision',
    previous?.reviewDecisions ?? [],
    next.reviewDecisions,
  );
}

export function assertWorkOfficePdfRecordClaims(
  claims: Y.Array<string>,
  content: WorkPdfCollaborationContent,
): void {
  const fingerprints = new Map<string, string>();
  for (const raw of claims.toArray()) {
    const claim = parsedClaim(raw);
    const identity = claimIdentity(claim.kind, claim.id);
    const existing = fingerprints.get(identity);
    if (existing !== undefined && existing !== claim.fingerprint) {
      throw new WorkOfficeCollaborationError(
        'office.collaboration.content_invalid',
        `The PDF ${claim.kind} ID '${claim.id}' was concurrently assigned to different records.`,
      );
    }
    fingerprints.set(identity, claim.fingerprint);
  }
  assertClaims(fingerprints, 'annotation', content.annotations);
  assertClaims(
    fingerprints,
    'signature-placement',
    content.signaturePlacements,
  );
  assertClaims(fingerprints, 'redaction', content.redactionProposals);
  assertClaims(fingerprints, 'page-operation', content.pageOperations);
  assertClaims(fingerprints, 'review-decision', content.reviewDecisions);
}

function appendClaims<T extends { id: string }>(
  claims: Y.Array<string>,
  kind: PdfRecordClaimKind,
  previous: readonly T[],
  next: readonly T[],
  fingerprint: (value: T) => unknown = (value) => value,
): void {
  const previousIds = new Set(previous.map(({ id }) => id));
  for (const value of next) {
    if (previousIds.has(value.id)) continue;
    const serialized = canonicalJson({
      fingerprint: canonicalJson(fingerprint(value)),
      id: value.id,
      kind,
    } satisfies PdfRecordClaim);
    if (!claims.toArray().includes(serialized)) claims.push([serialized]);
  }
}

function annotationFingerprint(
  value: WorkPdfCollaborationContent['annotations'][number],
): unknown {
  return value.source === 'base'
    ? { id: value.id, pageIndex: value.pageIndex, source: value.source }
    : value;
}

function assertClaims<T extends { id: string }>(
  claims: Map<string, string>,
  kind: PdfRecordClaimKind,
  values: readonly T[],
): void {
  for (const { id } of values) {
    if (!claims.has(claimIdentity(kind, id))) {
      invalidSharedPdf(`${kind} record claim`);
    }
  }
}

function parsedClaim(raw: unknown): PdfRecordClaim {
  if (typeof raw !== 'string') invalidSharedPdf('record claim');
  let value: unknown;
  try {
    value = JSON.parse(raw as string);
  } catch {
    invalidSharedPdf('record claim');
  }
  if (!isRecord(value) || canonicalJson(value) !== raw) {
    invalidSharedPdf('record claim');
  }
  if (
    Object.keys(value).some(
      (key) => key !== 'fingerprint' && key !== 'id' && key !== 'kind',
    )
  ) {
    invalidSharedPdf('record claim');
  }
  const kind = requiredClaimKind(value.kind);
  const id = requiredClaimString(value.id, 'record claim ID');
  const fingerprint = requiredClaimString(
    value.fingerprint,
    'record claim fingerprint',
  );
  if (!id.trim() || id.length > 512) invalidSharedPdf('record claim ID');
  try {
    if (canonicalJson(JSON.parse(fingerprint)) !== fingerprint) {
      invalidSharedPdf('record claim fingerprint');
    }
  } catch {
    invalidSharedPdf('record claim fingerprint');
  }
  return { fingerprint, id, kind };
}

function requiredClaimKind(value: unknown): PdfRecordClaimKind {
  if (
    value === 'annotation' ||
    value === 'page-operation' ||
    value === 'redaction' ||
    value === 'review-decision' ||
    value === 'signature-placement'
  ) {
    return value;
  }
  invalidSharedPdf('record claim kind');
}

function requiredClaimString(value: unknown, label: string): string {
  if (typeof value !== 'string') invalidSharedPdf(label);
  return value as string;
}

function claimIdentity(kind: PdfRecordClaimKind, id: string): string {
  return canonicalJson([kind, id]);
}
