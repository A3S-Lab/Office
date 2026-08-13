import type * as Y from 'yjs';
import type {
  WorkSpreadsheetContent,
  WorkSpreadsheetSheet,
} from '../features/work/work-types';
import { WorkOfficeCollaborationError } from './office-collaboration';
import {
  canonicalWorkOfficeCollaborationJson as canonicalJson,
  isWorkOfficeCollaborationRecord as isRecord,
} from './office-collaboration-json';
import { invalidWorkOfficeSpreadsheetShared as invalidSharedSpreadsheet } from './office-spreadsheet-collaboration-validation';

type SpreadsheetRecordClaimKind =
  | 'chart'
  | 'image'
  | 'named-range'
  | 'pivot-table'
  | 'sheet';

interface SpreadsheetRecordClaim {
  fingerprint: string;
  id: string;
  kind: SpreadsheetRecordClaimKind;
  parentId?: string;
}

export function appendWorkOfficeSpreadsheetRecordClaims(
  claims: Y.Array<string>,
  previous: WorkSpreadsheetContent | undefined,
  next: WorkSpreadsheetContent,
): void {
  const previousSheets = new Map(
    (previous?.sheets ?? []).map((sheet) => [sheet.id as string, sheet]),
  );
  for (const sheet of next.sheets) {
    const before = previousSheets.get(sheet.id as string);
    if (!before) appendClaim(claims, 'sheet', sheet.id as string, sheet);
    appendChildClaims(claims, before, sheet, 'image', 'images');
    appendChildClaims(claims, before, sheet, 'chart', 'charts');
    appendChildClaims(claims, before, sheet, 'pivot-table', 'pivotTables');
  }
  const previousNamedRanges = new Set(
    (previous?.namedRanges ?? []).map(({ id }) => id),
  );
  for (const range of next.namedRanges ?? []) {
    if (!previousNamedRanges.has(range.id)) {
      appendClaim(claims, 'named-range', range.id, range);
    }
  }
}

export function assertWorkOfficeSpreadsheetRecordClaims(
  claims: Y.Array<string>,
  content: WorkSpreadsheetContent,
): void {
  const fingerprints = new Map<string, string>();
  for (const raw of claims.toArray()) {
    const claim = parsedClaim(raw);
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
  for (const sheet of content.sheets) {
    assertClaimExists(fingerprints, 'sheet', sheet.id as string);
    for (const image of sheet.images ?? []) {
      assertClaimExists(fingerprints, 'image', image.id, sheet.id as string);
    }
    for (const chart of sheet.charts ?? []) {
      assertClaimExists(fingerprints, 'chart', chart.id, sheet.id as string);
    }
    for (const pivot of sheet.pivotTables ?? []) {
      assertClaimExists(
        fingerprints,
        'pivot-table',
        pivot.id,
        sheet.id as string,
      );
    }
  }
  for (const range of content.namedRanges ?? []) {
    assertClaimExists(fingerprints, 'named-range', range.id);
  }
}

function appendChildClaims(
  claims: Y.Array<string>,
  previous: WorkSpreadsheetSheet | undefined,
  next: WorkSpreadsheetSheet,
  kind: 'chart' | 'image' | 'pivot-table',
  key: 'charts' | 'images' | 'pivotTables',
): void {
  const previousIds = new Set(
    (previous?.[key] ?? []).map(({ id }) => id as string),
  );
  for (const value of next[key] ?? []) {
    if (!previousIds.has(value.id)) {
      appendClaim(claims, kind, value.id, value, next.id as string);
    }
  }
}

function appendClaim(
  claims: Y.Array<string>,
  kind: SpreadsheetRecordClaimKind,
  id: string,
  value: unknown,
  parentId?: string,
): void {
  const serialized = canonicalJson({
    fingerprint: canonicalJson(value),
    id,
    kind,
    ...(parentId === undefined ? {} : { parentId }),
  } satisfies SpreadsheetRecordClaim);
  if (!claims.toArray().includes(serialized)) claims.push([serialized]);
}

function parsedClaim(raw: unknown): SpreadsheetRecordClaim {
  if (typeof raw !== 'string') invalidSharedSpreadsheet('record claim');
  let value: unknown;
  try {
    value = JSON.parse(raw as string);
  } catch {
    invalidSharedSpreadsheet('record claim');
  }
  if (!isRecord(value) || canonicalJson(value) !== raw) {
    invalidSharedSpreadsheet('record claim');
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
    invalidSharedSpreadsheet('record claim');
  }
  const kind = requiredClaimKind(value.kind);
  const id = requiredClaimIdentifier(value.id, 'record claim');
  const fingerprint = requiredClaimString(
    value.fingerprint,
    'record claim fingerprint',
  );
  const parentId =
    value.parentId === undefined
      ? undefined
      : requiredClaimIdentifier(value.parentId, 'record claim parent');
  if (
    (kind === 'sheet' || kind === 'named-range') ===
    (parentId !== undefined)
  ) {
    invalidSharedSpreadsheet('record claim parent');
  }
  try {
    const fingerprintValue = JSON.parse(fingerprint);
    if (canonicalJson(fingerprintValue) !== fingerprint) {
      invalidSharedSpreadsheet('record claim fingerprint');
    }
  } catch {
    invalidSharedSpreadsheet('record claim fingerprint');
  }
  return { fingerprint, id, kind, parentId };
}

function requiredClaimKind(value: unknown): SpreadsheetRecordClaimKind {
  if (
    value === 'chart' ||
    value === 'image' ||
    value === 'named-range' ||
    value === 'pivot-table' ||
    value === 'sheet'
  ) {
    return value;
  }
  invalidSharedSpreadsheet('record claim kind');
}

function requiredClaimIdentifier(value: unknown, label: string): string {
  const result = requiredClaimString(value, label);
  if (!result || result !== result.trim() || result.length > 256) {
    invalidSharedSpreadsheet(label);
  }
  return result;
}

function requiredClaimString(value: unknown, label: string): string {
  if (typeof value !== 'string') invalidSharedSpreadsheet(label);
  return value as string;
}

function assertClaimExists(
  claims: Map<string, string>,
  kind: SpreadsheetRecordClaimKind,
  id: string,
  parentId?: string,
): void {
  if (claims.has(claimIdentity(kind, id, parentId))) return;
  invalidSharedSpreadsheet(`${claimLabel({ kind, parentId })} record claim`);
}

function claimIdentity(
  kind: SpreadsheetRecordClaimKind,
  id: string,
  parentId?: string,
): string {
  return canonicalJson([kind, parentId ?? null, id]);
}

function claimLabel(
  claim: Pick<SpreadsheetRecordClaim, 'kind' | 'parentId'>,
): string {
  if (claim.kind === 'named-range') return 'named range';
  if (claim.kind === 'pivot-table') {
    return `pivot table in sheet '${claim.parentId}'`;
  }
  if (claim.kind === 'sheet') return 'sheet';
  return `${claim.kind} in sheet '${claim.parentId}'`;
}
