import type * as Y from 'yjs';
import type {
  WorkPresentationContent,
  WorkSlideElement,
} from '../features/work/work-types';
import { WorkOfficeCollaborationError } from './office-collaboration';
import {
  canonicalWorkOfficeCollaborationJson as canonicalJson,
  isWorkOfficeCollaborationRecord as isRecord,
} from './office-collaboration-json';
import { invalidWorkOfficePresentationShared as invalidSharedPresentation } from './office-presentation-collaboration-validation';

export type WorkOfficePresentationContainerKind = 'layout' | 'master' | 'slide';

interface PresentationElementClaim {
  containerId: string;
  containerKind: WorkOfficePresentationContainerKind;
  fingerprint: string;
  id: string;
  kind: 'element';
}

interface PresentationElementContainer {
  id: string;
  elements: WorkSlideElement[];
}

export function appendWorkOfficePresentationRecordClaims(
  claims: Y.Array<string>,
  previous: WorkPresentationContent | undefined,
  next: WorkPresentationContent,
): void {
  appendContainerClaims(claims, 'slide', previous?.slides ?? [], next.slides);
  appendContainerClaims(
    claims,
    'master',
    previous?.masters ?? [],
    next.masters ?? [],
  );
  appendContainerClaims(
    claims,
    'layout',
    previous?.layouts ?? [],
    next.layouts ?? [],
  );
}

export function assertWorkOfficePresentationRecordClaims(
  claims: Y.Array<string>,
  content: WorkPresentationContent,
): void {
  const fingerprints = new Map<string, string>();
  for (const raw of claims.toArray()) {
    const claim = parsedClaim(raw);
    const identity = claimIdentity(
      claim.containerKind,
      claim.containerId,
      claim.id,
    );
    const existing = fingerprints.get(identity);
    if (existing !== undefined && existing !== claim.fingerprint) {
      throw new WorkOfficeCollaborationError(
        'office.collaboration.content_invalid',
        `The scene element in ${claim.containerKind} '${claim.containerId}' ID '${claim.id}' was concurrently assigned to different records.`,
      );
    }
    fingerprints.set(identity, claim.fingerprint);
  }

  assertClaimedElementIdentities(fingerprints, 'slide', content.slides);
  assertClaimedElementIdentities(fingerprints, 'master', content.masters ?? []);
  assertClaimedElementIdentities(fingerprints, 'layout', content.layouts ?? []);
}

function appendContainerClaims(
  claims: Y.Array<string>,
  containerKind: WorkOfficePresentationContainerKind,
  previous: readonly PresentationElementContainer[],
  next: readonly PresentationElementContainer[],
): void {
  const previousById = new Map(
    previous.map((container) => [container.id, container]),
  );
  for (const container of next) {
    const previousElementIds = new Set(
      (previousById.get(container.id)?.elements ?? []).map(({ id }) => id),
    );
    for (const element of container.elements) {
      if (previousElementIds.has(element.id)) continue;
      appendClaim(claims, containerKind, container.id, element);
    }
  }
}

function appendClaim(
  claims: Y.Array<string>,
  containerKind: WorkOfficePresentationContainerKind,
  containerId: string,
  element: WorkSlideElement,
): void {
  const serialized = canonicalJson({
    containerId,
    containerKind,
    fingerprint: canonicalJson(element),
    id: element.id,
    kind: 'element',
  } satisfies PresentationElementClaim);
  if (!claims.toArray().includes(serialized)) claims.push([serialized]);
}

function parsedClaim(raw: unknown): PresentationElementClaim {
  if (typeof raw !== 'string') invalidSharedPresentation('record claim');
  let value: unknown;
  try {
    value = JSON.parse(raw as string);
  } catch {
    invalidSharedPresentation('record claim');
  }
  if (!isRecord(value) || canonicalJson(value) !== raw) {
    invalidSharedPresentation('record claim');
  }
  if (
    Object.keys(value).some(
      (key) =>
        key !== 'containerId' &&
        key !== 'containerKind' &&
        key !== 'fingerprint' &&
        key !== 'id' &&
        key !== 'kind',
    )
  ) {
    invalidSharedPresentation('record claim');
  }
  if (value.kind !== 'element') {
    invalidSharedPresentation('record claim kind');
  }
  const containerKind = requiredContainerKind(value.containerKind);
  const containerId = requiredIdentifier(
    value.containerId,
    'record claim container',
  );
  const id = requiredIdentifier(value.id, 'record claim');
  const fingerprint = requiredString(
    value.fingerprint,
    'record claim fingerprint',
  );
  let claimed: unknown;
  try {
    claimed = JSON.parse(fingerprint);
  } catch {
    invalidSharedPresentation('record claim fingerprint');
  }
  if (
    !isRecord(claimed) ||
    canonicalJson(claimed) !== fingerprint ||
    claimed.id !== id ||
    typeof claimed.type !== 'string'
  ) {
    invalidSharedPresentation('record claim fingerprint');
  }
  return {
    containerId,
    containerKind,
    fingerprint,
    id,
    kind: 'element',
  };
}

function assertClaimedElementIdentities(
  claims: ReadonlyMap<string, string>,
  containerKind: WorkOfficePresentationContainerKind,
  containers: readonly PresentationElementContainer[],
): void {
  for (const container of containers) {
    for (const element of container.elements) {
      const fingerprint = claims.get(
        claimIdentity(containerKind, container.id, element.id),
      );
      // Claims were introduced after the first Presentation collaboration
      // format. Missing claims are therefore valid for legacy records.
      if (fingerprint === undefined) continue;
      let claimed: unknown;
      try {
        claimed = JSON.parse(fingerprint);
      } catch {
        invalidSharedPresentation('record claim fingerprint');
      }
      if (
        !isRecord(claimed) ||
        claimed.id !== element.id ||
        claimed.type !== element.type
      ) {
        invalidSharedPresentation('scene element immutable identity claim');
      }
    }
  }
}

function requiredContainerKind(
  value: unknown,
): WorkOfficePresentationContainerKind {
  if (value === 'layout' || value === 'master' || value === 'slide') {
    return value;
  }
  invalidSharedPresentation('record claim container kind');
}

function requiredIdentifier(value: unknown, label: string): string {
  const result = requiredString(value, label);
  if (!result || result !== result.trim() || result.length > 256) {
    invalidSharedPresentation(label);
  }
  return result;
}

function requiredString(value: unknown, label: string): string {
  if (typeof value !== 'string') invalidSharedPresentation(label);
  return value as string;
}

function claimIdentity(
  containerKind: WorkOfficePresentationContainerKind,
  containerId: string,
  id: string,
): string {
  return canonicalJson([containerKind, containerId, id]);
}
