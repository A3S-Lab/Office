import { WorkOfficeCollaborationError } from './office-collaboration';
import {
  cloneWorkOfficeCollaborationJson as cloneJsonValue,
  isWorkOfficeCollaborationRecord as isRecord,
} from './office-collaboration-json';
import type {
  WorkPdfCollaborationAnnotation,
  WorkPdfCollaborationContent,
  WorkPdfCollaborationDeletePagesOperation,
  WorkPdfCollaborationFormValue,
  WorkPdfCollaborationPageOperation,
  WorkPdfCollaborationRedactionProposal,
  WorkPdfCollaborationReorderPagesOperation,
  WorkPdfCollaborationReviewDecision,
  WorkPdfCollaborationRotatePagesOperation,
  WorkPdfCollaborationSignaturePlacement,
  WorkPdfCollaborationSource,
} from './office-pdf-collaboration-types';

const MAX_PDF_BYTES = 4 * 1024 * 1024 * 1024;
const MAX_PDF_PAGES = 1_000_000;
const MAX_PDF_RECORDS = 1_000_000;

export function validateWorkOfficePdfCollaborationContent(
  content: WorkPdfCollaborationContent,
): WorkPdfCollaborationContent {
  const record = requiredRecord(content, 'content');
  assertExactKeys(
    record,
    [
      'annotations',
      'formValues',
      'pageOperations',
      'redactionProposals',
      'reviewDecisions',
      'signaturePlacements',
      'source',
      'type',
    ],
    'content',
  );
  if (record.type !== 'pdf') invalid("a PDF content type of 'pdf'");
  const source = validateSource(record.source);
  const result: WorkPdfCollaborationContent = {
    type: 'pdf',
    source,
    annotations: validateRecords(record.annotations, 'annotation', (value) =>
      validateAnnotation(value, source),
    ),
    formValues: validateRecords(
      record.formValues,
      'form value',
      validateFormValue,
    ),
    signaturePlacements: validateRecords(
      record.signaturePlacements,
      'signature placement',
      (value) => validateSignaturePlacement(value, source),
    ),
    redactionProposals: validateRecords(
      record.redactionProposals,
      'redaction proposal',
      (value) => validateRedactionProposal(value, source),
    ),
    pageOperations: validateRecords(
      record.pageOperations,
      'page operation',
      (value) => validatePageOperation(value, source),
    ),
    reviewDecisions: validateRecords(
      record.reviewDecisions,
      'review decision',
      validateReviewDecision,
    ),
  };
  assertReferences(result);
  return result;
}

export function validateSharedWorkOfficePdfCollaborationContent(
  content: WorkPdfCollaborationContent,
): WorkPdfCollaborationContent {
  try {
    return validateWorkOfficePdfCollaborationContent(content);
  } catch (error) {
    if (error instanceof WorkOfficeCollaborationError) {
      throw new WorkOfficeCollaborationError(
        'office.collaboration.content_invalid',
        `The shared PDF collaboration content is invalid: ${error.message}`,
      );
    }
    throw error;
  }
}

export function invalidWorkOfficePdfShared(label: string): never {
  throw new WorkOfficeCollaborationError(
    'office.collaboration.content_invalid',
    `The shared PDF collaboration ${label} is invalid.`,
  );
}

function validateSource(value: unknown): WorkPdfCollaborationSource {
  const source = requiredRecord(value, 'source');
  assertExactKeys(source, ['byteLength', 'pageCount', 'sha256'], 'source');
  const sha256 = requiredString(source.sha256, 'source SHA-256');
  if (!/^[0-9a-f]{64}$/.test(sha256)) invalid('a lowercase SHA-256 source');
  return {
    sha256,
    byteLength: requiredInteger(
      source.byteLength,
      1,
      MAX_PDF_BYTES,
      'source byte length',
    ),
    pageCount: requiredInteger(
      source.pageCount,
      1,
      MAX_PDF_PAGES,
      'source page count',
    ),
  };
}

function validateAnnotation(
  value: unknown,
  source: WorkPdfCollaborationSource,
): WorkPdfCollaborationAnnotation {
  const record = requiredRecord(value, 'annotation');
  assertAllowedKeys(
    record,
    ['annotation', 'deleted', 'id', 'pageIndex', 'source'],
    'annotation',
  );
  const id = requiredIdentifier(record.id, 'annotation');
  const pageIndex = requiredPageIndex(record.pageIndex, source, 'annotation');
  if (record.source !== 'base' && record.source !== 'created') {
    invalid("an annotation source of 'base' or 'created'");
  }
  const annotation = validateJsonRecord(record.annotation, 'annotation value');
  if (annotation.id !== id || annotation.pageIndex !== pageIndex) {
    invalid('annotation value identity matching its record');
  }
  const result: WorkPdfCollaborationAnnotation = {
    id,
    pageIndex,
    source: record.source,
    annotation,
  };
  if (record.deleted !== undefined) {
    if (record.deleted !== true) invalid('an annotation tombstone of true');
    result.deleted = true;
  }
  return result;
}

function validateFormValue(value: unknown): WorkPdfCollaborationFormValue {
  const record = requiredRecord(value, 'form value');
  assertExactKeys(record, ['id', 'value'], 'form value');
  return {
    id: requiredIdentifier(record.id, 'form field'),
    value: requiredString(record.value, 'form field value'),
  };
}

function validateSignaturePlacement(
  value: unknown,
  source: WorkPdfCollaborationSource,
): WorkPdfCollaborationSignaturePlacement {
  const record = requiredRecord(value, 'signature placement');
  assertExactKeys(
    record,
    [
      'actorId',
      'annotationId',
      'appearance',
      'createdAt',
      'id',
      'kind',
      'pageIndex',
    ],
    'signature placement',
  );
  if (record.kind !== 'signature' && record.kind !== 'initials') {
    invalid("a signature placement kind of 'signature' or 'initials'");
  }
  const appearance = requiredRecord(record.appearance, 'signature appearance');
  assertExactKeys(
    appearance,
    ['assetId', 'byteLength', 'mimeType', 'sha256'],
    'signature appearance',
  );
  const sha256 = requiredString(appearance.sha256, 'signature SHA-256');
  if (!/^[0-9a-f]{64}$/.test(sha256)) {
    invalid('a lowercase SHA-256 signature appearance');
  }
  return {
    id: requiredIdentifier(record.id, 'signature placement'),
    annotationId: requiredIdentifier(
      record.annotationId,
      'signature annotation',
    ),
    pageIndex: requiredPageIndex(
      record.pageIndex,
      source,
      'signature placement',
    ),
    kind: record.kind,
    actorId: requiredIdentifier(record.actorId, 'signature actor'),
    createdAt: requiredTimestamp(record.createdAt, 'signature creation'),
    appearance: {
      assetId: requiredIdentifier(appearance.assetId, 'signature asset'),
      sha256,
      byteLength: requiredInteger(
        appearance.byteLength,
        1,
        MAX_PDF_BYTES,
        'signature byte length',
      ),
      mimeType: requiredString(appearance.mimeType, 'signature MIME type'),
    },
  };
}

function validateRedactionProposal(
  value: unknown,
  source: WorkPdfCollaborationSource,
): WorkPdfCollaborationRedactionProposal {
  const record = requiredRecord(value, 'redaction proposal');
  assertAllowedKeys(
    record,
    ['id', 'pageIndex', 'proposedAt', 'proposedBy', 'reason', 'rects', 'text'],
    'redaction proposal',
  );
  const rects = requiredArray(record.rects, 'redaction rectangles');
  if (rects.length === 0 || rects.length > 10_000) {
    invalid('between 1 and 10,000 redaction rectangles');
  }
  const result: WorkPdfCollaborationRedactionProposal = {
    id: requiredIdentifier(record.id, 'redaction proposal'),
    pageIndex: requiredPageIndex(record.pageIndex, source, 'redaction'),
    rects: rects.map((rect) => {
      const item = requiredRecord(rect, 'redaction rectangle');
      assertExactKeys(
        item,
        ['bottom', 'left', 'right', 'top'],
        'redaction rectangle',
      );
      const value = {
        left: requiredFiniteNumber(item.left, 'redaction left'),
        top: requiredFiniteNumber(item.top, 'redaction top'),
        right: requiredFiniteNumber(item.right, 'redaction right'),
        bottom: requiredFiniteNumber(item.bottom, 'redaction bottom'),
      };
      if (value.right <= value.left || value.bottom <= value.top) {
        invalid('a positive redaction rectangle');
      }
      return value;
    }),
    proposedBy: requiredIdentifier(record.proposedBy, 'redaction proposer'),
    proposedAt: requiredTimestamp(record.proposedAt, 'redaction proposal'),
  };
  if (record.reason !== undefined) {
    result.reason = requiredString(record.reason, 'redaction reason');
  }
  if (record.text !== undefined) {
    result.text = requiredString(record.text, 'redaction text');
  }
  return result;
}

function validatePageOperation(
  value: unknown,
  source: WorkPdfCollaborationSource,
): WorkPdfCollaborationPageOperation {
  const record = requiredRecord(value, 'page operation');
  const base = {
    id: requiredIdentifier(record.id, 'page operation'),
    proposedBy: requiredIdentifier(
      record.proposedBy,
      'page operation proposer',
    ),
    proposedAt: requiredTimestamp(record.proposedAt, 'page operation proposal'),
  };
  if (record.kind === 'rotate') {
    assertExactKeys(
      record,
      ['degrees', 'id', 'kind', 'pageIndices', 'proposedAt', 'proposedBy'],
      'rotate page operation',
    );
    if (
      record.degrees !== 90 &&
      record.degrees !== 180 &&
      record.degrees !== 270
    ) {
      invalid('a page rotation of 90, 180, or 270 degrees');
    }
    return {
      ...base,
      kind: 'rotate',
      pageIndices: requiredPageIndices(record.pageIndices, source, false),
      degrees: record.degrees,
    } satisfies WorkPdfCollaborationRotatePagesOperation;
  }
  if (record.kind === 'delete') {
    assertExactKeys(
      record,
      ['id', 'kind', 'pageIndices', 'proposedAt', 'proposedBy'],
      'delete page operation',
    );
    const pageIndices = requiredPageIndices(record.pageIndices, source, false);
    if (pageIndices.length >= source.pageCount) {
      invalid('a page deletion that retains at least one source page');
    }
    return {
      ...base,
      kind: 'delete',
      pageIndices,
    } satisfies WorkPdfCollaborationDeletePagesOperation;
  }
  if (record.kind === 'reorder') {
    assertExactKeys(
      record,
      ['id', 'kind', 'pageOrder', 'proposedAt', 'proposedBy'],
      'reorder page operation',
    );
    const pageOrder = requiredPageIndices(record.pageOrder, source, true);
    return {
      ...base,
      kind: 'reorder',
      pageOrder,
    } satisfies WorkPdfCollaborationReorderPagesOperation;
  }
  invalid('a supported PDF page operation kind');
}

function validateReviewDecision(
  value: unknown,
): WorkPdfCollaborationReviewDecision {
  const record = requiredRecord(value, 'review decision');
  assertExactKeys(
    record,
    ['actorId', 'createdAt', 'decision', 'id', 'targetId', 'targetKind'],
    'review decision',
  );
  if (
    record.targetKind !== 'redaction' &&
    record.targetKind !== 'page-operation'
  ) {
    invalid('a supported review target kind');
  }
  if (record.decision !== 'approve' && record.decision !== 'reject') {
    invalid("a review decision of 'approve' or 'reject'");
  }
  return {
    id: requiredIdentifier(record.id, 'review decision'),
    targetKind: record.targetKind,
    targetId: requiredIdentifier(record.targetId, 'review target'),
    decision: record.decision,
    actorId: requiredIdentifier(record.actorId, 'review actor'),
    createdAt: requiredTimestamp(record.createdAt, 'review decision'),
  };
}

function assertReferences(content: WorkPdfCollaborationContent): void {
  const annotations = new Map(
    content.annotations.map((value) => [value.id, value]),
  );
  for (const placement of content.signaturePlacements) {
    const annotation = annotations.get(placement.annotationId);
    if (
      !annotation ||
      annotation.deleted ||
      annotation.source !== 'created' ||
      annotation.pageIndex !== placement.pageIndex
    ) {
      invalid(
        `signature placement '${placement.id}' to reference a live created annotation on the same page`,
      );
    }
  }
  const redactions = new Set(content.redactionProposals.map(({ id }) => id));
  const pageOperations = new Set(content.pageOperations.map(({ id }) => id));
  const decidedTargets = new Set<string>();
  for (const decision of content.reviewDecisions) {
    const targets =
      decision.targetKind === 'redaction' ? redactions : pageOperations;
    if (!targets.has(decision.targetId)) {
      invalid(
        `review decision '${decision.id}' to reference an existing ${decision.targetKind}`,
      );
    }
    const target = `${decision.targetKind}:${decision.targetId}`;
    if (decidedTargets.has(target)) {
      invalid(`at most one final review decision for '${target}'`);
    }
    decidedTargets.add(target);
  }
}

function validateRecords<T extends { id: string }>(
  value: unknown,
  label: string,
  validate: (value: unknown) => T,
): T[] {
  const records = requiredArray(value, `${label}s`);
  if (records.length > MAX_PDF_RECORDS)
    invalid(`at most ${MAX_PDF_RECORDS} ${label}s`);
  const ids = new Set<string>();
  return records.map((item) => {
    const record = validate(item);
    if (ids.has(record.id))
      invalid(`a unique ${label} ID; '${record.id}' is repeated`);
    ids.add(record.id);
    return record;
  });
}

function requiredPageIndices(
  value: unknown,
  source: WorkPdfCollaborationSource,
  complete: boolean,
): number[] {
  const pages = requiredArray(value, complete ? 'page order' : 'page indices');
  if (
    (complete && pages.length !== source.pageCount) ||
    (!complete && pages.length === 0)
  ) {
    invalid(
      complete
        ? 'a complete source-page permutation'
        : 'at least one page index',
    );
  }
  const result = pages.map((page) =>
    requiredPageIndex(page, source, 'page operation'),
  );
  if (new Set(result).size !== result.length) invalid('unique page indices');
  return result;
}

function validateJsonRecord(
  value: unknown,
  label: string,
): Record<string, unknown> {
  const record = requiredRecord(value, label);
  try {
    return cloneJsonValue(record) as Record<string, unknown>;
  } catch (error) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.content_invalid',
      `PDF collaboration requires a JSON-compatible ${label}: ${error instanceof Error ? error.message : String(error)}`,
    );
  }
}

function requiredPageIndex(
  value: unknown,
  source: WorkPdfCollaborationSource,
  label: string,
): number {
  return requiredInteger(value, 0, source.pageCount - 1, `${label} page index`);
}

function requiredTimestamp(value: unknown, label: string): string {
  const timestamp = requiredString(value, `${label} timestamp`);
  if (!Number.isFinite(Date.parse(timestamp)))
    invalid(`an ISO-compatible ${label} timestamp`);
  return timestamp;
}

function requiredIdentifier(value: unknown, label: string): string {
  const result = requiredString(value, `${label} ID`);
  if (!result || result !== result.trim() || result.length > 512) {
    invalid(`a ${label} ID containing 1 to 512 characters`);
  }
  return result;
}

function requiredString(value: unknown, label: string): string {
  if (typeof value !== 'string') invalid(`a string ${label}`);
  return value as string;
}

function requiredFiniteNumber(value: unknown, label: string): number {
  if (typeof value !== 'number' || !Number.isFinite(value)) {
    invalid(`a finite number for ${label}`);
  }
  return value as number;
}

function requiredInteger(
  value: unknown,
  minimum: number,
  maximum: number,
  label: string,
): number {
  if (
    !Number.isSafeInteger(value) ||
    (value as number) < minimum ||
    (value as number) > maximum
  ) {
    invalid(`an integer ${label} between ${minimum} and ${maximum}`);
  }
  return value as number;
}

function requiredRecord(
  value: unknown,
  label: string,
): Record<string, unknown> {
  if (!isRecord(value)) invalid(`a valid ${label} record`);
  return value as Record<string, unknown>;
}

function requiredArray(value: unknown, label: string): unknown[] {
  if (!Array.isArray(value)) invalid(`an array of ${label}`);
  return value as unknown[];
}

function assertExactKeys(
  record: Record<string, unknown>,
  keys: string[],
  label: string,
): void {
  assertAllowedKeys(record, keys, label);
  if (keys.some((key) => !Object.hasOwn(record, key))) {
    invalid(`a complete ${label} record`);
  }
}

function assertAllowedKeys(
  record: Record<string, unknown>,
  keys: string[],
  label: string,
): void {
  const allowed = new Set(keys);
  if (Object.keys(record).some((key) => !allowed.has(key))) {
    invalid(`a ${label} record without unknown fields`);
  }
}

function invalid(expected: string): never {
  throw new WorkOfficeCollaborationError(
    'office.collaboration.content_invalid',
    `PDF collaboration requires ${expected}.`,
  );
}
