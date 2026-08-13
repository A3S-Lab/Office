import type * as Y from 'yjs';
import { WorkOfficeCollaborationError } from './office-collaboration';
import { canonicalWorkOfficeCollaborationJson as canonicalJson } from './office-collaboration-json';
import {
  appendWorkOfficePdfRecordClaims,
  assertWorkOfficePdfRecordClaims,
} from './office-pdf-collaboration-claims';
import { assertWorkOfficePdfPatchSafe } from './office-pdf-collaboration-conflicts';
import {
  assertWorkOfficePdfRecordCollectionEmpty,
  patchWorkOfficePdfRecords,
  readWorkOfficePdfRecords,
  type WorkOfficePdfRecordCollectionRoots,
  workOfficePdfRecordCollectionRoots,
  workOfficePdfRecordCollectionUndoScope,
  workOfficePdfRecordFieldChanged,
} from './office-pdf-collaboration-records';
import type {
  WorkPdfCollaborationAnnotation,
  WorkPdfCollaborationContent,
  WorkPdfCollaborationFormValue,
  WorkPdfCollaborationPageOperation,
  WorkPdfCollaborationRedactionProposal,
  WorkPdfCollaborationReviewDecision,
  WorkPdfCollaborationSignaturePlacement,
  WorkPdfCollaborationSource,
} from './office-pdf-collaboration-types';
import {
  invalidWorkOfficePdfShared as invalidSharedPdf,
  validateSharedWorkOfficePdfCollaborationContent,
} from './office-pdf-collaboration-validation';

export { validateWorkOfficePdfCollaborationContent } from './office-pdf-collaboration-validation';

export const PDF_SOURCE_ROOT = 'pdf.source';
export const PDF_SOURCE_IDENTITIES_ROOT = 'pdf.source-identities';
export const PDF_RECORD_CLAIMS_ROOT = 'pdf.record-claims';

export interface WorkOfficePdfRoots {
  source: Y.Map<unknown>;
  sourceIdentities: Y.Array<string>;
  annotations: WorkOfficePdfRecordCollectionRoots;
  formValues: WorkOfficePdfRecordCollectionRoots;
  signaturePlacements: WorkOfficePdfRecordCollectionRoots;
  redactionProposals: WorkOfficePdfRecordCollectionRoots;
  pageOperations: WorkOfficePdfRecordCollectionRoots;
  reviewDecisions: WorkOfficePdfRecordCollectionRoots;
  recordClaims: Y.Array<string>;
}

export function workOfficePdfRoots(
  document: Y.Doc,
  rootName: (suffix: string) => string,
): WorkOfficePdfRoots {
  return {
    source: document.getMap(rootName(PDF_SOURCE_ROOT)),
    sourceIdentities: document.getArray(rootName(PDF_SOURCE_IDENTITIES_ROOT)),
    annotations: workOfficePdfRecordCollectionRoots(
      document,
      rootName,
      'annotations',
    ),
    formValues: workOfficePdfRecordCollectionRoots(
      document,
      rootName,
      'form-values',
    ),
    signaturePlacements: workOfficePdfRecordCollectionRoots(
      document,
      rootName,
      'signature-placements',
    ),
    redactionProposals: workOfficePdfRecordCollectionRoots(
      document,
      rootName,
      'redaction-proposals',
    ),
    pageOperations: workOfficePdfRecordCollectionRoots(
      document,
      rootName,
      'page-operations',
    ),
    reviewDecisions: workOfficePdfRecordCollectionRoots(
      document,
      rootName,
      'review-decisions',
    ),
    recordClaims: document.getArray(rootName(PDF_RECORD_CLAIMS_ROOT)),
  };
}

export function initializeWorkOfficePdfRoots(
  roots: WorkOfficePdfRoots,
  content: WorkPdfCollaborationContent,
): void {
  patchSource(roots.source, content.source);
  roots.sourceIdentities.push([canonicalJson(content.source)]);
  appendWorkOfficePdfRecordClaims(roots.recordClaims, undefined, content);
  patchCollections(roots, undefined, content);
}

export function readWorkOfficePdfRoots(
  roots: WorkOfficePdfRoots,
): WorkPdfCollaborationContent {
  const content: WorkPdfCollaborationContent = {
    type: 'pdf',
    source: readSource(roots.source),
    annotations: readWorkOfficePdfRecords<
      WorkPdfCollaborationAnnotation & Record<string, unknown>
    >(roots.annotations, 'annotation'),
    formValues: readWorkOfficePdfRecords<
      WorkPdfCollaborationFormValue & Record<string, unknown>
    >(roots.formValues, 'form value'),
    signaturePlacements: readWorkOfficePdfRecords<
      WorkPdfCollaborationSignaturePlacement & Record<string, unknown>
    >(roots.signaturePlacements, 'signature placement'),
    redactionProposals: readWorkOfficePdfRecords<
      WorkPdfCollaborationRedactionProposal & Record<string, unknown>
    >(roots.redactionProposals, 'redaction proposal'),
    pageOperations: readWorkOfficePdfRecords<
      WorkPdfCollaborationPageOperation & Record<string, unknown>
    >(roots.pageOperations, 'page operation'),
    reviewDecisions: readWorkOfficePdfRecords<
      WorkPdfCollaborationReviewDecision & Record<string, unknown>
    >(roots.reviewDecisions, 'review decision'),
  };
  const validated = validateSharedWorkOfficePdfCollaborationContent(content);
  assertSourceIdentity(roots.sourceIdentities, validated.source);
  assertWorkOfficePdfRecordClaims(roots.recordClaims, validated);
  return validated;
}

export function patchWorkOfficePdfRoots(
  roots: WorkOfficePdfRoots,
  previous: WorkPdfCollaborationContent,
  next: WorkPdfCollaborationContent,
): void {
  const shared = readWorkOfficePdfRoots(roots);
  assertWorkOfficePdfPatchSafe(previous, next, shared);
  appendWorkOfficePdfRecordClaims(roots.recordClaims, previous, next);
  patchCollections(roots, previous, next);
}

export function assertWorkOfficePdfRootsEmpty(roots: WorkOfficePdfRoots): void {
  if (
    roots.source.size > 0 ||
    roots.sourceIdentities.length > 0 ||
    !assertWorkOfficePdfRecordCollectionEmpty(roots.annotations) ||
    !assertWorkOfficePdfRecordCollectionEmpty(roots.formValues) ||
    !assertWorkOfficePdfRecordCollectionEmpty(roots.signaturePlacements) ||
    !assertWorkOfficePdfRecordCollectionEmpty(roots.redactionProposals) ||
    !assertWorkOfficePdfRecordCollectionEmpty(roots.pageOperations) ||
    !assertWorkOfficePdfRecordCollectionEmpty(roots.reviewDecisions) ||
    roots.recordClaims.length > 0
  ) {
    throw new WorkOfficeCollaborationError(
      'office.collaboration.bootstrap_ambiguous',
      'The PDF collaboration roots contain data without initialized metadata.',
    );
  }
}

export function workOfficePdfUndoScope(
  roots: WorkOfficePdfRoots,
): Array<Y.Map<unknown> | Y.Array<string>> {
  return [
    ...workOfficePdfRecordCollectionUndoScope(roots.annotations),
    ...workOfficePdfRecordCollectionUndoScope(roots.formValues),
  ];
}

export function workOfficePdfIrreversibleScope(
  roots: WorkOfficePdfRoots,
): Array<Y.Map<unknown> | Y.Array<string>> {
  return [
    ...workOfficePdfRecordCollectionUndoScope(roots.signaturePlacements),
    ...workOfficePdfRecordCollectionUndoScope(roots.redactionProposals),
    ...workOfficePdfRecordCollectionUndoScope(roots.pageOperations),
    ...workOfficePdfRecordCollectionUndoScope(roots.reviewDecisions),
  ];
}

export function workOfficePdfIrreversibleChange(
  roots: WorkOfficePdfRoots,
  transaction: Y.Transaction,
): boolean {
  const changedParents = transaction.changedParentTypes;
  return (
    workOfficePdfRecordFieldChanged(transaction, roots.annotations, [
      'deleted',
    ]) ||
    workOfficePdfIrreversibleScope(roots).some((root) =>
      changedParents.has(
        root as unknown as Y.AbstractType<Y.YEvent<Y.AbstractType<unknown>>>,
      ),
    )
  );
}

function patchCollections(
  roots: WorkOfficePdfRoots,
  previous: WorkPdfCollaborationContent | undefined,
  next: WorkPdfCollaborationContent,
): void {
  patchWorkOfficePdfRecords(
    roots.annotations,
    previous?.annotations ?? [],
    next.annotations,
    'annotation',
  );
  patchWorkOfficePdfRecords(
    roots.formValues,
    previous?.formValues ?? [],
    next.formValues,
    'form value',
  );
  patchWorkOfficePdfRecords(
    roots.signaturePlacements,
    previous?.signaturePlacements ?? [],
    next.signaturePlacements,
    'signature placement',
  );
  patchWorkOfficePdfRecords(
    roots.redactionProposals,
    previous?.redactionProposals ?? [],
    next.redactionProposals,
    'redaction proposal',
  );
  patchWorkOfficePdfRecords(
    roots.pageOperations,
    previous?.pageOperations ?? [],
    next.pageOperations,
    'page operation',
  );
  patchWorkOfficePdfRecords(
    roots.reviewDecisions,
    previous?.reviewDecisions ?? [],
    next.reviewDecisions,
    'review decision',
  );
}

function patchSource(
  target: Y.Map<unknown>,
  source: WorkPdfCollaborationSource,
): void {
  target.set('sha256', source.sha256);
  target.set('byteLength', source.byteLength);
  target.set('pageCount', source.pageCount);
}

function readSource(target: Y.Map<unknown>): WorkPdfCollaborationSource {
  const sha256 = target.get('sha256');
  const byteLength = target.get('byteLength');
  const pageCount = target.get('pageCount');
  if (
    target.size !== 3 ||
    typeof sha256 !== 'string' ||
    typeof byteLength !== 'number' ||
    typeof pageCount !== 'number'
  ) {
    invalidSharedPdf('source identity');
  }
  return { sha256, byteLength, pageCount };
}

function assertSourceIdentity(
  identities: Y.Array<string>,
  source: WorkPdfCollaborationSource,
): void {
  if (identities.length === 0) return;
  if (identities.length !== 1 || identities.get(0) !== canonicalJson(source)) {
    invalidSharedPdf('immutable source identity');
  }
}
