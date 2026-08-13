export interface WorkPdfCollaborationSource {
  /** Lowercase SHA-256 digest of the immutable source bytes. */
  sha256: string;
  byteLength: number;
  pageCount: number;
}

export interface WorkPdfCollaborationAnnotation {
  id: string;
  pageIndex: number;
  source: 'base' | 'created';
  /** Portable, JSON-compatible annotation object including matching id/pageIndex. */
  annotation: Record<string, unknown>;
  /** Durable tombstone. Annotation records are never physically removed. */
  deleted?: true;
}

export interface WorkPdfCollaborationFormValue {
  /** Stable fully-qualified PDF field name. */
  id: string;
  value: string;
}

export interface WorkPdfCollaborationSignatureAppearance {
  /** Host-owned encrypted/private asset reference; signature bytes are not in Yjs. */
  assetId: string;
  sha256: string;
  byteLength: number;
  mimeType: string;
}

export interface WorkPdfCollaborationSignaturePlacement {
  id: string;
  annotationId: string;
  pageIndex: number;
  kind: 'signature' | 'initials';
  actorId: string;
  createdAt: string;
  appearance: WorkPdfCollaborationSignatureAppearance;
}

export interface WorkPdfCollaborationRect {
  left: number;
  top: number;
  right: number;
  bottom: number;
}

export interface WorkPdfCollaborationRedactionProposal {
  id: string;
  pageIndex: number;
  rects: WorkPdfCollaborationRect[];
  proposedBy: string;
  proposedAt: string;
  reason?: string;
  text?: string;
}

interface WorkPdfCollaborationPageOperationBase {
  id: string;
  proposedBy: string;
  proposedAt: string;
}

export interface WorkPdfCollaborationRotatePagesOperation
  extends WorkPdfCollaborationPageOperationBase {
  kind: 'rotate';
  pageIndices: number[];
  degrees: 90 | 180 | 270;
}

export interface WorkPdfCollaborationDeletePagesOperation
  extends WorkPdfCollaborationPageOperationBase {
  kind: 'delete';
  pageIndices: number[];
}

export interface WorkPdfCollaborationReorderPagesOperation
  extends WorkPdfCollaborationPageOperationBase {
  kind: 'reorder';
  /** Complete zero-based permutation of immutable source pages. */
  pageOrder: number[];
}

export type WorkPdfCollaborationPageOperation =
  | WorkPdfCollaborationRotatePagesOperation
  | WorkPdfCollaborationDeletePagesOperation
  | WorkPdfCollaborationReorderPagesOperation;

export interface WorkPdfCollaborationReviewDecision {
  id: string;
  targetKind: 'redaction' | 'page-operation';
  targetId: string;
  decision: 'approve' | 'reject';
  actorId: string;
  createdAt: string;
}

export interface WorkPdfCollaborationContent {
  type: 'pdf';
  source: WorkPdfCollaborationSource;
  annotations: WorkPdfCollaborationAnnotation[];
  formValues: WorkPdfCollaborationFormValue[];
  signaturePlacements: WorkPdfCollaborationSignaturePlacement[];
  redactionProposals: WorkPdfCollaborationRedactionProposal[];
  pageOperations: WorkPdfCollaborationPageOperation[];
  reviewDecisions: WorkPdfCollaborationReviewDecision[];
}

export function createWorkPdfCollaborationContent(
  source: WorkPdfCollaborationSource,
): WorkPdfCollaborationContent {
  return {
    type: 'pdf',
    source,
    annotations: [],
    formValues: [],
    signaturePlacements: [],
    redactionProposals: [],
    pageOperations: [],
    reviewDecisions: [],
  };
}
