import type { PdfCollaborationContent } from '../../src/core';

export const PDF_COLLABORATION_SOURCE = {
  sha256: '0123456789abcdef0123456789abcdef0123456789abcdef0123456789abcdef',
  byteLength: 4_096,
  pageCount: 3,
} as const;

export function pdfCollaborationFixture(): PdfCollaborationContent {
  return {
    type: 'pdf',
    source: { ...PDF_COLLABORATION_SOURCE },
    annotations: [
      {
        id: 'annotation-base-1',
        pageIndex: 0,
        source: 'base',
        annotation: {
          id: 'annotation-base-1',
          pageIndex: 0,
          type: 9,
          rect: { left: 10, top: 20, right: 120, bottom: 36 },
          color: '#ffff00',
          opacity: 0.5,
          contents: 'Imported highlight',
        },
      },
    ],
    formValues: [{ id: 'Applicant.Name', value: 'Ada' }],
    signaturePlacements: [],
    redactionProposals: [],
    pageOperations: [],
    reviewDecisions: [],
  };
}
