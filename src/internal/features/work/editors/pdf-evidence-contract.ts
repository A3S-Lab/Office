export const PDF_EVIDENCE_COORDINATE_BASIS = 1_000_000 as const;

export interface PdfEvidenceBounds {
  bottom: number;
  left: number;
  right: number;
  top: number;
}

export interface PdfEvidenceRegion {
  bounds: PdfEvidenceBounds;
  id: string;
  label?: string;
  sourceRegionIds: readonly string[];
  targetIds: readonly string[];
}

export interface PdfEvidencePage {
  canvasHeight: number;
  canvasId: string;
  canvasKind: 'source';
  canvasWidth: number;
  coordinateBasis: typeof PDF_EVIDENCE_COORDINATE_BASIS;
  pageNumber: number;
  regions: readonly PdfEvidenceRegion[];
  renderProfileSha256: string;
  rotationDegrees: 0 | 90 | 180 | 270;
  sourceSha256: string;
}

export interface PdfEvidenceRegionLocation {
  pageNumber: number;
}

export interface PdfEvidenceOverlay {
  coordinateBasis: typeof PDF_EVIDENCE_COORDINATE_BASIS;
  loadPage: (
    pageNumber: number,
    signal: AbortSignal,
  ) => Promise<PdfEvidencePage | null>;
  locateRegion?: (
    regionId: string,
    signal: AbortSignal,
  ) => Promise<PdfEvidenceRegionLocation | null>;
  renderProfileSha256: string;
  sourceSha256: string;
}
