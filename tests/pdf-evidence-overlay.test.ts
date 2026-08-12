import { expect, test } from '@rstest/core';
import {
  PDF_EVIDENCE_COORDINATE_BASIS,
  type PdfEvidenceOverlay,
  type PdfEvidencePage,
  pdfEvidenceBoundsToPageRect,
  pdfEvidenceHorizontalCenterOffset,
  pdfEvidencePointToUnrotatedPage,
  validatePdfEvidencePage,
} from '../src/internal/features/work/editors/pdf-evidence-overlay';

const SOURCE_SHA256 = 'a'.repeat(64);
const RENDER_PROFILE_SHA256 = 'b'.repeat(64);

const overlay: PdfEvidenceOverlay = {
  coordinateBasis: PDF_EVIDENCE_COORDINATE_BASIS,
  loadPage: async () => null,
  renderProfileSha256: RENDER_PROFILE_SHA256,
  sourceSha256: SOURCE_SHA256,
};

const page: PdfEvidencePage = {
  canvasHeight: 400,
  canvasId: 'canvas-1',
  canvasKind: 'source',
  canvasWidth: 200,
  coordinateBasis: PDF_EVIDENCE_COORDINATE_BASIS,
  pageNumber: 1,
  regions: [
    {
      bounds: { bottom: 750_000, left: 100_000, right: 600_000, top: 250_000 },
      id: 'evidence-1',
      sourceRegionIds: ['source-region-1'],
      targetIds: ['node-1'],
    },
  ],
  renderProfileSha256: RENDER_PROFILE_SHA256,
  rotationDegrees: 0,
  sourceSha256: SOURCE_SHA256,
};

type EvidenceRuntime = Parameters<typeof validatePdfEvidencePage>[2];

function runtimeForPage(
  width = 200,
  height = 400,
  rotation: 0 | 1 | 2 | 3 = 0,
): EvidenceRuntime {
  return {
    documentManager: {
      getActiveDocument: () => ({
        pages: [{ rotation, size: { height, width } }],
      }),
    },
  } as unknown as EvidenceRuntime;
}

test('maps normalized evidence geometry onto the immutable page canvas', () => {
  expect(pdfEvidenceBoundsToPageRect(page.regions[0].bounds, 200, 400)).toEqual(
    {
      origin: { x: 20, y: 100 },
      size: { height: 200, width: 100 },
    },
  );

  expect(() =>
    validatePdfEvidencePage(page, overlay, runtimeForPage(), 1),
  ).not.toThrow();
});

test('includes EmbedPDF scroller centering when content is narrower than its viewport', () => {
  expect(pdfEvidenceHorizontalCenterOffset(806, 10, 595, 0.44)).toBeCloseTo(
    262.1,
  );
  expect(pdfEvidenceHorizontalCenterOffset(400, 10, 595, 1)).toBe(0);
});

test('restores rotated evidence points before using EmbedPDF page navigation', () => {
  expect(
    pdfEvidencePointToUnrotatedPage({ x: 100, y: 40 }, 200, 400, 0),
  ).toEqual({
    x: 100,
    y: 40,
  });
  expect(
    pdfEvidencePointToUnrotatedPage({ x: 300, y: 40 }, 200, 400, 1),
  ).toEqual({
    x: 40,
    y: 100,
  });
  expect(
    pdfEvidencePointToUnrotatedPage({ x: 100, y: 360 }, 200, 400, 2),
  ).toEqual({
    x: 100,
    y: 40,
  });
  expect(
    pdfEvidencePointToUnrotatedPage({ x: 300, y: 160 }, 200, 400, 3),
  ).toEqual({
    x: 40,
    y: 300,
  });
});

test('rejects stale identities, fabricated fine geometry, and missing source receipts', () => {
  expect(() =>
    validatePdfEvidencePage(
      { ...page, renderProfileSha256: 'c'.repeat(64) },
      overlay,
      runtimeForPage(),
      1,
    ),
  ).toThrow('identity is invalid');

  expect(() =>
    validatePdfEvidencePage(
      {
        ...page,
        regions: [
          {
            ...page.regions[0],
            bounds: { bottom: 200, left: 100, right: 100, top: 100 },
          },
        ],
      },
      overlay,
      runtimeForPage(),
      1,
    ),
  ).toThrow('bounds are invalid');

  expect(() =>
    validatePdfEvidencePage(
      {
        ...page,
        regions: [{ ...page.regions[0], sourceRegionIds: [] }],
      },
      overlay,
      runtimeForPage(),
      1,
    ),
  ).toThrow('identity receipts are invalid');
});
