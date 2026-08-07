import { expect, test } from '@rstest/core';
import {
  clampDocumentZoom,
  documentZoomForFit,
} from '../src/internal/features/work/editors/document-zoom';

test('clamps document zoom to the supported range', () => {
  expect(clampDocumentZoom(24)).toBe(50);
  expect(clampDocumentZoom(125.6)).toBe(126);
  expect(clampDocumentZoom(240)).toBe(200);
});

test('fits page width to the available viewport width', () => {
  expect(
    documentZoomForFit('width', {
      pageHeight: 1123,
      pageWidth: 794,
      viewportHeight: 800,
      viewportWidth: 850,
      viewportPadding: { top: 28, right: 28, bottom: 70, left: 28 },
    }),
  ).toBe(100);
});

test('fits a whole page using the limiting viewport dimension', () => {
  expect(
    documentZoomForFit('page', {
      pageHeight: 1123,
      pageWidth: 794,
      viewportHeight: 800,
      viewportWidth: 1200,
      viewportPadding: { top: 28, right: 28, bottom: 70, left: 28 },
    }),
  ).toBe(63);
});
