import { expect, test } from '@rstest/core';
import { createDocumentPaginationMeasurementRange } from '../src/internal/features/work/editors/document-pagination-measurement-range';

test('restores the consumed prefix when cooperative pagination is aborted', () => {
  const range = createDocumentPaginationMeasurementRange();
  const initialPass = range.begin();
  expect(initialPass.from).toBe(0);

  range.invalidate(900);
  range.restore(initialPass);

  expect(range.begin().from).toBe(0);
});

test('keeps later invalidations after a pagination pass commits', () => {
  const range = createDocumentPaginationMeasurementRange();
  const initialPass = range.begin();
  range.invalidate(900);
  range.commit(initialPass);

  expect(range.begin().from).toBe(900);
});

test('marks an externally revised document dirty when no edit was observed', () => {
  const range = createDocumentPaginationMeasurementRange();
  const initialPass = range.begin();
  range.commit(initialPass);
  range.ensureDirty();

  expect(range.begin().from).toBe(0);
});
