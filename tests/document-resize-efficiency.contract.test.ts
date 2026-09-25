import { readFileSync } from 'node:fs';
import { expect, test } from '@rstest/core';

const commentsPanel = readFileSync(
  'src/internal/features/work/editors/document-comments-panel.tsx',
  'utf8',
);
const pagination = readFileSync(
  'src/internal/features/work/editors/use-document-pagination.ts',
  'utf8',
);

test('comment connectors keep mark class sync off the geometry measure path', () => {
  expect(commentsPanel).toContain('syncDocumentCommentMarkClasses');
  expect(commentsPanel).toContain('scheduleMeasureAfterWindowResize');
  // Observer must not tear down when the comment window slides.
  expect(commentsPanel).not.toMatch(
    /window\.addEventListener\('resize',\s*scheduleMeasure\)/,
  );
  expect(commentsPanel).toMatch(
    /}, \[\s*draft,\s*editor,\s*scheduleMeasure,\s*scheduleMeasureAfterMutations,\s*scheduleMeasureAfterWindowResize,\s*scheduleWindowUpdate,\s*surfaceRef,\s*\]/,
  );
});

test('document pagination coalesces window resize into a trailing settle', () => {
  expect(pagination).toContain('windowResizeTimer');
  expect(pagination).toContain('setTimeout(() => {');
  expect(pagination).toContain('paginationViewportTriggers');
});
