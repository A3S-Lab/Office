import { readdirSync, readFileSync, statSync } from 'node:fs';
import { join } from 'node:path';
import { expect, test } from '@rstest/core';

// Every raw-HTML sink renders markup that has already passed an allow-list
// sanitizer (`sanitizeDocumentHtml` or `sanitizeDocumentPageChromeHtml`) or
// comes from ProseMirror-rendered editor DOM. Adding a sink requires that
// review, then an entry here.
const REVIEWED_SINKS: Record<string, number> = {
  'src/internal/features/work/components/work-document-pages.tsx': 5,
  'src/internal/features/work/editors/document-editor.tsx': 2,
  'src/internal/features/work/work-document-pagination-decorations.ts': 3,
};

const SINK = /dangerouslySetInnerHTML|insertAdjacentHTML\(/g;

function sourceFiles(directory: string): string[] {
  return readdirSync(directory).flatMap((entry) => {
    const path = join(directory, entry);
    if (statSync(path).isDirectory()) return sourceFiles(path);
    return /\.(ts|tsx)$/.test(entry) && !/\.test\.tsx?$/.test(entry)
      ? [path]
      : [];
  });
}

test('raw HTML sinks stay limited to reviewed, sanitized call sites', () => {
  const found: Record<string, number> = {};
  for (const file of sourceFiles('src')) {
    const count = readFileSync(file, 'utf8').match(SINK)?.length ?? 0;
    if (count) found[file] = count;
  }
  expect(found).toEqual(REVIEWED_SINKS);
});

test('rendered Writer pages use the allow-list sanitizer', () => {
  const pages = readFileSync(
    'src/internal/features/work/work-document-pages.ts',
    'utf8',
  );
  expect(pages).toContain('sanitizeDocumentHtml(');
  expect(pages).not.toMatch(/function sanitizeDocument\(/);
});
