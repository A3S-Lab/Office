import { expect, test } from '@rstest/core';
import {
  documentationEntryUrl,
  legacyDocsPath,
} from '../playground/src/site-routes';

test('keeps the Playground at the site root and opens the documentation center', () => {
  expect(
    documentationEntryUrl('https://a3s-lab.github.io/Office/index.html'),
  ).toBe('https://a3s-lab.github.io/Office/docs/index.html');
  expect(documentationEntryUrl('http://127.0.0.1:4175/')).toBe(
    'http://127.0.0.1:4175/docs/index.html',
  );
});

test('keeps legacy guide hashes on dedicated documentation routes', () => {
  expect(legacyDocsPath('#guide')).toBe('docs/guide/index.html');
  expect(legacyDocsPath('#guide/components')).toBe(
    'docs/components/index.html',
  );
  expect(legacyDocsPath('#guide/api')).toBe('docs/components/document.html');
  expect(legacyDocsPath('#guide/cli')).toBe('docs/automation/index.html');
  expect(legacyDocsPath('#unknown')).toBeNull();
});
