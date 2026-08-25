import { expect, test } from '@rstest/core';
import {
  collaborationServerDocumentationUrl,
  documentationEntryUrl,
  legacyDocsPath,
} from '../playground/src/site-routes';

test('keeps documentation at the site root and resolves it from Playground', () => {
  expect(
    documentationEntryUrl(
      'https://a3s-lab.github.io/Office/playground/index.html',
    ),
  ).toBe('https://a3s-lab.github.io/Office/');
  expect(documentationEntryUrl('http://127.0.0.1:4175/playground/')).toBe(
    'http://127.0.0.1:4175/',
  );
  expect(
    collaborationServerDocumentationUrl(
      'https://a3s-lab.github.io/Office/playground/',
    ),
  ).toBe(
    'https://a3s-lab.github.io/Office/components/collaboration-server.html',
  );
});

test('keeps legacy guide hashes on dedicated documentation routes', () => {
  expect(legacyDocsPath('#guide')).toBe('../guide/');
  expect(legacyDocsPath('#guide/components')).toBe('../components/');
  expect(legacyDocsPath('#guide/api')).toBe('../components/document.html');
  expect(legacyDocsPath('#guide/cli')).toBe('../automation/');
  expect(legacyDocsPath('#unknown')).toBeNull();
});
