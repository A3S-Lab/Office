import { expect, test } from '@rstest/core';
import {
  collaborationServerDocumentationUrl,
  documentationEntryUrl,
  legacyDocsPath,
} from '../playground/src/site-routes';

test('resolves the sibling documentation mount from Playground', () => {
  expect(
    documentationEntryUrl(
      'https://a3s-lab.github.io/Office/playground/index.html',
    ),
  ).toBe('https://a3s-lab.github.io/Office/docs/');
  expect(documentationEntryUrl('http://127.0.0.1:4175/playground/')).toBe(
    'http://127.0.0.1:4175/docs/',
  );
  expect(
    collaborationServerDocumentationUrl(
      'https://a3s-lab.github.io/Office/playground/',
    ),
  ).toBe(
    'https://a3s-lab.github.io/Office/docs/components/collaboration-server.html',
  );
});

test('keeps legacy guide hashes on dedicated documentation routes', () => {
  expect(legacyDocsPath('#guide')).toBe('../docs/guide/');
  expect(legacyDocsPath('#guide/components')).toBe('../docs/components/');
  expect(legacyDocsPath('#guide/api')).toBe('../docs/components/document.html');
  expect(legacyDocsPath('#guide/cli')).toBe('../docs/automation/');
  expect(legacyDocsPath('#unknown')).toBeNull();
});
