import { describe, expect, test } from '@rstest/core';
import {
  combineDocumentHtmlFingerprintSegments,
  createDocumentHtmlFingerprintSegment,
  documentHtmlFingerprint,
  documentHtmlFingerprintForSegment,
  documentHtmlFingerprintMatches,
} from '../src/internal/features/work/work-document-html-fingerprint';

describe('document HTML fingerprint', () => {
  test('combines independently hashed segments into the exact full fingerprint', () => {
    const parts = ['<section>', '<p>one</p>', '<p>two</p>', '</section>'];
    const combined = combineDocumentHtmlFingerprintSegments(
      parts.map((part) => createDocumentHtmlFingerprintSegment(part)),
    );

    expect(documentHtmlFingerprintForSegment(combined)).toBe(
      documentHtmlFingerprint(parts.join('')),
    );
    expect(documentHtmlFingerprint(parts.join(''))).toMatch(/^p1:/);
    expect(documentHtmlFingerprint(parts.join(''))).not.toBe(
      documentHtmlFingerprint([...parts].reverse().join('')),
    );
  });

  test('accepts persisted legacy fingerprints while rejecting stale content', () => {
    const html = '<p>Structured content</p>';

    expect(documentHtmlFingerprintMatches(html, 'p:1ht88x0')).toBe(true);
    expect(
      documentHtmlFingerprintMatches('<p>Changed content</p>', 'p:1ht88x0'),
    ).toBe(false);
    expect(
      documentHtmlFingerprintMatches(html, documentHtmlFingerprint(html)),
    ).toBe(true);
  });
});
