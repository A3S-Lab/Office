import { describe, expect, test } from '@rstest/core';
import { normalizeDocumentHref } from '../src/internal/features/work/work-document-links';

describe('document links', () => {
  test('accepts complete web, email, and document links', () => {
    expect(normalizeDocumentHref(' https://a3s.dev/docs ')).toBe(
      'https://a3s.dev/docs',
    );
    expect(normalizeDocumentHref('http://localhost:29653/')).toBe(
      'http://localhost:29653/',
    );
    expect(normalizeDocumentHref('mailto:hello@a3s.dev')).toBe(
      'mailto:hello@a3s.dev',
    );
    expect(normalizeDocumentHref('#section-2')).toBe('#section-2');
  });

  test('rejects incomplete and unsafe links', () => {
    expect(normalizeDocumentHref('https://')).toBeNull();
    expect(normalizeDocumentHref('javascript:alert(1)')).toBeNull();
    expect(normalizeDocumentHref('mailto:')).toBeNull();
    expect(normalizeDocumentHref('#')).toBeNull();
  });
});
