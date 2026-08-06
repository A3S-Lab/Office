import { describe, expect, test } from '@rstest/core';
import { documentWordLineHeightFactor } from '../src/internal/features/work/work-document-word-line-metrics';

describe('Word-compatible document line metrics', () => {
  test.each([
    ['Arial', 1.15],
    ['Times New Roman', 1.15],
    ['Calibri', 1.2207],
    ['Segoe UI', 1.3301],
    ['Microsoft YaHei', 1.7143],
    ['SimSun', 1.2976],
    ['DengXian', 1.3548],
  ])('uses the measured WPS single-line factor for %s', (family, factor) => {
    expect(documentWordLineHeightFactor(family)).toBe(factor);
  });

  test('uses the first resolved CSS family and a stable fallback', () => {
    expect(documentWordLineHeightFactor('"Microsoft YaHei", Arial')).toBe(
      1.7143,
    );
    expect(documentWordLineHeightFactor('Unregistered Fixture Font')).toBe(
      1.15,
    );
  });
});
