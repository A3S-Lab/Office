import { describe, expect, test } from '@rstest/core';
import {
  cycleDocumentSelectionCaseText,
  detectDocumentSelectionCase,
  nextDocumentSelectionCase,
  transformDocumentSelectionCase,
} from '../src/internal/features/work/work-document-text-case';

describe('document selection case cycle', () => {
  test('detects lower, upper, and title case', () => {
    expect(detectDocumentSelectionCase('hello world')).toBe('lower');
    expect(detectDocumentSelectionCase('HELLO WORLD')).toBe('upper');
    expect(detectDocumentSelectionCase('Hello World')).toBe('title');
  });

  test('cycles lower → upper → title → lower', () => {
    expect(nextDocumentSelectionCase('lower')).toBe('upper');
    expect(nextDocumentSelectionCase('upper')).toBe('title');
    expect(nextDocumentSelectionCase('title')).toBe('lower');
    expect(cycleDocumentSelectionCaseText('hello world')).toBe('HELLO WORLD');
    expect(cycleDocumentSelectionCaseText('HELLO WORLD')).toBe('Hello World');
    expect(cycleDocumentSelectionCaseText('Hello World')).toBe('hello world');
  });

  test('transforms selected text for each case target', () => {
    expect(transformDocumentSelectionCase("o'reilly", 'title')).toBe("O'reilly");
    expect(transformDocumentSelectionCase('Hello', 'lower')).toBe('hello');
    expect(transformDocumentSelectionCase('Hello', 'upper')).toBe('HELLO');
  });
});
