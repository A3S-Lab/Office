import { expect, test } from '@rstest/core';
import {
  nextPresentationBlankScreen,
  presentationBlankScreenLabel,
} from '../src/internal/features/work/editors/presentation-slideshow-blank';

test('toggles WPS slideshow black and white blank screens', () => {
  expect(nextPresentationBlankScreen('off', 'b')).toBe('black');
  expect(nextPresentationBlankScreen('black', 'B')).toBe('off');
  expect(nextPresentationBlankScreen('off', '.')).toBe('black');
  expect(nextPresentationBlankScreen('black', '.')).toBe('off');

  expect(nextPresentationBlankScreen('off', 'w')).toBe('white');
  expect(nextPresentationBlankScreen('white', 'W')).toBe('off');
  expect(nextPresentationBlankScreen('off', ',')).toBe('white');
  expect(nextPresentationBlankScreen('white', ',')).toBe('off');

  expect(nextPresentationBlankScreen('black', 'w')).toBe('white');
  expect(nextPresentationBlankScreen('white', 'b')).toBe('black');
  expect(nextPresentationBlankScreen('off', 'ArrowRight')).toBeNull();
  expect(presentationBlankScreenLabel('black')).toBe('黑屏');
  expect(presentationBlankScreenLabel('white')).toBe('白屏');
});
