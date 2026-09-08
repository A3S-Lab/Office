import { expect, test } from '@rstest/core';
import {
  appendPresentationGotoDigit,
  presentationGotoEscapeAction,
  resolvePresentationGotoIndex,
} from '../src/internal/features/work/editors/presentation-slideshow-goto';

test('accumulates WPS slideshow digit buffer and clamps overlong input', () => {
  expect(appendPresentationGotoDigit('', '2')).toBe('2');
  expect(appendPresentationGotoDigit('2', '5')).toBe('25');
  expect(appendPresentationGotoDigit('1234', '5')).toBe('1234');
  expect(appendPresentationGotoDigit('', 'a')).toBeNull();
  expect(appendPresentationGotoDigit('', 'Enter')).toBeNull();
  expect(appendPresentationGotoDigit('', '.')).toBeNull();
});

test('resolves 1-based digit buffer to a clamped 0-based slide index', () => {
  expect(resolvePresentationGotoIndex('', 5)).toBeNull();
  expect(resolvePresentationGotoIndex('1', 5)).toBe(0);
  expect(resolvePresentationGotoIndex('3', 5)).toBe(2);
  expect(resolvePresentationGotoIndex('9', 5)).toBe(4);
  expect(resolvePresentationGotoIndex('0', 5)).toBeNull();
  expect(resolvePresentationGotoIndex('2', 0)).toBeNull();
});

test('Escape clears a pending digit buffer before exiting the slideshow', () => {
  expect(presentationGotoEscapeAction('12')).toBe('clear-buffer');
  expect(presentationGotoEscapeAction('')).toBe('exit');
});
