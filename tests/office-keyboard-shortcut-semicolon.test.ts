import { expect, test } from '@rstest/core';
import { matchesOfficeEditorKeyboardShortcut } from '../src/internal/keyboard-shortcuts';

test('matches shifted semicolon shortcuts by physical key code', () => {
  const event = new KeyboardEvent('keydown', {
    code: 'Semicolon',
    ctrlKey: true,
    key: ':',
    shiftKey: true,
  });

  expect(matchesOfficeEditorKeyboardShortcut(event, 'Control-Shift-;')).toBe(
    true,
  );
});

test('matches shifted and unshifted apostrophe shortcuts by physical key code', () => {
  const formula = new KeyboardEvent('keydown', {
    code: 'Quote',
    ctrlKey: true,
    key: "'",
  });
  const value = new KeyboardEvent('keydown', {
    code: 'Quote',
    ctrlKey: true,
    key: '"',
    shiftKey: true,
  });

  expect(matchesOfficeEditorKeyboardShortcut(formula, "Control-'")).toBe(true);
  expect(matchesOfficeEditorKeyboardShortcut(value, "Control-Shift-'")).toBe(
    true,
  );
});
