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
