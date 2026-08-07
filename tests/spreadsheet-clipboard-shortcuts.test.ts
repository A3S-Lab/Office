import { expect, test } from '@rstest/core';
import { runSpreadsheetClipboardShortcut } from '../src/internal/features/work/editors/spreadsheet-clipboard-shortcuts';

test('routes grid clipboard shortcuts through the typed command port', () => {
  const calls: string[] = [];
  for (const shortcut of [
    { key: 'c', modifier: 'meta', action: 'copy' },
    { key: 'x', modifier: 'control', action: 'cut' },
    { key: 'v', modifier: 'meta', action: 'paste' },
  ] as const) {
    const event = new KeyboardEvent('keydown', {
      cancelable: true,
      ctrlKey: shortcut.modifier === 'control',
      key: shortcut.key,
      metaKey: shortcut.modifier === 'meta',
    });
    expect(
      runSpreadsheetClipboardShortcut(
        event,
        () => true,
        () => {
          calls.push(shortcut.action);
          return true;
        },
      ),
    ).toBe(true);
    expect(event.defaultPrevented).toBe(true);
  }
  expect(calls).toEqual(['copy', 'cut', 'paste']);
});

test('leaves unavailable, native text, and modified shortcuts untouched', () => {
  const input = document.createElement('input');
  document.body.append(input);
  const unavailable = new KeyboardEvent('keydown', {
    cancelable: true,
    key: 'c',
    metaKey: true,
  });
  expect(
    runSpreadsheetClipboardShortcut(
      unavailable,
      () => false,
      () => true,
    ),
  ).toBe(false);
  expect(unavailable.defaultPrevented).toBe(false);

  let nativeHandled = true;
  input.addEventListener('keydown', (event) => {
    nativeHandled = runSpreadsheetClipboardShortcut(
      event,
      () => true,
      () => true,
    );
  });
  input.dispatchEvent(
    new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      key: 'x',
      ctrlKey: true,
    }),
  );
  expect(nativeHandled).toBe(false);

  const modified = new KeyboardEvent('keydown', {
    cancelable: true,
    key: 'v',
    metaKey: true,
    shiftKey: true,
  });
  expect(
    runSpreadsheetClipboardShortcut(
      modified,
      () => true,
      () => true,
    ),
  ).toBe(false);
  expect(modified.defaultPrevented).toBe(false);

  const ignoredSurface = document.createElement('div');
  ignoredSurface.dataset.officeShortcuts = 'ignore';
  const ignoredButton = document.createElement('button');
  ignoredSurface.append(ignoredButton);
  document.body.append(ignoredSurface);
  let ignoredHandled = true;
  ignoredButton.addEventListener('keydown', (event) => {
    ignoredHandled = runSpreadsheetClipboardShortcut(
      event,
      () => true,
      () => true,
    );
  });
  ignoredButton.dispatchEvent(
    new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      key: 'c',
      metaKey: true,
    }),
  );
  expect(ignoredHandled).toBe(false);
  ignoredSurface.remove();
  input.remove();
});

test('does not consume a shortcut when execution declines it', () => {
  const copy = new KeyboardEvent('keydown', {
    cancelable: true,
    key: 'c',
    metaKey: true,
  });
  expect(
    runSpreadsheetClipboardShortcut(
      copy,
      () => true,
      () => false,
    ),
  ).toBe(false);
  expect(copy.defaultPrevented).toBe(false);
});
