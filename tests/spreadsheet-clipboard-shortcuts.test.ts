import { expect, test } from '@rstest/core';
import { waitFor } from '@testing-library/react';
import { runSpreadsheetClipboardShortcut } from '../src/internal/features/work/editors/spreadsheet-clipboard-shortcuts';

test('owns grid copy, cut, and paste shortcuts without clipboard permission', async () => {
  const calls: string[] = [];
  let clipboardText = '';
  const options = {
    clipboard: {
      readText: async () => clipboardText,
      writeText: async (value: string) => {
        clipboardText = value;
        calls.push(`copy:${value}`);
      },
    },
    clearSelectedCells: () => {
      calls.push('clear');
      return true;
    },
    pasteCells: (values: readonly (readonly unknown[])[]) => {
      calls.push(`paste:${JSON.stringify(values)}`);
      return true;
    },
    readSelectionText: () => 'A3S\tOffice',
    restoreFocus: () => calls.push('focus'),
  };

  const copy = new KeyboardEvent('keydown', {
    cancelable: true,
    key: 'c',
    metaKey: true,
  });
  expect(runSpreadsheetClipboardShortcut(copy, options)).toBe(true);
  expect(copy.defaultPrevented).toBe(true);
  await waitFor(() => expect(calls).toEqual(['copy:A3S\tOffice', 'focus']));

  calls.length = 0;
  const cut = new KeyboardEvent('keydown', {
    cancelable: true,
    ctrlKey: true,
    key: 'x',
  });
  expect(runSpreadsheetClipboardShortcut(cut, options)).toBe(true);
  expect(cut.defaultPrevented).toBe(true);
  await waitFor(() =>
    expect(calls).toEqual(['copy:A3S\tOffice', 'clear', 'focus']),
  );

  calls.length = 0;
  clipboardText = '项目\t金额\n研发\t120';
  const paste = new KeyboardEvent('keydown', {
    cancelable: true,
    key: 'v',
    metaKey: true,
  });
  expect(runSpreadsheetClipboardShortcut(paste, options)).toBe(true);
  expect(paste.defaultPrevented).toBe(true);
  await waitFor(() =>
    expect(calls).toEqual(['paste:[["项目","金额"],["研发","120"]]', 'focus']),
  );
});

test('leaves native text editing and modified shortcuts untouched', () => {
  const input = document.createElement('input');
  document.body.append(input);
  let handled = true;
  input.addEventListener('keydown', (event) => {
    handled = runSpreadsheetClipboardShortcut(event, {
      clipboard: {
        readText: async () => '',
        writeText: async () => undefined,
      },
      clearSelectedCells: () => true,
      pasteCells: () => true,
      readSelectionText: () => 'A3S',
      restoreFocus: () => undefined,
    });
  });
  const nativeCopy = new KeyboardEvent('keydown', {
    bubbles: true,
    cancelable: true,
    key: 'c',
    metaKey: true,
  });
  input.dispatchEvent(nativeCopy);
  expect(handled).toBe(false);
  expect(nativeCopy.defaultPrevented).toBe(false);

  const modifiedCopy = new KeyboardEvent('keydown', {
    cancelable: true,
    key: 'c',
    metaKey: true,
    shiftKey: true,
  });
  expect(
    runSpreadsheetClipboardShortcut(modifiedCopy, {
      clipboard: {
        readText: async () => '',
        writeText: async () => undefined,
      },
      clearSelectedCells: () => true,
      pasteCells: () => true,
      readSelectionText: () => 'A3S',
      restoreFocus: () => undefined,
    }),
  ).toBe(false);
  expect(modifiedCopy.defaultPrevented).toBe(false);

  const menu = document.createElement('div');
  menu.dataset.officeShortcuts = 'ignore';
  const menuItem = document.createElement('button');
  menu.append(menuItem);
  document.body.append(menu);
  let menuHandled = true;
  menuItem.addEventListener('keydown', (event) => {
    menuHandled = runSpreadsheetClipboardShortcut(event, {
      clipboard: {
        readText: async () => '',
        writeText: async () => undefined,
      },
      clearSelectedCells: () => true,
      pasteCells: () => true,
      readSelectionText: () => 'A3S',
      restoreFocus: () => undefined,
    });
  });
  const menuCopy = new KeyboardEvent('keydown', {
    bubbles: true,
    cancelable: true,
    key: 'c',
    metaKey: true,
  });
  menuItem.dispatchEvent(menuCopy);
  expect(menuHandled).toBe(false);
  expect(menuCopy.defaultPrevented).toBe(false);
  menu.remove();
  input.remove();
});

test('copies a blank selected cell instead of leaving stale local content', async () => {
  const writes: string[] = [];
  const copy = new KeyboardEvent('keydown', {
    cancelable: true,
    key: 'c',
    metaKey: true,
  });
  expect(
    runSpreadsheetClipboardShortcut(copy, {
      clipboard: {
        readText: async () => '',
        writeText: async (value) => {
          writes.push(value);
        },
      },
      clearSelectedCells: () => true,
      pasteCells: () => true,
      readSelectionText: () => '',
      restoreFocus: () => undefined,
    }),
  ).toBe(true);
  await waitFor(() => expect(writes).toEqual(['']));
});
