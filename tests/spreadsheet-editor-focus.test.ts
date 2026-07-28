import { expect, test } from '@rstest/core';
import { focusSpreadsheetGrid } from '../src/internal/features/work/editors/spreadsheet-editor';

test('restores focus to the interactive spreadsheet overlay', async () => {
  const container = document.createElement('div');
  const cellArea = document.createElement('div');
  cellArea.className = 'fortune-cell-area';
  cellArea.tabIndex = -1;
  const overlay = document.createElement('main');
  overlay.className = 'fortune-sheet-overlay';
  overlay.tabIndex = -1;
  container.append(cellArea, overlay);
  document.body.append(container);

  focusSpreadsheetGrid(container);
  expect(document.activeElement).toBe(overlay);

  await waitForAnimationFrames(3);
  const remountedOverlay = document.createElement('main');
  remountedOverlay.className = 'fortune-sheet-overlay';
  remountedOverlay.tabIndex = -1;
  overlay.replaceWith(remountedOverlay);
  await waitForAnimationFrames(3);

  expect(document.activeElement).toBe(remountedOverlay);
  container.remove();
});

test('restores grid focus when a worksheet command trigger regains focus during remount', async () => {
  const trigger = document.createElement('button');
  const container = document.createElement('div');
  const overlay = document.createElement('main');
  overlay.className = 'fortune-sheet-overlay';
  overlay.tabIndex = -1;
  container.append(overlay);
  document.body.append(trigger, container);
  trigger.focus();

  focusSpreadsheetGrid(container);
  expect(document.activeElement).toBe(overlay);

  const remountedOverlay = document.createElement('main');
  remountedOverlay.className = 'fortune-sheet-overlay';
  remountedOverlay.tabIndex = -1;
  overlay.replaceWith(remountedOverlay);
  trigger.focus();
  await waitForAnimationFrames(3);

  expect(document.activeElement).toBe(remountedOverlay);
  trigger.remove();
  container.remove();
});

test('waits for a lazily mounted spreadsheet overlay before restoring focus', async () => {
  const container = document.createElement('div');
  document.body.append(container);

  focusSpreadsheetGrid(container);
  await waitForAnimationFrames(15);

  const overlay = document.createElement('main');
  overlay.className = 'fortune-sheet-overlay';
  overlay.tabIndex = -1;
  container.append(overlay);
  await waitForAnimationFrames(2);

  expect(document.activeElement).toBe(overlay);
  container.remove();
});

test('recovers from a delayed workbook blur after the grid is mounted', async () => {
  const container = document.createElement('div');
  const overlay = document.createElement('main');
  overlay.className = 'fortune-sheet-overlay';
  overlay.tabIndex = -1;
  container.append(overlay);
  document.body.append(container);

  focusSpreadsheetGrid(container);
  await waitForAnimationFrames(15);
  overlay.blur();
  expect(document.activeElement).toBe(document.body);
  await waitForAnimationFrames(2);

  expect(document.activeElement).toBe(overlay);
  container.remove();
});

function waitForAnimationFrames(count: number): Promise<void> {
  return new Promise((resolve) => {
    const next = (remaining: number) => {
      if (remaining <= 0) {
        resolve();
        return;
      }
      requestAnimationFrame(() => next(remaining - 1));
    };
    next(count);
  });
}
