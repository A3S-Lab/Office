import { expect, test } from '@rstest/core';
import { waitFor } from '@testing-library/react';
import type { SpreadsheetEditorCommands } from '../src/internal/features/work/editors/spreadsheet-command-controller';
import {
  focusSpreadsheetGrid,
  spreadsheetCommandsWithGridFocus,
} from '../src/internal/features/work/editors/spreadsheet-editor';
import {
  isSpreadsheetCellEditingTarget,
  isSpreadsheetNativeTextUndoTarget,
} from '../src/internal/features/work/editors/spreadsheet-editor-support';

test('recognizes live cell and formula editors before restoring grid focus', () => {
  const fortune = document.createElement('div');
  fortune.className = 'fortune-container';
  const inputBox = document.createElement('div');
  inputBox.className = 'luckysheet-input-box';
  inputBox.style.zIndex = '19';
  const cellInput = document.createElement('div');
  cellInput.className = 'luckysheet-cell-input';
  cellInput.contentEditable = 'true';
  inputBox.append(cellInput);
  const hiddenInputBox = document.createElement('div');
  hiddenInputBox.className = 'luckysheet-input-box';
  hiddenInputBox.style.zIndex = '-1';
  const hiddenCellInput = document.createElement('div');
  hiddenCellInput.className = 'luckysheet-cell-input';
  hiddenCellInput.contentEditable = 'true';
  hiddenInputBox.append(hiddenCellInput);
  const formulaInput = document.createElement('div');
  formulaInput.className = 'fortune-fx-input';
  formulaInput.contentEditable = 'true';
  const grid = document.createElement('div');
  grid.className = 'fortune-sheet-overlay';
  const unrelatedInput = document.createElement('input');
  fortune.append(inputBox, hiddenInputBox, formulaInput, grid);

  expect(isSpreadsheetCellEditingTarget(cellInput)).toBe(true);
  expect(isSpreadsheetCellEditingTarget(hiddenCellInput)).toBe(false);
  expect(isSpreadsheetCellEditingTarget(formulaInput)).toBe(true);
  expect(isSpreadsheetCellEditingTarget(unrelatedInput)).toBe(false);
  expect(isSpreadsheetNativeTextUndoTarget(cellInput)).toBe(true);
  expect(isSpreadsheetNativeTextUndoTarget(hiddenCellInput)).toBe(false);
  expect(isSpreadsheetNativeTextUndoTarget(formulaInput)).toBe(true);
  expect(isSpreadsheetNativeTextUndoTarget(grid)).toBe(false);
  expect(isSpreadsheetNativeTextUndoTarget(unrelatedInput)).toBe(true);
});

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

  await waitFor(() => expect(document.activeElement).toBe(overlay));
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

test('does not steal focus when a hidden cell editor resumes editing', async () => {
  const container = document.createElement('div');
  const overlay = document.createElement('main');
  overlay.className = 'fortune-sheet-overlay';
  overlay.tabIndex = -1;
  const inputBox = document.createElement('div');
  inputBox.className = 'luckysheet-input-box';
  inputBox.style.zIndex = '-1';
  const cellInput = document.createElement('div');
  cellInput.className = 'luckysheet-cell-input';
  cellInput.contentEditable = 'true';
  cellInput.tabIndex = 0;
  inputBox.append(cellInput);
  container.append(overlay, inputBox);
  document.body.append(container);
  cellInput.focus();

  focusSpreadsheetGrid(container);
  expect(document.activeElement).toBe(overlay);

  inputBox.style.zIndex = '19';
  cellInput.focus();
  await waitForAnimationFrames(3);

  expect(document.activeElement).toBe(cellInput);
  container.remove();
});

test('returns grid focus after successful ribbon commands only', () => {
  const calls: string[] = [];
  const record =
    (name: string, result = true) =>
    (...args: unknown[]) => {
      calls.push(`${name}:${args.join(',')}`);
      return result;
    };
  const commands = {
    activateFormatPainter: record('activateFormatPainter'),
    applyFormatPainter: record('applyFormatPainter'),
    cancelFormatPainter: record('cancelFormatPainter'),
    copySelection: record('copySelection'),
    cutSelection: record('cutSelection', false),
    openAutoFilterMenu: record('openAutoFilterMenu'),
    pasteSelection: record('pasteSelection'),
    redo: record('redo'),
    setCellFormat: record('setCellFormat'),
    setGridLines: record('setGridLines'),
    setZoom: record('setZoom'),
    toggleCellMerge: record('toggleCellMerge'),
    toggleAutoFilter: record('toggleAutoFilter'),
    undo: record('undo', false),
  } as unknown as SpreadsheetEditorCommands;
  const focused: string[] = [];
  const ribbon = spreadsheetCommandsWithGridFocus(commands, () =>
    focused.push('grid'),
  );

  expect(ribbon.setCellFormat('bl', 1)).toBe(true);
  expect(ribbon.setGridLines(false)).toBe(true);
  expect(ribbon.toggleAutoFilter()).toBe(true);
  expect(ribbon.openAutoFilterMenu()).toBe(true);
  expect(ribbon.activateFormatPainter('once')).toBe(true);
  expect(ribbon.cancelFormatPainter()).toBe(true);
  expect(ribbon.copySelection()).toBe(true);
  expect(ribbon.cutSelection()).toBe(false);
  expect(ribbon.pasteSelection()).toBe(true);
  expect(ribbon.undo()).toBe(false);
  expect(ribbon.setZoom(125)).toBe(true);

  expect(calls).toEqual([
    'setCellFormat:bl,1',
    'setGridLines:false',
    'toggleAutoFilter:',
    'openAutoFilterMenu:',
    'activateFormatPainter:once',
    'cancelFormatPainter:',
    'copySelection:',
    'cutSelection:',
    'pasteSelection:',
    'undo:',
    'setZoom:125',
  ]);
  expect(focused).toEqual([
    'grid',
    'grid',
    'grid',
    'grid',
    'grid',
    'grid',
    'grid',
  ]);
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
