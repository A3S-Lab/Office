import { expect, test } from '@rstest/core';
import { createOfficeEditorRuntime } from '../src/internal/features/work/editors/office-editor-extension';
import type {
  SpreadsheetCommandContext,
  SpreadsheetEditorCommands,
  SpreadsheetNavigationCommandPort,
} from '../src/internal/features/work/editors/spreadsheet-command-controller';
import { createSpreadsheetNavigationExtension } from '../src/internal/features/work/editors/spreadsheet-navigation-command';

test('routes WPS Find and Go To commands through one typed navigation port', () => {
  const fixture = navigationFixture();
  const editor = createOfficeEditorRuntime<
    SpreadsheetCommandContext,
    SpreadsheetEditorCommands
  >(fixture.context, [createSpreadsheetNavigationExtension()]);

  expect(editor.can().openFind()).toBe(true);
  expect(editor.commands.openFind()).toBe(true);
  expect(editor.can().openGoTo()).toBe(true);
  expect(editor.commands.openGoTo()).toBe(true);
  expect(fixture.calls).toEqual(['find', 'go-to']);

  fixture.navigation.canOpenGoTo = false;
  expect(editor.can().openGoTo()).toBe(false);
  expect(editor.commands.openGoTo()).toBe(false);
  expect(fixture.calls).toEqual(['find', 'go-to']);
});

test('owns Ctrl+F, Ctrl+G, and F5 only on the live spreadsheet grid', () => {
  const fixture = navigationFixture();
  const editor = createOfficeEditorRuntime<
    SpreadsheetCommandContext,
    SpreadsheetEditorCommands
  >(fixture.context, [createSpreadsheetNavigationExtension()]);
  const container = document.createElement('div');
  container.className = 'fortune-container';
  const grid = document.createElement('div');
  grid.className = 'fortune-sheet-overlay';
  const formula = document.createElement('div');
  formula.className = 'fortune-fx-input';
  formula.contentEditable = 'true';
  container.append(grid, formula);
  const hostInput = document.createElement('input');
  document.body.append(container, hostInput);

  const dispatch = (target: HTMLElement, init: KeyboardEventInit) => {
    let handled = false;
    target.addEventListener(
      'keydown',
      (event) => {
        handled = editor.handleKeyDown(event);
      },
      { once: true },
    );
    const event = new KeyboardEvent('keydown', {
      bubbles: true,
      cancelable: true,
      ...init,
    });
    target.dispatchEvent(event);
    return { event, handled };
  };

  for (const init of [
    { ctrlKey: true, key: 'f' },
    { ctrlKey: true, key: 'g' },
    { key: 'F5' },
  ]) {
    const result = dispatch(grid, init);
    expect(result.handled).toBe(true);
    expect(result.event.defaultPrevented).toBe(true);
  }
  expect(fixture.calls).toEqual(['find', 'go-to', 'go-to']);

  for (const target of [hostInput, formula]) {
    const result = dispatch(target, { ctrlKey: true, key: 'g' });
    expect(result.handled).toBe(false);
    expect(result.event.defaultPrevented).toBe(false);
  }
  const repeat = dispatch(grid, { ctrlKey: true, key: 'g', repeat: true });
  expect(repeat.handled).toBe(false);
  expect(fixture.calls).toEqual(['find', 'go-to', 'go-to']);

  hostInput.remove();
  container.remove();
});

function navigationFixture(): {
  calls: string[];
  context: SpreadsheetCommandContext;
  navigation: SpreadsheetNavigationCommandPort;
} {
  const calls: string[] = [];
  const navigation: SpreadsheetNavigationCommandPort = {
    canOpenFind: true,
    canOpenGoTo: true,
    openFind: () => {
      calls.push('find');
      return true;
    },
    openGoTo: () => {
      calls.push('go-to');
      return true;
    },
  };
  return {
    calls,
    navigation,
    context: {
      activeSheetId: 'sheet-1',
      autoFilter: {
        active: false,
        canOpenMenu: false,
        canToggle: false,
        openMenu: () => false,
        toggle: () => false,
      },
      calculation: null,
      clipboard: {
        canCopySelection: false,
        canCutSelection: false,
        canOpenPasteSpecial: false,
        canPasteSelection: false,
        canPasteSpecial: () => false,
        copySelection: () => false,
        cutSelection: () => false,
        openPasteSpecial: () => false,
        pasteSelection: () => false,
        pasteSpecial: () => false,
      },
      content: {
        type: 'spreadsheet',
        sheets: [{ id: 'sheet-1', name: 'Sheet 1', status: 1 }],
      },
      editable: true,
      fallbackRange: { row: [0, 0], column: [0, 0] },
      formulaBar: null,
      formatCells: { canOpen: false, open: () => false },
      formatPainter: {
        active: false,
        canActivate: false,
        mode: null,
        activate: () => false,
        applySelection: () => false,
        cancel: () => false,
      },
      history: null,
      navigation,
      onChange: () => undefined,
      selection: null,
      targetSheetId: 'sheet-1',
      toolbarCell: null,
      view: null,
      workbook: null,
    },
  };
}
