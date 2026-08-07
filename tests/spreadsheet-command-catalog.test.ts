import { expect, test } from '@rstest/core';
import {
  spreadsheetCommandCatalog,
  spreadsheetRibbonTabs,
} from '../src/internal/features/work/editors/spreadsheet-command-catalog';

test('keeps the WPS spreadsheet information architecture stable', () => {
  expect(spreadsheetRibbonTabs.map((tab) => tab.id)).toEqual([
    'home',
    'insert',
    'pageLayout',
    'formulas',
    'data',
    'review',
    'view',
  ]);

  expect(spreadsheetCommandCatalog.undo.location).toEqual({
    area: 'quickAccess',
  });
  expect(spreadsheetCommandCatalog.redo.location).toEqual({
    area: 'quickAccess',
  });
  expect(spreadsheetCommandCatalog.paste.location).toEqual({
    area: 'ribbon',
    group: 'clipboard',
    tab: 'home',
  });
  expect(spreadsheetCommandCatalog.cut.shortcut).toEqual({
    aria: 'Control+X Meta+X',
    editor: ['Mod-x'],
    label: 'Cmd/Ctrl+X',
  });
  expect(spreadsheetCommandCatalog.copy.shortcut).toEqual({
    aria: 'Control+C Meta+C',
    editor: ['Mod-c'],
    label: 'Cmd/Ctrl+C',
  });
  expect(spreadsheetCommandCatalog.paste.shortcut).toEqual({
    aria: 'Control+V Meta+V',
    editor: ['Mod-v'],
    label: 'Cmd/Ctrl+V',
  });
  expect(spreadsheetCommandCatalog.formatPainter.location).toEqual({
    area: 'ribbon',
    group: 'clipboard',
    tab: 'home',
  });
  expect(spreadsheetCommandCatalog.conditionalFormatting.location).toEqual({
    area: 'ribbon',
    group: 'styles',
    tab: 'home',
  });
  expect(spreadsheetCommandCatalog.insertChart.location).toEqual({
    area: 'ribbon',
    group: 'charts',
    tab: 'insert',
  });
  expect(spreadsheetCommandCatalog.sortAscending.location).toEqual({
    area: 'ribbon',
    group: 'sortAndFilter',
    tab: 'data',
  });
  expect(spreadsheetCommandCatalog.recalculateWorkbook.shortcut).toEqual({
    aria: 'F9',
    editor: ['F9'],
    label: 'F9',
  });
});
