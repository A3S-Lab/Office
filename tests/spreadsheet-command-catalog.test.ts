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
  expect(spreadsheetCommandCatalog.strike.location).toEqual({
    area: 'ribbon',
    group: 'font',
    tab: 'home',
  });
  expect(spreadsheetCommandCatalog.strike.shortcut).toEqual({
    aria: 'Control+5 Meta+5',
    editor: ['Mod-5'],
    label: 'Cmd/Ctrl+5',
  });
  for (const command of [
    spreadsheetCommandCatalog.numberFormatGeneral,
    spreadsheetCommandCatalog.numberFormatNumber,
    spreadsheetCommandCatalog.numberFormatCurrency,
    spreadsheetCommandCatalog.numberFormatAccounting,
    spreadsheetCommandCatalog.numberFormatPercent,
    spreadsheetCommandCatalog.numberFormatDate,
    spreadsheetCommandCatalog.numberFormatTime,
    spreadsheetCommandCatalog.numberFormatScientific,
    spreadsheetCommandCatalog.numberFormatFraction,
    spreadsheetCommandCatalog.numberFormatText,
    spreadsheetCommandCatalog.decreaseDecimalPlaces,
    spreadsheetCommandCatalog.increaseDecimalPlaces,
  ]) {
    expect(command.location).toEqual({
      area: 'ribbon',
      group: 'number',
      tab: 'home',
    });
  }
  expect(spreadsheetCommandCatalog.numberFormatGeneral.shortcut).toEqual({
    aria: 'Control+Shift+~ Meta+Shift+~',
    editor: ['Mod-Shift-~'],
    label: 'Cmd/Ctrl+Shift+~',
  });
  expect(spreadsheetCommandCatalog.numberFormatNumber.shortcut).toEqual({
    aria: 'Control+Shift+! Meta+Shift+!',
    editor: ['Mod-Shift-!'],
    label: 'Cmd/Ctrl+Shift+!',
  });
  expect(spreadsheetCommandCatalog.numberFormatCurrency.shortcut).toEqual({
    aria: 'Control+Shift+$ Meta+Shift+$',
    editor: ['Mod-Shift-$'],
    label: 'Cmd/Ctrl+Shift+$',
  });
  expect(spreadsheetCommandCatalog.numberFormatPercent.shortcut).toEqual({
    aria: 'Control+Shift+% Meta+Shift+%',
    editor: ['Mod-Shift-%'],
    label: 'Cmd/Ctrl+Shift+%',
  });
  expect(spreadsheetCommandCatalog.numberFormatDate.shortcut).toEqual({
    aria: 'Control+Shift+# Meta+Shift+#',
    editor: ['Mod-Shift-#'],
    label: 'Cmd/Ctrl+Shift+#',
  });
  expect(spreadsheetCommandCatalog.numberFormatTime.shortcut).toEqual({
    aria: 'Control+Shift+@ Meta+Shift+@',
    editor: ['Mod-Shift-@'],
    label: 'Cmd/Ctrl+Shift+@',
  });
  expect(spreadsheetCommandCatalog.numberFormatScientific.shortcut).toEqual({
    aria: 'Control+Shift+^ Meta+Shift+^',
    editor: ['Mod-Shift-^'],
    label: 'Cmd/Ctrl+Shift+^',
  });
  expect(spreadsheetCommandCatalog.decreaseDecimalPlaces).toMatchObject({
    id: 'number.decreaseDecimalPlaces',
    label: '减少小数位',
  });
  expect(spreadsheetCommandCatalog.increaseDecimalPlaces).toMatchObject({
    id: 'number.increaseDecimalPlaces',
    label: '增加小数位',
  });
  for (const command of [
    spreadsheetCommandCatalog.borderTop,
    spreadsheetCommandCatalog.borderBottom,
    spreadsheetCommandCatalog.borderLeft,
    spreadsheetCommandCatalog.borderRight,
    spreadsheetCommandCatalog.borderNone,
    spreadsheetCommandCatalog.borderAll,
    spreadsheetCommandCatalog.borderOutside,
    spreadsheetCommandCatalog.borderInside,
    spreadsheetCommandCatalog.borderHorizontal,
    spreadsheetCommandCatalog.borderVertical,
    spreadsheetCommandCatalog.borderDiagonal,
  ]) {
    expect(command.location).toEqual({
      area: 'ribbon',
      group: 'font',
      tab: 'home',
    });
  }
  expect(spreadsheetCommandCatalog.conditionalFormatting.location).toEqual({
    area: 'ribbon',
    group: 'styles',
    tab: 'home',
  });
  expect(spreadsheetCommandCatalog.cellStyles.location).toEqual({
    area: 'ribbon',
    group: 'styles',
    tab: 'home',
  });
  expect(spreadsheetCommandCatalog.mergeAndCenter.location).toEqual({
    area: 'ribbon',
    group: 'alignment',
    tab: 'home',
  });
  expect(spreadsheetCommandCatalog.mergeAndCenter.shortcut).toEqual({
    aria: 'Control+M',
    editor: ['Control-m'],
    label: 'Ctrl+M',
  });
  for (const command of [
    spreadsheetCommandCatalog.mergeCells,
    spreadsheetCommandCatalog.mergeAcross,
    spreadsheetCommandCatalog.unmergeCells,
    spreadsheetCommandCatalog.unmergeAndFill,
  ]) {
    expect(command.location).toEqual({
      area: 'ribbon',
      group: 'alignment',
      tab: 'home',
    });
  }
  expect(spreadsheetCommandCatalog.insertRowsAbove.location).toEqual({
    area: 'ribbon',
    group: 'cells',
    tab: 'home',
  });
  expect(spreadsheetCommandCatalog.insertRowsBelow.location).toEqual({
    area: 'ribbon',
    group: 'cells',
    tab: 'home',
  });
  expect(spreadsheetCommandCatalog.insertColumnsLeft.location).toEqual({
    area: 'ribbon',
    group: 'cells',
    tab: 'home',
  });
  expect(spreadsheetCommandCatalog.insertColumnsRight.location).toEqual({
    area: 'ribbon',
    group: 'cells',
    tab: 'home',
  });
  expect(spreadsheetCommandCatalog.deleteRows.location).toEqual({
    area: 'ribbon',
    group: 'cells',
    tab: 'home',
  });
  expect(spreadsheetCommandCatalog.deleteColumns.location).toEqual({
    area: 'ribbon',
    group: 'cells',
    tab: 'home',
  });
  for (const command of [
    spreadsheetCommandCatalog.fillDown,
    spreadsheetCommandCatalog.fillRight,
    spreadsheetCommandCatalog.fillUp,
    spreadsheetCommandCatalog.fillLeft,
    spreadsheetCommandCatalog.clearAll,
    spreadsheetCommandCatalog.clearFormats,
    spreadsheetCommandCatalog.clearContents,
    spreadsheetCommandCatalog.clearComments,
    spreadsheetCommandCatalog.clearHyperlinks,
  ]) {
    expect(command.location).toEqual({
      area: 'ribbon',
      group: 'editing',
      tab: 'home',
    });
  }
  expect(spreadsheetCommandCatalog.clearContents.shortcut).toEqual({
    aria: 'Delete Backspace',
    editor: ['Delete', 'Backspace'],
    label: 'Delete/Backspace',
  });
  expect(spreadsheetCommandCatalog.fillDown.shortcut).toEqual({
    aria: 'Control+D Meta+D',
    editor: ['Mod-d'],
    label: 'Cmd/Ctrl+D',
  });
  expect(spreadsheetCommandCatalog.fillRight.shortcut).toEqual({
    aria: 'Control+R Meta+R',
    editor: ['Mod-r'],
    label: 'Cmd/Ctrl+R',
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
  expect(spreadsheetCommandCatalog.autoFilter.location).toEqual({
    area: 'ribbon',
    group: 'sortAndFilter',
    tab: 'data',
  });
  expect(spreadsheetCommandCatalog.autoFilter.shortcut).toEqual({
    aria: 'Control+Shift+L Meta+Shift+L',
    editor: ['Mod-Shift-l'],
    label: 'Cmd/Ctrl+Shift+L',
  });
  expect(spreadsheetCommandCatalog.autoFilter.menuShortcut).toEqual({
    aria: 'Alt+ArrowDown',
    editor: ['Alt-ArrowDown'],
    label: 'Alt+↓',
  });
  expect(spreadsheetCommandCatalog.recalculateWorkbook.shortcut).toEqual({
    aria: 'F9',
    editor: ['F9'],
    label: 'F9',
  });
  expect(spreadsheetCommandCatalog.freezePanes.location).toEqual({
    area: 'ribbon',
    group: 'window',
    tab: 'view',
  });
});
