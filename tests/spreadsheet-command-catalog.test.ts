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
  expect(spreadsheetCommandCatalog.bold.shortcut).toEqual({
    aria: 'Control+B Meta+B Control+2',
    editor: ['Mod-b', 'Control-2'],
    label: 'Cmd/Ctrl+B 或 Ctrl+2',
  });
  expect(spreadsheetCommandCatalog.italic.shortcut).toEqual({
    aria: 'Control+I Meta+I Control+3',
    editor: ['Mod-i', 'Control-3'],
    label: 'Cmd/Ctrl+I 或 Ctrl+3',
  });
  expect(spreadsheetCommandCatalog.underline.shortcut).toEqual({
    aria: 'Control+U Meta+U Control+4',
    editor: ['Mod-u', 'Control-4'],
    label: 'Cmd/Ctrl+U 或 Ctrl+4',
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
  expect(spreadsheetCommandCatalog.growFont.shortcut).toEqual({
    aria: 'Control+Shift+. Meta+Shift+. Control+] Meta+]',
    editor: ['Mod-Shift-.', 'Mod-]'],
    label: 'Cmd/Ctrl+Shift+. 或 Cmd/Ctrl+]',
  });
  expect(spreadsheetCommandCatalog.shrinkFont.shortcut).toEqual({
    aria: 'Control+Shift+, Meta+Shift+, Control+[ Meta+[',
    editor: ['Mod-Shift-,', 'Mod-['],
    label: 'Cmd/Ctrl+Shift+, 或 Cmd/Ctrl+[',
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
    spreadsheetCommandCatalog.insertCurrentDate,
    spreadsheetCommandCatalog.insertCurrentTime,
  ]) {
    expect(command.location).toEqual({
      area: 'ribbon',
      group: 'number',
      tab: 'home',
    });
  }
  expect(spreadsheetCommandCatalog.borderOutside.shortcut).toEqual({
    aria: 'Control+Shift+& Meta+Shift+&',
    editor: ['Mod-Shift-&', 'Mod-Shift-7'],
    label: 'Cmd/Ctrl+Shift+&',
  });
  expect(spreadsheetCommandCatalog.borderNone.shortcut).toEqual({
    aria: 'Control+Shift+_ Meta+Shift+_',
    editor: ['Mod-Shift-_', 'Mod-Shift-Minus'],
    label: 'Cmd/Ctrl+Shift+_',
  });
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
  expect(spreadsheetCommandCatalog.insertCurrentDate.shortcut).toEqual({
    aria: 'Control+;',
    editor: ['Control-;'],
    label: 'Ctrl+;',
  });
  expect(spreadsheetCommandCatalog.insertCurrentTime.shortcut).toEqual({
    aria: 'Control+Shift+;',
    editor: ['Control-Shift-;'],
    label: 'Ctrl+Shift+;',
  });
  expect(spreadsheetCommandCatalog.decreaseDecimalPlaces).toMatchObject({
    id: 'number.decreaseDecimalPlaces',
    label: '减少小数位',
  });
  expect(spreadsheetCommandCatalog.increaseDecimalPlaces).toMatchObject({
    id: 'number.increaseDecimalPlaces',
    label: '增加小数位',
  });
  expect(spreadsheetCommandCatalog.formatCells.location).toEqual({
    area: 'ribbon',
    group: 'number',
    tab: 'home',
  });
  expect(spreadsheetCommandCatalog.formatCells.shortcut).toEqual({
    aria: 'Control+1 Meta+1',
    editor: ['Mod-1'],
    label: 'Cmd/Ctrl+1',
  });
  expect(spreadsheetCommandCatalog.formatCellsFont).toMatchObject({
    id: 'font.formatCellsFont',
    label: '设置字体格式',
    location: { area: 'ribbon', group: 'font', tab: 'home' },
    shortcut: {
      aria: 'Control+Shift+F Meta+Shift+F',
      editor: ['Mod-Shift-f'],
      label: 'Cmd/Ctrl+Shift+F',
    },
  });
  expect(spreadsheetCommandCatalog.formatCellsFontSize).toMatchObject({
    id: 'font.formatCellsFontSize',
    label: '设置字号格式',
    location: { area: 'ribbon', group: 'font', tab: 'home' },
    shortcut: {
      aria: 'Control+Shift+P Meta+Shift+P',
      editor: ['Mod-Shift-p'],
      label: 'Cmd/Ctrl+Shift+P',
    },
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
    spreadsheetCommandCatalog.borderDiagonalDown,
    spreadsheetCommandCatalog.borderDiagonalUp,
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
    spreadsheetCommandCatalog.textOrientationHorizontal,
    spreadsheetCommandCatalog.textOrientationAngleCounterclockwise,
    spreadsheetCommandCatalog.textOrientationAngleClockwise,
    spreadsheetCommandCatalog.textOrientationVertical,
    spreadsheetCommandCatalog.textOrientationRotateUp,
    spreadsheetCommandCatalog.textOrientationRotateDown,
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
  expect(spreadsheetCommandCatalog.hideRows.shortcut).toEqual({
    aria: 'Control+9 Meta+9',
    editor: ['Mod-9'],
    label: 'Cmd/Ctrl+9',
  });
  expect(spreadsheetCommandCatalog.hideColumns.shortcut).toEqual({
    aria: 'Control+0 Meta+0',
    editor: ['Mod-0'],
    label: 'Cmd/Ctrl+0',
  });
  expect(spreadsheetCommandCatalog.unhideRows.shortcut).toEqual({
    aria: 'Control+Shift+9 Meta+Shift+9',
    editor: ['Mod-Shift-9'],
    label: 'Cmd/Ctrl+Shift+9',
  });
  expect(spreadsheetCommandCatalog.unhideColumns.shortcut).toEqual({
    aria: 'Control+Shift+0 Meta+Shift+0',
    editor: ['Mod-Shift-0'],
    label: 'Cmd/Ctrl+Shift+0',
  });
  for (const command of [
    spreadsheetCommandCatalog.hideRows,
    spreadsheetCommandCatalog.hideColumns,
    spreadsheetCommandCatalog.unhideRows,
    spreadsheetCommandCatalog.unhideColumns,
  ]) {
    expect(command.location).toEqual({
      area: 'ribbon',
      group: 'cells',
      tab: 'home',
    });
  }
  for (const command of [
    spreadsheetCommandCatalog.autoSum,
    spreadsheetCommandCatalog.autoAverage,
    spreadsheetCommandCatalog.autoCount,
    spreadsheetCommandCatalog.autoMaximum,
    spreadsheetCommandCatalog.autoMinimum,
    spreadsheetCommandCatalog.fillDown,
    spreadsheetCommandCatalog.fillRight,
    spreadsheetCommandCatalog.fillUp,
    spreadsheetCommandCatalog.fillLeft,
    spreadsheetCommandCatalog.copyFormulaFromAbove,
    spreadsheetCommandCatalog.copyValueFromAbove,
    spreadsheetCommandCatalog.clearAll,
    spreadsheetCommandCatalog.clearFormats,
    spreadsheetCommandCatalog.clearContents,
    spreadsheetCommandCatalog.clearComments,
    spreadsheetCommandCatalog.clearHyperlinks,
    spreadsheetCommandCatalog.find,
    spreadsheetCommandCatalog.findAndSelect,
    spreadsheetCommandCatalog.goTo,
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
  expect(spreadsheetCommandCatalog.autoSum.shortcut).toEqual({
    aria: 'Alt+=',
    editor: ['Alt-equal'],
    label: 'Alt+=',
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
  expect(spreadsheetCommandCatalog.copyFormulaFromAbove.shortcut).toEqual({
    aria: "Control+'",
    editor: ["Control-'"],
    label: "Ctrl+'",
  });
  expect(spreadsheetCommandCatalog.copyValueFromAbove.shortcut).toEqual({
    aria: "Control+Shift+'",
    editor: ["Control-Shift-'"],
    label: "Ctrl+Shift+'",
  });
  expect(spreadsheetCommandCatalog.find.shortcut).toEqual({
    aria: 'Control+F Meta+F',
    editor: ['Mod-f'],
    label: 'Cmd/Ctrl+F',
  });
  expect(spreadsheetCommandCatalog.goTo.shortcut).toEqual({
    aria: 'Control+G F5',
    editor: ['Control-g', 'F5'],
    label: 'Ctrl+G 或 F5',
  });
  expect(spreadsheetCommandCatalog.insertChart.location).toEqual({
    area: 'ribbon',
    group: 'charts',
    tab: 'insert',
  });
  expect(spreadsheetCommandCatalog.hyperlink.location).toEqual({
    area: 'ribbon',
    group: 'links',
    tab: 'insert',
  });
  expect(spreadsheetCommandCatalog.hyperlink.shortcut).toEqual({
    aria: 'Control+K Meta+K',
    editor: ['Mod-k'],
    label: 'Cmd/Ctrl+K',
  });
  expect(spreadsheetCommandCatalog.sortAscending.location).toEqual({
    area: 'ribbon',
    group: 'sortAndFilter',
    tab: 'data',
  });
  expect(spreadsheetCommandCatalog.customSort).toMatchObject({
    id: 'data.customSort',
    label: '自定义排序',
    location: {
      area: 'ribbon',
      group: 'sortAndFilter',
      tab: 'data',
    },
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
  expect(spreadsheetCommandCatalog.dataValidation).toMatchObject({
    id: 'data.validation',
    label: '数据验证',
    location: {
      area: 'ribbon',
      group: 'dataTools',
      tab: 'data',
    },
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
