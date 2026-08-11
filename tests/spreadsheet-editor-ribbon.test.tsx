import { expect, test } from '@rstest/core';
import {
  fireEvent,
  render,
  screen,
  waitFor,
  within,
} from '@testing-library/react';
import type {
  SpreadsheetEditorCanCommands,
  SpreadsheetEditorCommands,
} from '../src/internal/features/work/editors/spreadsheet-command-controller';
import { SpreadsheetEditorRibbon } from '../src/internal/features/work/editors/spreadsheet-editor-ribbon';

test('uses the shared quick access and collapsible adaptive ribbon', () => {
  const { container } = render(
    <SpreadsheetEditorRibbon
      activeTab="home"
      can={spreadsheetCan()}
      commands={spreadsheetCommands(() => true)}
      content={{ type: 'spreadsheet', sheets: [] }}
      gridLinesVisible
      findOpen={false}
      panel={null}
      toolbarCell={null}
      onTabChange={() => undefined}
      onOpenFind={() => undefined}
      onTogglePanel={() => undefined}
    />,
  );

  expect(
    screen.getByRole('toolbar', { name: '快速访问工具栏' }),
  ).toBeInTheDocument();
  expect(screen.getByRole('button', { name: '撤销' })).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Z Meta+Z',
  );
  expect(screen.getByRole('button', { name: '重做' })).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+Z Meta+Shift+Z Control+Y Meta+Y',
  );
  expect(screen.queryByText('撤销与恢复')).not.toBeInTheDocument();

  const ribbon = container.querySelector('.work-spreadsheet-ribbon');
  expect(ribbon).not.toHaveAttribute('data-collapsed');
  fireEvent.doubleClick(screen.getByRole('tab', { name: '开始' }));
  expect(ribbon).toHaveAttribute('data-collapsed', 'true');
  expect(screen.getByRole('button', { name: '展开功能区' })).toHaveAttribute(
    'aria-expanded',
    'false',
  );
});

test('routes number-format controls through typed spreadsheet commands', () => {
  const formats: Array<{ attribute: string; value: unknown }> = [];
  const commands = spreadsheetCommands((attribute, value) => {
    formats.push({ attribute, value });
    return true;
  });

  render(
    <SpreadsheetEditorRibbon
      activeTab="home"
      can={spreadsheetCan()}
      commands={commands}
      content={{ type: 'spreadsheet', sheets: [] }}
      gridLinesVisible
      findOpen={false}
      panel={null}
      toolbarCell={{ v: 0.5, m: '50.0%', ct: { fa: '0.0%', t: 'n' } }}
      onTabChange={() => undefined}
      onOpenFind={() => undefined}
      onTogglePanel={() => undefined}
    />,
  );

  const formatSelect = screen.getByRole('combobox', { name: '数字格式' });
  expect(formatSelect).toHaveTextContent('百分比');
  fireEvent.click(formatSelect);
  fireEvent.click(screen.getByRole('option', { name: '数字' }));
  fireEvent.click(screen.getByRole('button', { name: '百分比格式' }));
  fireEvent.click(screen.getByRole('button', { name: '减少小数位' }));
  fireEvent.click(screen.getByRole('button', { name: '增加小数位' }));

  expect(formats).toEqual([
    { attribute: 'ct', value: { fa: '#,##0.00', t: 'n' } },
    { attribute: 'ct', value: { fa: '0.00%', t: 'n' } },
    { attribute: 'ct', value: { fa: '0%', t: 'n' } },
    { attribute: 'ct', value: { fa: '0.00%', t: 'n' } },
  ]);
});

test('routes the WPS Home clipboard group through typed commands', () => {
  const actions: string[] = [];
  render(
    <SpreadsheetEditorRibbon
      activeTab="home"
      can={spreadsheetCan()}
      commands={spreadsheetCommands(
        () => true,
        () => true,
        {
          copySelection: () => {
            actions.push('copy');
            return true;
          },
          cutSelection: () => {
            actions.push('cut');
            return true;
          },
          pasteSelection: () => {
            actions.push('paste');
            return true;
          },
          activateFormatPainter: (mode) => {
            actions.push(`format-painter:${mode}`);
            return true;
          },
          cancelFormatPainter: () => {
            actions.push('format-painter:cancel');
            return true;
          },
        },
      )}
      content={{ type: 'spreadsheet', sheets: [] }}
      gridLinesVisible
      findOpen={false}
      panel={null}
      toolbarCell={null}
      onTabChange={() => undefined}
      onOpenFind={() => undefined}
      onTogglePanel={() => undefined}
    />,
  );

  const clipboard = screen.getByRole('region', { name: '剪贴板' });
  const paste = within(clipboard).getByRole('button', { name: '粘贴' });
  const cut = within(clipboard).getByRole('button', { name: '剪切' });
  const copy = within(clipboard).getByRole('button', { name: '复制' });
  const formatPainter = within(clipboard).getByRole('button', {
    name: '格式刷',
  });
  expect(paste).toHaveAttribute('aria-keyshortcuts', 'Control+V Meta+V');
  expect(cut).toHaveAttribute('aria-keyshortcuts', 'Control+X Meta+X');
  expect(copy).toHaveAttribute('aria-keyshortcuts', 'Control+C Meta+C');
  expect(formatPainter).toHaveAttribute('aria-pressed', 'false');
  expect(formatPainter).toHaveAttribute(
    'title',
    '格式刷（单击应用一次，双击锁定连续应用）',
  );

  fireEvent.click(paste);
  fireEvent.click(cut);
  fireEvent.click(copy);
  fireEvent.click(formatPainter);
  fireEvent.doubleClick(formatPainter);
  expect(actions).toEqual([
    'paste',
    'cut',
    'copy',
    'format-painter:once',
    'format-painter:locked',
  ]);
});

test('exposes locked format-painter state and exits on another click', () => {
  const actions: string[] = [];
  render(
    <SpreadsheetEditorRibbon
      activeTab="home"
      can={spreadsheetCan()}
      commands={spreadsheetCommands(
        () => true,
        () => true,
        {
          cancelFormatPainter: () => {
            actions.push('cancel');
            return true;
          },
        },
      )}
      content={{ type: 'spreadsheet', sheets: [] }}
      formatPainterMode="locked"
      gridLinesVisible
      findOpen={false}
      panel={null}
      toolbarCell={null}
      onTabChange={() => undefined}
      onOpenFind={() => undefined}
      onTogglePanel={() => undefined}
    />,
  );

  const formatPainter = screen.getByRole('button', { name: '格式刷' });
  expect(formatPainter).toHaveAttribute('aria-pressed', 'true');
  expect(formatPainter).toHaveAttribute(
    'title',
    '格式刷已锁定（再次点击或按 Escape 退出）',
  );
  expect(formatPainter).toHaveTextContent('连续');
  fireEvent.click(formatPainter);
  expect(actions).toEqual(['cancel']);
});

test('operates WPS row and column actions from the Home cells group', async () => {
  const actions: string[] = [];
  const commands = spreadsheetCommands(
    () => true,
    () => true,
    {
      deleteSelectedStructure: (axis) => {
        actions.push(`delete:${axis}`);
        return true;
      },
      insertSelectedStructure: (axis, position) => {
        actions.push(`insert:${axis}:${position}`);
        return true;
      },
    },
  );

  render(
    <SpreadsheetEditorRibbon
      activeTab="home"
      can={spreadsheetCan()}
      commands={commands}
      content={{ type: 'spreadsheet', sheets: [] }}
      gridLinesVisible
      findOpen={false}
      panel={null}
      toolbarCell={null}
      onTabChange={() => undefined}
      onOpenFind={() => undefined}
      onTogglePanel={() => undefined}
    />,
  );

  const trigger = screen.getByRole('button', { name: '行和列' });
  expect(trigger).toHaveAttribute('aria-haspopup', 'menu');
  fireEvent.click(trigger);

  const menu = screen.getByRole('menu', { name: '行和列选项' });
  const insertAbove = within(menu).getByRole('menuitem', {
    name: '在上方插入行',
  });
  const deleteColumns = within(menu).getByRole('menuitem', {
    name: '删除所选列',
  });
  expect(
    within(menu).getByRole('menuitem', { name: '在下方插入行' }),
  ).toBeInTheDocument();
  expect(
    within(menu).getByRole('menuitem', { name: '在左侧插入列' }),
  ).toBeInTheDocument();
  expect(
    within(menu).getByRole('menuitem', { name: '在右侧插入列' }),
  ).toBeInTheDocument();
  expect(
    within(menu).getByRole('menuitem', { name: '删除所选行' }),
  ).toBeInTheDocument();

  await waitFor(() => expect(insertAbove).toHaveFocus());
  fireEvent.keyDown(menu, { key: 'End' });
  expect(deleteColumns).toHaveFocus();
  fireEvent.click(deleteColumns);
  expect(actions).toEqual(['delete:column']);
  expect(screen.queryByRole('menu', { name: '行和列选项' })).toBeNull();
  expect(trigger).toHaveFocus();

  fireEvent.click(trigger);
  fireEvent.click(
    within(screen.getByRole('menu', { name: '行和列选项' })).getByRole(
      'menuitem',
      { name: '在上方插入行' },
    ),
  );
  expect(actions).toEqual(['delete:column', 'insert:row:before']);
});

test('disables unavailable WPS row and column actions independently', () => {
  const can = spreadsheetCan();
  can.insertSelectedStructure = (axis, position) =>
    axis === 'row' && position === 'before';
  can.deleteSelectedStructure = (axis) => axis === 'row';

  render(
    <SpreadsheetEditorRibbon
      activeTab="home"
      can={can}
      commands={spreadsheetCommands(() => true)}
      content={{ type: 'spreadsheet', sheets: [] }}
      gridLinesVisible
      findOpen={false}
      panel={null}
      toolbarCell={null}
      onTabChange={() => undefined}
      onOpenFind={() => undefined}
      onTogglePanel={() => undefined}
    />,
  );

  fireEvent.click(screen.getByRole('button', { name: '行和列' }));
  const menu = screen.getByRole('menu', { name: '行和列选项' });
  expect(
    within(menu).getByRole('menuitem', { name: '在上方插入行' }),
  ).toBeEnabled();
  expect(
    within(menu).getByRole('menuitem', { name: '在下方插入行' }),
  ).toBeDisabled();
  expect(
    within(menu).getByRole('menuitem', { name: '在左侧插入列' }),
  ).toBeDisabled();
  expect(
    within(menu).getByRole('menuitem', { name: '删除所选行' }),
  ).toBeEnabled();
  expect(
    within(menu).getByRole('menuitem', { name: '删除所选列' }),
  ).toBeDisabled();
});

test('routes font, vertical alignment, and wrapping through cell formats', () => {
  const formats: Array<{ attribute: string; value: unknown }> = [];
  render(
    <SpreadsheetEditorRibbon
      activeTab="home"
      can={spreadsheetCan()}
      commands={spreadsheetCommands((attribute, value) => {
        formats.push({ attribute, value });
        return true;
      })}
      content={{ type: 'spreadsheet', sheets: [] }}
      gridLinesVisible
      findOpen={false}
      panel={null}
      toolbarCell={{ ff: 'Arial', vt: 0, tb: '2' }}
      onTabChange={() => undefined}
      onOpenFind={() => undefined}
      onTogglePanel={() => undefined}
    />,
  );

  const font = screen.getByRole('combobox', { name: '字体' });
  expect(font).toHaveTextContent('Arial');
  fireEvent.click(font);
  const simSun = screen.getByRole('option', { name: '宋体' });
  expect(simSun.querySelector('span')).toHaveAttribute(
    'style',
    'font-family: SimSun, "Songti SC", serif;',
  );
  fireEvent.click(simSun);
  fireEvent.click(screen.getByRole('button', { name: '底端对齐' }));
  fireEvent.click(screen.getByRole('button', { name: '自动换行' }));

  expect(formats).toEqual([
    { attribute: 'ff', value: 'SimSun' },
    { attribute: 'vt', value: 2 },
    { attribute: 'tb', value: '1' },
  ]);
  expect(screen.getByRole('button', { name: '加粗' })).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+B Meta+B',
  );
  expect(screen.getByRole('button', { name: '斜体' })).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+I Meta+I',
  );
  expect(screen.getByRole('button', { name: '下划线' })).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+U Meta+U',
  );
});

test('omits empty resource counts from spreadsheet ribbon actions', () => {
  render(
    <SpreadsheetEditorRibbon
      activeTab="insert"
      can={spreadsheetCan()}
      commands={spreadsheetCommands(() => true)}
      content={{ type: 'spreadsheet', sheets: [] }}
      gridLinesVisible
      findOpen={false}
      panel={null}
      toolbarCell={null}
      onTabChange={() => undefined}
      onOpenFind={() => undefined}
      onTogglePanel={() => undefined}
    />,
  );

  expect(screen.getByRole('button', { name: '插入图表' })).toBeInTheDocument();
  expect(
    screen.queryByRole('button', { name: '条件格式' }),
  ).not.toBeInTheDocument();
  expect(screen.queryByText('0')).not.toBeInTheDocument();
});

test('keeps conditional formatting in Home and sorting in Data', () => {
  const panels: string[] = [];
  const sorts: string[] = [];
  const { rerender } = render(
    <SpreadsheetEditorRibbon
      activeTab="home"
      can={spreadsheetCan()}
      commands={spreadsheetCommands(
        () => true,
        (direction) => {
          sorts.push(direction);
          return true;
        },
      )}
      content={{ type: 'spreadsheet', sheets: [] }}
      gridLinesVisible
      findOpen={false}
      panel={null}
      toolbarCell={null}
      onTabChange={() => undefined}
      onOpenFind={() => undefined}
      onTogglePanel={(panel) => panels.push(panel)}
    />,
  );

  fireEvent.click(screen.getByRole('button', { name: '条件格式' }));
  expect(panels).toEqual(['conditional-formatting']);

  rerender(
    <SpreadsheetEditorRibbon
      activeTab="data"
      can={spreadsheetCan()}
      commands={spreadsheetCommands(
        () => true,
        (direction) => {
          sorts.push(direction);
          return true;
        },
      )}
      content={{ type: 'spreadsheet', sheets: [] }}
      gridLinesVisible
      findOpen={false}
      panel={null}
      toolbarCell={null}
      onTabChange={() => undefined}
      onOpenFind={() => undefined}
      onTogglePanel={() => undefined}
    />,
  );

  fireEvent.click(screen.getByRole('button', { name: '升序' }));
  fireEvent.click(screen.getByRole('button', { name: '降序' }));
  expect(sorts).toEqual(['ascending', 'descending']);
});

test('exposes the spreadsheet find shortcut through the home ribbon', () => {
  let openCount = 0;
  render(
    <SpreadsheetEditorRibbon
      activeTab="home"
      can={spreadsheetCan()}
      commands={spreadsheetCommands(() => true)}
      content={{ type: 'spreadsheet', sheets: [] }}
      findOpen={false}
      gridLinesVisible
      panel={null}
      toolbarCell={null}
      onOpenFind={() => {
        openCount += 1;
      }}
      onTabChange={() => undefined}
      onTogglePanel={() => undefined}
    />,
  );

  const find = screen.getByRole('button', { name: '查找' });
  expect(find).toHaveAttribute('aria-keyshortcuts', 'Control+F Meta+F');
  fireEvent.click(find);
  expect(openCount).toBe(1);
});

test('operates the WPS Clear menu from the Home editing group', () => {
  const actions: string[] = [];
  render(
    <SpreadsheetEditorRibbon
      activeTab="home"
      can={spreadsheetCan()}
      commands={spreadsheetCommands(
        () => true,
        () => true,
        {
          clearSelectedCells: (mode = 'contents') => {
            actions.push(mode);
            return true;
          },
        },
      )}
      content={{ type: 'spreadsheet', sheets: [] }}
      findOpen={false}
      gridLinesVisible
      panel={null}
      toolbarCell={null}
      onOpenFind={() => undefined}
      onTabChange={() => undefined}
      onTogglePanel={() => undefined}
    />,
  );

  const editing = screen.getByRole('region', { name: '编辑' });
  const trigger = within(editing).getByRole('button', { name: '清除' });
  expect(trigger).toHaveAttribute('aria-haspopup', 'menu');
  fireEvent.click(trigger);

  const menu = screen.getByRole('menu', { name: '清除选项' });
  expect(
    within(menu).getByRole('menuitem', { name: '清除全部' }),
  ).toBeInTheDocument();
  expect(
    within(menu).getByRole('menuitem', { name: '清除格式' }),
  ).toBeInTheDocument();
  const contents = within(menu).getByRole('menuitem', { name: '清除内容' });
  expect(contents).toHaveAttribute('aria-keyshortcuts', 'Delete Backspace');
  expect(
    within(menu).getByRole('menuitem', { name: '清除批注' }),
  ).toBeInTheDocument();
  const hyperlinks = within(menu).getByRole('menuitem', {
    name: '清除超链接',
  });
  fireEvent.keyDown(menu, { key: 'End' });
  expect(hyperlinks).toHaveFocus();
  fireEvent.click(hyperlinks);
  expect(actions).toEqual(['hyperlinks']);
  expect(screen.queryByRole('menu', { name: '清除选项' })).toBeNull();
});

test('places the WPS merge split control in the Home alignment group', () => {
  const actions: string[] = [];
  render(
    <SpreadsheetEditorRibbon
      activeTab="home"
      can={spreadsheetCan()}
      commands={spreadsheetCommands(
        () => true,
        () => true,
        {
          mergeSelectedCells: (command) => {
            actions.push(command);
            return true;
          },
        },
      )}
      content={{ type: 'spreadsheet', sheets: [] }}
      gridLinesVisible
      findOpen={false}
      panel={null}
      toolbarCell={null}
      onTabChange={() => undefined}
      onOpenFind={() => undefined}
      onTogglePanel={() => undefined}
    />,
  );

  const alignment = screen.getByRole('region', { name: '对齐' });
  const cells = screen.getByRole('region', { name: '单元格' });
  const mergeAndCenter = within(alignment).getByRole('button', {
    name: '合并居中',
  });
  expect(mergeAndCenter).toHaveAttribute('aria-keyshortcuts', 'Control+M');
  expect(
    within(cells).queryByRole('button', { name: '合并单元格' }),
  ).toBeNull();

  fireEvent.click(mergeAndCenter);
  expect(actions).toEqual(['merge-and-center']);

  fireEvent.click(
    within(alignment).getByRole('button', { name: '更多合并方式' }),
  );
  const menu = screen.getByRole('menu', { name: '合并选项' });
  expect(
    within(menu).getByRole('menuitem', { name: '合并居中' }),
  ).toBeInTheDocument();
  expect(
    within(menu).getByRole('menuitem', { name: '合并单元格' }),
  ).toBeInTheDocument();
  expect(
    within(menu).getByRole('menuitem', { name: '跨行合并' }),
  ).toBeInTheDocument();
  expect(
    within(menu).getByRole('menuitem', { name: '取消合并单元格' }),
  ).toBeInTheDocument();
  const unmergeAndFill = within(menu).getByRole('menuitem', {
    name: '取消合并并填充',
  });
  fireEvent.keyDown(menu, { key: 'End' });
  expect(unmergeAndFill).toHaveFocus();
  fireEvent.click(unmergeAndFill);

  expect(actions).toEqual(['merge-and-center', 'unmerge-and-fill']);
  expect(screen.queryByRole('menu', { name: '合并选项' })).toBeNull();
});

test('exposes WPS AutoFilter state and routes the Data ribbon action', () => {
  const actions: string[] = [];
  render(
    <SpreadsheetEditorRibbon
      activeTab="data"
      autoFilterActive
      can={spreadsheetCan()}
      commands={spreadsheetCommands(
        () => true,
        () => true,
        {
          toggleAutoFilter: () => {
            actions.push('toggle-filter');
            return true;
          },
        },
      )}
      content={{ type: 'spreadsheet', sheets: [] }}
      findOpen={false}
      gridLinesVisible
      panel={null}
      toolbarCell={null}
      onOpenFind={() => undefined}
      onTabChange={() => undefined}
      onTogglePanel={() => undefined}
    />,
  );

  const filter = screen.getByRole('button', { name: '自动筛选' });
  expect(filter).toHaveAttribute('aria-pressed', 'true');
  expect(filter).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+L Meta+Shift+L',
  );
  expect(filter).toHaveAttribute(
    'title',
    '自动筛选（Cmd/Ctrl+Shift+L）；表头菜单（Alt+↓）',
  );
  fireEvent.click(filter);
  expect(actions).toEqual(['toggle-filter']);
});

test('operates WPS Freeze Panes from the View window group', () => {
  const actions: string[] = [];
  render(
    <SpreadsheetEditorRibbon
      activeTab="view"
      freezePanesActive
      freezePanesSelection={{
        row: [2, 2],
        column: [1, 1],
        row_focus: 2,
        column_focus: 1,
      }}
      can={spreadsheetCan()}
      commands={spreadsheetCommands(
        () => true,
        () => true,
        {
          setFreezePanes: (preset) => {
            actions.push(preset);
            return true;
          },
        },
      )}
      content={{ type: 'spreadsheet', sheets: [] }}
      findOpen={false}
      gridLinesVisible
      panel={null}
      toolbarCell={null}
      onOpenFind={() => undefined}
      onTabChange={() => undefined}
      onTogglePanel={() => undefined}
    />,
  );

  const trigger = screen.getByRole('button', { name: '冻结窗格' });
  expect(trigger).toHaveAttribute('aria-haspopup', 'menu');
  expect(trigger).toHaveAttribute('aria-pressed', 'true');
  fireEvent.click(trigger);

  const menu = screen.getByRole('menu', { name: '冻结窗格选项' });
  expect(
    within(menu).getByRole('menuitem', { name: '取消冻结窗格' }),
  ).toBeInTheDocument();
  expect(
    within(menu).getByRole('menuitem', {
      name: '冻结至第 2 行、A 列',
    }),
  ).toBeInTheDocument();
  expect(
    within(menu).getByRole('menuitem', { name: '冻结首行' }),
  ).toBeInTheDocument();
  const firstColumn = within(menu).getByRole('menuitem', {
    name: '冻结首列',
  });
  fireEvent.keyDown(menu, { key: 'End' });
  expect(firstColumn).toHaveFocus();
  fireEvent.click(firstColumn);
  expect(actions).toEqual(['firstColumn']);
  expect(screen.queryByRole('menu', { name: '冻结窗格选项' })).toBeNull();
});

function spreadsheetCan(): SpreadsheetEditorCanCommands {
  return {
    activateSheet: () => true,
    addSheet: () => true,
    clearSelectedCells: () => true,
    activateFormatPainter: () => true,
    applyFormatPainter: () => true,
    cancelFormatPainter: () => true,
    copySelection: () => true,
    cutSelection: () => true,
    deleteSelectedStructure: () => true,
    deleteSheet: () => true,
    duplicateSheet: () => true,
    hideSheet: () => true,
    insertSelectedStructure: () => true,
    mergeSelectedCells: () => true,
    moveSheet: () => true,
    moveSelection: () => true,
    openAutoFilterMenu: () => true,
    pasteCells: () => true,
    pasteSelection: () => true,
    recalculateFormula: () => true,
    renameSheet: () => true,
    redo: () => false,
    setCellFormat: () => true,
    setFreezePanes: () => true,
    setGridLines: () => true,
    setSheetColor: () => true,
    setSelectedStructureHidden: () => true,
    setSelectedStructureSize: () => true,
    setSpreadsheetContent: () => true,
    setZoom: () => true,
    selectCellRange: () => true,
    sortSelectedCells: () => true,
    toggleAutoFilter: () => true,
    undo: () => false,
  };
}

function spreadsheetCommands(
  setCellFormat: SpreadsheetEditorCommands['setCellFormat'],
  sortSelectedCells: SpreadsheetEditorCommands['sortSelectedCells'] = () =>
    true,
  overrides: Partial<
    Pick<
      SpreadsheetEditorCommands,
      | 'activateFormatPainter'
      | 'cancelFormatPainter'
      | 'clearSelectedCells'
      | 'copySelection'
      | 'cutSelection'
      | 'deleteSelectedStructure'
      | 'insertSelectedStructure'
      | 'mergeSelectedCells'
      | 'pasteSelection'
      | 'setFreezePanes'
      | 'toggleAutoFilter'
    >
  > = {},
): SpreadsheetEditorCommands {
  return {
    activateSheet: () => true,
    activateFormatPainter: overrides.activateFormatPainter ?? (() => true),
    addSheet: () => true,
    applyFormatPainter: () => true,
    cancelFormatPainter: overrides.cancelFormatPainter ?? (() => true),
    clearSelectedCells: overrides.clearSelectedCells ?? (() => true),
    copySelection: overrides.copySelection ?? (() => true),
    cutSelection: overrides.cutSelection ?? (() => true),
    deleteSelectedStructure: overrides.deleteSelectedStructure ?? (() => true),
    deleteSheet: () => true,
    duplicateSheet: () => true,
    hideSheet: () => true,
    insertSelectedStructure: overrides.insertSelectedStructure ?? (() => true),
    mergeSelectedCells: overrides.mergeSelectedCells ?? (() => true),
    moveSheet: () => true,
    moveSelection: () => true,
    openAutoFilterMenu: () => true,
    pasteCells: () => true,
    pasteSelection: overrides.pasteSelection ?? (() => true),
    recalculateFormula: () => true,
    renameSheet: () => true,
    redo: () => false,
    setCellFormat,
    setFreezePanes: overrides.setFreezePanes ?? (() => true),
    setGridLines: () => true,
    setSheetColor: () => true,
    setSelectedStructureHidden: () => true,
    setSelectedStructureSize: () => true,
    setSpreadsheetContent: () => true,
    setZoom: () => true,
    selectCellRange: () => true,
    sortSelectedCells,
    toggleAutoFilter: overrides.toggleAutoFilter ?? (() => true),
    undo: () => false,
  };
}
