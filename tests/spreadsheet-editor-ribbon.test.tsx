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
  const decimalAdjustments: string[] = [];
  const formatCellsOpens: string[] = [];
  const commands = spreadsheetCommands(
    (attribute, value) => {
      formats.push({ attribute, value });
      return true;
    },
    undefined,
    {
      adjustDecimalPlaces: (direction) => {
        decimalAdjustments.push(direction);
        return true;
      },
      openFormatCells: () => {
        formatCellsOpens.push('open');
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
      toolbarCell={{ v: 0.5, m: '50.0%', ct: { fa: '0.0%', t: 'n' } }}
      onTabChange={() => undefined}
      onTogglePanel={() => undefined}
    />,
  );

  const formatSelect = screen.getByRole('combobox', { name: '数字格式' });
  expect(formatSelect).toHaveTextContent('百分比');
  fireEvent.click(formatSelect);
  expect(screen.getAllByRole('option')).toHaveLength(11);
  expect(screen.getByRole('option', { name: '货币' })).toHaveTextContent(
    'Cmd/Ctrl+Shift+$',
  );
  expect(screen.getByRole('option', { name: '短日期' })).toHaveTextContent(
    'Cmd/Ctrl+Shift+#',
  );
  fireEvent.click(screen.getByRole('option', { name: '数字' }));
  fireEvent.click(formatSelect);
  fireEvent.click(screen.getByRole('option', { name: '短日期' }));
  fireEvent.click(formatSelect);
  fireEvent.click(screen.getByRole('option', { name: '文本' }));
  const currency = screen.getByRole('button', { name: '货币格式' });
  expect(currency).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+$ Meta+Shift+$',
  );
  fireEvent.click(currency);
  const percent = screen.getByRole('button', { name: '百分比格式' });
  expect(percent).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+% Meta+Shift+%',
  );
  fireEvent.click(percent);
  fireEvent.click(screen.getByRole('button', { name: '减少小数位' }));
  fireEvent.click(screen.getByRole('button', { name: '增加小数位' }));
  const formatCells = screen.getByRole('button', {
    name: '设置单元格格式',
  });
  expect(formatCells).toHaveAttribute('aria-keyshortcuts', 'Control+1 Meta+1');
  expect(formatCells).toHaveAttribute('title', '设置单元格格式（Cmd/Ctrl+1）');
  fireEvent.click(formatCells);

  expect(formats).toEqual([
    { attribute: 'ct', value: { fa: '#,##0.00', t: 'n' } },
    { attribute: 'ct', value: { fa: 'yyyy-MM-dd', t: 'd' } },
    { attribute: 'ct', value: { fa: '@', t: 's' } },
    { attribute: 'ct', value: { fa: '[$¥-804]#,##0.00', t: 'n' } },
    { attribute: 'ct', value: { fa: '0.00%', t: 'n' } },
  ]);
  expect(decimalAdjustments).toEqual(['decrease', 'increase']);
  expect(formatCellsOpens).toEqual(['open']);
});

test('applies grouped WPS cell styles from a preview gallery', async () => {
  const styles: string[] = [];
  render(
    <SpreadsheetEditorRibbon
      activeTab="home"
      can={spreadsheetCan()}
      commands={spreadsheetCommands(
        () => true,
        () => true,
        {
          applyCellStyle: (preset) => {
            styles.push(preset);
            return true;
          },
        },
      )}
      content={{ type: 'spreadsheet', sheets: [] }}
      gridLinesVisible
      findOpen={false}
      panel={null}
      toolbarCell={{ bg: '#c6efce', fc: '#006100' }}
      onTabChange={() => undefined}
      onTogglePanel={() => undefined}
    />,
  );

  const trigger = screen.getByRole('button', { name: '单元格样式' });
  expect(trigger).toHaveAttribute('aria-haspopup', 'menu');
  expect(trigger).toHaveAttribute('title', '单元格样式（当前：好）');
  fireEvent.click(trigger);

  const menu = screen.getByRole('menu', { name: '单元格样式库' });
  expect(within(menu).getByRole('group', { name: '常用' })).toBeInTheDocument();
  expect(
    within(menu).getByRole('group', { name: '数据和模型' }),
  ).toBeInTheDocument();
  expect(
    within(menu).getByRole('group', { name: '标题和汇总' }),
  ).toBeInTheDocument();
  expect(within(menu).getAllByRole('menuitemradio')).toHaveLength(17);
  const good = within(menu).getByRole('menuitemradio', {
    name: '应用单元格样式：好',
  });
  expect(good).toHaveAttribute('aria-checked', 'true');
  expect(
    good.querySelector('.work-spreadsheet-cell-style-preview'),
  ).toHaveStyle({
    backgroundColor: '#c6efce',
    color: '#006100',
  });
  await waitFor(() =>
    expect(
      within(menu).getByRole('menuitemradio', {
        name: '应用单元格样式：常规',
      }),
    ).toHaveFocus(),
  );
  const total = within(menu).getByRole('menuitemradio', {
    name: '应用单元格样式：总计',
  });
  const neutral = within(menu).getByRole('menuitemradio', {
    name: '应用单元格样式：适中',
  });
  fireEvent.keyDown(menu, { key: 'ArrowRight' });
  expect(good).toHaveFocus();
  fireEvent.keyDown(menu, { key: 'ArrowDown' });
  expect(neutral).toHaveFocus();
  fireEvent.keyDown(menu, { key: 'End' });
  expect(total).toHaveFocus();

  fireEvent.click(
    within(menu).getByRole('menuitemradio', {
      name: '应用单元格样式：输入',
    }),
  );
  expect(styles).toEqual(['input']);
  expect(screen.queryByRole('menu', { name: '单元格样式库' })).toBeNull();
  expect(trigger).toHaveFocus();
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
          pasteSpecial: (content) => {
            actions.push(`paste-special:${content}`);
            return true;
          },
          openPasteSpecial: () => {
            actions.push('open-paste-special');
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
  const pasteDisclosure = within(clipboard).getByTitle('更多粘贴方式');
  fireEvent.click(pasteDisclosure);
  let pasteMenu = screen.getByRole('menu', { name: '粘贴选项' });
  fireEvent.click(within(pasteMenu).getByRole('menuitem', { name: '值' }));
  expect(pasteDisclosure).toHaveFocus();
  fireEvent.click(pasteDisclosure);
  pasteMenu = screen.getByRole('menu', { name: '粘贴选项' });
  fireEvent.click(
    within(pasteMenu).getByRole('menuitem', { name: /选择性粘贴/ }),
  );
  expect(actions).toEqual([
    'paste',
    'cut',
    'copy',
    'format-painter:once',
    'format-painter:locked',
    'paste-special:values',
    'open-paste-special',
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
      setSelectedStructureHidden: (axis, hidden) => {
        actions.push(`${hidden ? 'hide' : 'unhide'}:${axis}`);
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
  const hideRows = within(menu).getByRole('menuitem', {
    name: '隐藏所选行',
  });
  const unhideColumns = within(menu).getByRole('menuitem', {
    name: '取消隐藏所选列',
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
  expect(hideRows).toHaveAttribute('aria-keyshortcuts', 'Control+9 Meta+9');
  expect(hideRows).toHaveTextContent('Cmd/Ctrl+9');
  expect(unhideColumns).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+0 Meta+Shift+0',
  );

  await waitFor(() => expect(insertAbove).toHaveFocus());
  fireEvent.keyDown(menu, { key: 'End' });
  expect(unhideColumns).toHaveFocus();
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

  fireEvent.click(trigger);
  fireEvent.click(
    within(screen.getByRole('menu', { name: '行和列选项' })).getByRole(
      'menuitem',
      { name: '隐藏所选行' },
    ),
  );
  expect(actions).toEqual(['delete:column', 'insert:row:before', 'hide:row']);
});

test('disables unavailable WPS row and column actions independently', () => {
  const can = spreadsheetCan();
  can.insertSelectedStructure = (axis, position) =>
    axis === 'row' && position === 'before';
  can.deleteSelectedStructure = (axis) => axis === 'row';
  can.setSelectedStructureHidden = (axis, hidden) => axis === 'row' && hidden;

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
  expect(
    within(menu).getByRole('menuitem', { name: '隐藏所选行' }),
  ).toBeEnabled();
  expect(
    within(menu).getByRole('menuitem', { name: '取消隐藏所选行' }),
  ).toBeDisabled();
});

test('offers the six WPS text orientations as an accessible radio menu', async () => {
  const orientations: string[] = [];
  render(
    <SpreadsheetEditorRibbon
      activeTab="home"
      can={spreadsheetCan()}
      commands={spreadsheetCommands(
        () => true,
        () => true,
        {
          setTextOrientation: (orientation) => {
            orientations.push(orientation);
            return true;
          },
        },
      )}
      content={{ type: 'spreadsheet', sheets: [] }}
      gridLinesVisible
      findOpen={false}
      panel={null}
      toolbarCell={{ rt: 135 }}
      onTabChange={() => undefined}
      onTogglePanel={() => undefined}
    />,
  );

  const alignment = screen.getByRole('region', { name: '对齐' });
  const trigger = within(alignment).getByRole('button', { name: '文字方向' });
  expect(trigger).toHaveAttribute('aria-haspopup', 'menu');
  expect(trigger).toHaveAttribute('aria-pressed', 'true');
  expect(trigger).toHaveAttribute('title', '文字方向（当前：顺时针倾斜）');
  fireEvent.click(trigger);

  const menu = screen.getByRole('menu', { name: '文字方向选项' });
  const options = within(menu).getAllByRole('menuitemradio');
  expect(options).toHaveLength(6);
  expect(
    within(menu).getByRole('menuitemradio', { name: '顺时针倾斜' }),
  ).toHaveAttribute('aria-checked', 'true');
  const horizontal = within(menu).getByRole('menuitemradio', {
    name: '横排文字',
  });
  const rotateDown = within(menu).getByRole('menuitemradio', {
    name: '向下旋转文字',
  });
  await waitFor(() => expect(horizontal).toHaveFocus());
  fireEvent.keyDown(menu, { key: 'End' });
  expect(rotateDown).toHaveFocus();
  fireEvent.click(
    within(menu).getByRole('menuitemradio', { name: '竖排文字' }),
  );
  expect(orientations).toEqual(['vertical']);
  expect(screen.queryByRole('menu', { name: '文字方向选项' })).toBeNull();
  expect(trigger).toHaveFocus();
});

test('routes WPS font, underline styles, vertical alignment, and wrapping through cell formats', () => {
  const formats: Array<{ attribute: string; value: unknown }> = [];
  const fontSteps: string[] = [];
  render(
    <SpreadsheetEditorRibbon
      activeTab="home"
      can={spreadsheetCan()}
      commands={spreadsheetCommands(
        (attribute, value) => {
          formats.push({ attribute, value });
          return true;
        },
        undefined,
        {
          adjustFontSize: (direction) => {
            fontSteps.push(direction);
            return true;
          },
        },
      )}
      content={{ type: 'spreadsheet', sheets: [] }}
      gridLinesVisible
      findOpen={false}
      panel={null}
      toolbarCell={{ cl: 0, ff: 'Arial', un: 4, vt: 0, tb: '2' }}
      onTabChange={() => undefined}
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
  const growFont = screen.getByRole('button', { name: '增大字号' });
  const shrinkFont = screen.getByRole('button', { name: '减小字号' });
  fireEvent.click(growFont);
  fireEvent.click(shrinkFont);
  const underline = screen.getByRole('button', { name: '下划线' });
  expect(underline).toHaveAttribute('aria-pressed', 'true');
  fireEvent.click(underline);
  fireEvent.click(screen.getByRole('button', { name: '更多下划线' }));
  const underlineMenu = screen.getByRole('menu', { name: '下划线样式' });
  expect(
    within(underlineMenu).getByRole('menuitemradio', {
      name: '双会计用下划线',
    }),
  ).toHaveAttribute('aria-checked', 'true');
  fireEvent.click(
    within(underlineMenu).getByRole('menuitemradio', {
      name: '双会计用下划线',
    }),
  );
  fireEvent.click(screen.getByRole('button', { name: '删除线' }));
  fireEvent.click(screen.getByRole('button', { name: '底端对齐' }));
  fireEvent.click(screen.getByRole('button', { name: '自动换行' }));

  expect(formats).toEqual([
    { attribute: 'ff', value: 'SimSun' },
    { attribute: 'un', value: 0 },
    { attribute: 'un', value: 4 },
    { attribute: 'cl', value: 1 },
    { attribute: 'vt', value: 2 },
    { attribute: 'tb', value: '1' },
  ]);
  expect(fontSteps).toEqual(['grow', 'shrink']);
  expect(growFont).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+. Meta+Shift+. Control+] Meta+]',
  );
  expect(growFont).toHaveAttribute(
    'title',
    '增大字号（Cmd/Ctrl+Shift+. 或 Cmd/Ctrl+]）',
  );
  expect(shrinkFont).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+, Meta+Shift+, Control+[ Meta+[',
  );
  expect(screen.getByRole('button', { name: '加粗' })).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+B Meta+B Control+2',
  );
  expect(screen.getByRole('button', { name: '斜体' })).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+I Meta+I Control+3',
  );
  expect(underline).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+U Meta+U Control+4',
  );
  expect(screen.getByRole('button', { name: '删除线' })).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+5 Meta+5',
  );
});

test('removes direct font and fill colors from the WPS color menus', () => {
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
      toolbarCell={{ bg: '#fff2cc', fc: '#1155cc' }}
      onTabChange={() => undefined}
      onTogglePanel={() => undefined}
    />,
  );

  const textColor = screen.getByRole('button', { name: '文字颜色' });
  fireEvent.click(textColor);
  fireEvent.click(screen.getByRole('button', { name: '自动颜色' }));
  expect(textColor).toHaveFocus();

  const fillColor = screen.getByRole('button', { name: '填充颜色' });
  fireEvent.click(fillColor);
  fireEvent.click(screen.getByRole('button', { name: '无填充' }));
  expect(fillColor).toHaveFocus();

  expect(formats).toEqual([
    { attribute: 'fc', value: undefined },
    { attribute: 'bg', value: undefined },
  ]);
});

test('omits empty resource counts from spreadsheet ribbon actions', () => {
  const actions: string[] = [];
  render(
    <SpreadsheetEditorRibbon
      activeTab="insert"
      can={spreadsheetCan()}
      commands={spreadsheetCommands(() => true, undefined, {
        openHyperlink: () => {
          actions.push('hyperlink');
          return true;
        },
      })}
      content={{ type: 'spreadsheet', sheets: [] }}
      gridLinesVisible
      findOpen={false}
      panel={null}
      toolbarCell={null}
      onTabChange={() => undefined}
      onTogglePanel={() => undefined}
    />,
  );

  expect(screen.getByRole('button', { name: '插入图表' })).toBeInTheDocument();
  const links = screen.getByRole('region', { name: '链接' });
  const hyperlink = within(links).getByRole('button', { name: '超链接' });
  expect(hyperlink).toHaveAttribute('aria-keyshortcuts', 'Control+K Meta+K');
  expect(hyperlink).toHaveAttribute('title', '超链接（Cmd/Ctrl+K）');
  fireEvent.click(hyperlink);
  expect(actions).toEqual(['hyperlink']);
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
      onTogglePanel={() => undefined}
    />,
  );

  fireEvent.click(screen.getByRole('button', { name: '升序' }));
  fireEvent.click(screen.getByRole('button', { name: '降序' }));
  expect(sorts).toEqual(['ascending', 'descending']);
});

test('places data validation in the WPS Data tools group', () => {
  const actions: string[] = [];
  render(
    <SpreadsheetEditorRibbon
      activeTab="data"
      can={spreadsheetCan()}
      commands={spreadsheetCommands(() => true, undefined, {
        openDataValidation: () => {
          actions.push('data-validation');
          return true;
        },
      })}
      content={{ type: 'spreadsheet', sheets: [] }}
      gridLinesVisible
      findOpen={false}
      panel={null}
      toolbarCell={null}
      onTabChange={() => undefined}
      onTogglePanel={() => undefined}
    />,
  );

  const dataTools = screen.getByRole('region', { name: '数据工具' });
  fireEvent.click(within(dataTools).getByRole('button', { name: '数据验证' }));
  expect(actions).toEqual(['data-validation']);
});

test('operates the WPS Find and Select menu from the Home ribbon', async () => {
  const actions: string[] = [];
  render(
    <SpreadsheetEditorRibbon
      activeTab="home"
      can={spreadsheetCan()}
      commands={spreadsheetCommands(() => true, undefined, {
        openFind: () => {
          actions.push('find');
          return true;
        },
        openGoTo: () => {
          actions.push('go-to');
          return true;
        },
      })}
      content={{ type: 'spreadsheet', sheets: [] }}
      findOpen={false}
      gridLinesVisible
      panel={null}
      toolbarCell={null}
      onTabChange={() => undefined}
      onTogglePanel={() => undefined}
    />,
  );

  const trigger = screen.getByRole('button', { name: '查找和选择' });
  expect(trigger).toHaveAttribute('aria-haspopup', 'menu');
  expect(trigger).toHaveAttribute('aria-pressed', 'false');
  fireEvent.click(trigger);

  const menu = screen.getByRole('menu', { name: '查找和选择选项' });
  const find = within(menu).getByRole('menuitem', { name: '查找' });
  const goTo = within(menu).getByRole('menuitem', { name: '定位' });
  expect(find).toHaveAttribute('aria-keyshortcuts', 'Control+F Meta+F');
  expect(find).toHaveTextContent('Cmd/Ctrl+F');
  expect(goTo).toHaveAttribute('aria-keyshortcuts', 'Control+G F5');
  expect(goTo).toHaveTextContent('Ctrl+G 或 F5');
  await waitFor(() => expect(find).toHaveFocus());
  fireEvent.keyDown(menu, { key: 'ArrowDown' });
  expect(goTo).toHaveFocus();
  fireEvent.keyDown(menu, { key: 'Home' });
  expect(find).toHaveFocus();
  fireEvent.keyDown(menu, { key: 'End' });
  expect(goTo).toHaveFocus();
  fireEvent.click(goTo);
  expect(actions).toEqual(['go-to']);
  expect(screen.queryByRole('menu', { name: '查找和选择选项' })).toBeNull();
  expect(trigger).toHaveFocus();

  fireEvent.click(trigger);
  fireEvent.click(
    within(screen.getByRole('menu', { name: '查找和选择选项' })).getByRole(
      'menuitem',
      { name: '查找' },
    ),
  );
  expect(actions).toEqual(['go-to', 'find']);
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

test('operates the WPS AutoSum split command and aggregate menu from Home', async () => {
  const actions: string[] = [];
  render(
    <SpreadsheetEditorRibbon
      activeTab="home"
      can={spreadsheetCan()}
      commands={spreadsheetCommands(
        () => true,
        () => true,
        {
          applyAutoSum: (functionName) => {
            actions.push(functionName);
            return true;
          },
        },
      )}
      content={{ type: 'spreadsheet', sheets: [] }}
      findOpen={false}
      gridLinesVisible
      panel={null}
      toolbarCell={null}
      onTabChange={() => undefined}
      onTogglePanel={() => undefined}
    />,
  );

  const editing = screen.getByRole('region', { name: '编辑' });
  const primary = within(editing).getByRole('button', { name: '自动求和' });
  const disclosure = within(editing).getByRole('button', {
    name: '更多自动计算方式',
  });
  expect(primary).toHaveAttribute('aria-keyshortcuts', 'Alt+=');
  expect(primary).toHaveAttribute('title', '自动求和（Alt+=）');
  fireEvent.click(primary);
  expect(actions).toEqual(['sum']);

  expect(disclosure).toHaveAttribute('aria-haspopup', 'menu');
  fireEvent.click(disclosure);
  const menu = screen.getByRole('menu', { name: '自动计算选项' });
  const sum = within(menu).getByRole('menuitem', { name: '自动求和' });
  const average = within(menu).getByRole('menuitem', { name: '平均值' });
  const count = within(menu).getByRole('menuitem', { name: '计数' });
  const maximum = within(menu).getByRole('menuitem', { name: '最大值' });
  const minimum = within(menu).getByRole('menuitem', { name: '最小值' });
  expect(sum).toHaveAttribute('aria-keyshortcuts', 'Alt+=');
  await waitFor(() => expect(sum).toHaveFocus());
  fireEvent.keyDown(menu, { key: 'ArrowDown' });
  expect(average).toHaveFocus();
  fireEvent.keyDown(menu, { key: 'End' });
  expect(minimum).toHaveFocus();
  expect(count).toBeEnabled();
  expect(maximum).toBeEnabled();
  fireEvent.click(average);
  expect(actions).toEqual(['sum', 'average']);
  expect(screen.queryByRole('menu', { name: '自动计算选项' })).toBeNull();
  expect(disclosure).toHaveFocus();
});

test('operates the keyboard-accessible WPS Fill menu from Home', async () => {
  const actions: string[] = [];
  render(
    <SpreadsheetEditorRibbon
      activeTab="home"
      can={spreadsheetCan()}
      commands={spreadsheetCommands(
        () => true,
        () => true,
        {
          fillSelectedCells: (direction) => {
            actions.push(direction);
            return true;
          },
        },
      )}
      content={{ type: 'spreadsheet', sheets: [] }}
      findOpen={false}
      gridLinesVisible
      panel={null}
      toolbarCell={null}
      onTabChange={() => undefined}
      onTogglePanel={() => undefined}
    />,
  );

  const editing = screen.getByRole('region', { name: '编辑' });
  const trigger = within(editing).getByRole('button', { name: '填充' });
  expect(trigger).toHaveAttribute('aria-haspopup', 'menu');
  fireEvent.click(trigger);

  const menu = screen.getByRole('menu', { name: '填充选项' });
  const down = within(menu).getByRole('menuitem', { name: '向下填充' });
  const right = within(menu).getByRole('menuitem', { name: '向右填充' });
  const up = within(menu).getByRole('menuitem', { name: '向上填充' });
  const left = within(menu).getByRole('menuitem', { name: '向左填充' });
  expect(down).toHaveAttribute('aria-keyshortcuts', 'Control+D Meta+D');
  expect(right).toHaveAttribute('aria-keyshortcuts', 'Control+R Meta+R');
  expect(up).not.toHaveAttribute('aria-keyshortcuts');
  expect(left).not.toHaveAttribute('aria-keyshortcuts');
  await waitFor(() => expect(down).toHaveFocus());
  fireEvent.keyDown(menu, { key: 'End' });
  expect(left).toHaveFocus();
  fireEvent.keyDown(menu, { key: 'Home' });
  expect(down).toHaveFocus();
  fireEvent.keyDown(menu, { key: 'ArrowDown' });
  expect(right).toHaveFocus();
  fireEvent.click(right);
  expect(actions).toEqual(['right']);
  expect(screen.queryByRole('menu', { name: '填充选项' })).toBeNull();
  expect(trigger).toHaveFocus();
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

test('shows a contextual Table Design ribbon with all 60 built-in styles', () => {
  const actions: string[] = [];
  const table = {
    id: 'table-1',
    name: 'SalesTable',
    range: {
      row: [0, 2] as [number, number],
      column: [0, 2] as [number, number],
    },
    columns: [{ name: 'Region' }, { name: 'Revenue' }, { name: 'Status' }],
    filters: [],
    headerRow: true,
    totalsRow: false,
    style: { family: 'medium' as const, number: 2 },
    showFirstColumn: false,
    showLastColumn: false,
    showRowStripes: true,
    showColumnStripes: false,
  };
  render(
    <SpreadsheetEditorRibbon
      activeTab="tableDesign"
      activeTable={table}
      activeTableSheetId="sheet-1"
      can={spreadsheetCan()}
      commands={spreadsheetCommands(
        () => true,
        () => true,
        {
          convertTableToRange: () => {
            actions.push('convert');
            return true;
          },
          updateTable: (_sheetId, _tableId, patch) => {
            actions.push(JSON.stringify(patch));
            return true;
          },
        },
      )}
      content={{ type: 'spreadsheet', sheets: [] }}
      findOpen={false}
      gridLinesVisible
      panel={null}
      toolbarCell={null}
      onTabChange={() => undefined}
      onTogglePanel={() => undefined}
    />,
  );

  const contextualTab = screen.getByRole('tab', { name: '表格设计' });
  expect(contextualTab).toHaveAttribute('data-contextual', 'true');
  expect(contextualTab).toHaveAttribute('aria-selected', 'true');

  const name = screen.getByRole('textbox', { name: '表格名称' });
  expect(name).toHaveValue('SalesTable');
  fireEvent.change(name, { target: { value: 'Sales_2026' } });
  fireEvent.keyDown(name, { key: 'Enter' });

  const tableStyleTrigger = screen.getByRole('button', { name: '表格样式' });
  fireEvent.click(tableStyleTrigger);
  const gallery = screen.getByRole('menu', { name: '表格样式库' });
  expect(within(gallery).getAllByRole('menuitemradio')).toHaveLength(60);
  expect(
    within(gallery).getByRole('group', { name: '浅色' }),
  ).toBeInTheDocument();
  expect(
    within(gallery).getByRole('group', { name: '中等' }),
  ).toBeInTheDocument();
  expect(
    within(gallery).getByRole('group', { name: '深色' }),
  ).toBeInTheDocument();
  const light1 = within(gallery).getByRole('menuitemradio', {
    name: '应用表格样式：浅色 1',
  });
  const light2 = within(gallery).getByRole('menuitemradio', {
    name: '应用表格样式：浅色 2',
  });
  const light9 = within(gallery).getByRole('menuitemradio', {
    name: '应用表格样式：浅色 9',
  });
  light1.focus();
  fireEvent.keyDown(gallery, { key: 'ArrowRight' });
  expect(light2).toHaveFocus();
  fireEvent.keyDown(gallery, { key: 'ArrowDown' });
  expect(light9).toHaveFocus();
  fireEvent.click(
    within(gallery).getByRole('menuitemradio', {
      name: '应用表格样式：深色 11',
    }),
  );
  expect(tableStyleTrigger).toHaveFocus();

  fireEvent.click(screen.getByRole('button', { name: '首列' }));
  fireEvent.click(screen.getByRole('button', { name: '列条纹' }));
  fireEvent.click(screen.getByRole('button', { name: '转换为区域' }));

  expect(actions).toEqual([
    JSON.stringify({ name: 'Sales_2026' }),
    JSON.stringify({ style: { family: 'dark', number: 11 } }),
    JSON.stringify({ showFirstColumn: true }),
    JSON.stringify({ showColumnStripes: true }),
    'convert',
  ]);
});

function spreadsheetCan(): SpreadsheetEditorCanCommands {
  return {
    activateSheet: () => true,
    addSheet: () => true,
    adjustDecimalPlaces: () => true,
    adjustFontSize: () => true,
    applyCellStyle: () => true,
    applyCellFormat: () => true,
    applyDataValidation: () => true,
    applyAutoSum: () => true,
    applyHyperlink: () => true,
    applyTable: () => true,
    clearSelectedCells: () => true,
    activateFormatPainter: () => true,
    applyFormatPainter: () => true,
    cancelFormatPainter: () => true,
    copySelection: () => true,
    cutSelection: () => true,
    deleteSelectedStructure: () => true,
    deleteSheet: () => true,
    duplicateSheet: () => true,
    fillSelectedCells: () => true,
    hideSheet: () => true,
    insertSelectedStructure: () => true,
    insertCurrentDateTime: () => true,
    mergeSelectedCells: () => true,
    moveSheet: () => true,
    moveSelection: () => true,
    openAutoFilterMenu: () => true,
    openDataValidation: () => true,
    openFind: () => true,
    openFormatCells: () => true,
    openGoTo: () => true,
    openHyperlink: () => true,
    openPasteSpecial: () => true,
    openTable: () => true,
    pasteCells: () => true,
    pasteSelection: () => true,
    pasteSpecial: () => true,
    recalculateFormula: () => true,
    removeHyperlink: () => true,
    removeDataValidation: () => true,
    renameSheet: () => true,
    redo: () => false,
    setCellFormat: () => true,
    setTextOrientation: () => true,
    setFreezePanes: () => true,
    setGridLines: () => true,
    setSelectedCellBorders: () => true,
    setSheetColor: () => true,
    setSelectedStructureHidden: () => true,
    setSelectedStructureSize: () => true,
    setSpreadsheetContent: () => true,
    setZoom: () => true,
    selectCellRange: () => true,
    sortSelectedCells: () => true,
    toggleAutoFilter: () => true,
    updateTable: () => true,
    convertTableToRange: () => true,
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
      | 'adjustDecimalPlaces'
      | 'adjustFontSize'
      | 'activateFormatPainter'
      | 'applyCellStyle'
      | 'applyCellFormat'
      | 'applyDataValidation'
      | 'applyAutoSum'
      | 'applyHyperlink'
      | 'cancelFormatPainter'
      | 'clearSelectedCells'
      | 'copySelection'
      | 'convertTableToRange'
      | 'cutSelection'
      | 'deleteSelectedStructure'
      | 'fillSelectedCells'
      | 'insertSelectedStructure'
      | 'insertCurrentDateTime'
      | 'mergeSelectedCells'
      | 'openFind'
      | 'openDataValidation'
      | 'openFormatCells'
      | 'openGoTo'
      | 'openHyperlink'
      | 'openPasteSpecial'
      | 'openTable'
      | 'pasteSelection'
      | 'pasteSpecial'
      | 'removeDataValidation'
      | 'setFreezePanes'
      | 'setSelectedStructureHidden'
      | 'setTextOrientation'
      | 'toggleAutoFilter'
      | 'updateTable'
    >
  > = {},
): SpreadsheetEditorCommands {
  return {
    activateSheet: () => true,
    activateFormatPainter: overrides.activateFormatPainter ?? (() => true),
    addSheet: () => true,
    adjustDecimalPlaces: overrides.adjustDecimalPlaces ?? (() => true),
    adjustFontSize: overrides.adjustFontSize ?? (() => true),
    applyCellStyle: overrides.applyCellStyle ?? (() => true),
    applyCellFormat: overrides.applyCellFormat ?? (() => true),
    applyDataValidation: overrides.applyDataValidation ?? (() => true),
    applyAutoSum: overrides.applyAutoSum ?? (() => true),
    applyHyperlink: overrides.applyHyperlink ?? (() => true),
    applyTable: () => true,
    applyFormatPainter: () => true,
    cancelFormatPainter: overrides.cancelFormatPainter ?? (() => true),
    clearSelectedCells: overrides.clearSelectedCells ?? (() => true),
    copySelection: overrides.copySelection ?? (() => true),
    cutSelection: overrides.cutSelection ?? (() => true),
    deleteSelectedStructure: overrides.deleteSelectedStructure ?? (() => true),
    deleteSheet: () => true,
    duplicateSheet: () => true,
    fillSelectedCells: overrides.fillSelectedCells ?? (() => true),
    hideSheet: () => true,
    insertSelectedStructure: overrides.insertSelectedStructure ?? (() => true),
    insertCurrentDateTime: overrides.insertCurrentDateTime ?? (() => true),
    mergeSelectedCells: overrides.mergeSelectedCells ?? (() => true),
    moveSheet: () => true,
    moveSelection: () => true,
    openAutoFilterMenu: () => true,
    openDataValidation: overrides.openDataValidation ?? (() => true),
    openFind: overrides.openFind ?? (() => true),
    openFormatCells: overrides.openFormatCells ?? (() => true),
    openGoTo: overrides.openGoTo ?? (() => true),
    openHyperlink: overrides.openHyperlink ?? (() => true),
    openPasteSpecial: overrides.openPasteSpecial ?? (() => true),
    openTable: overrides.openTable ?? (() => true),
    pasteCells: () => true,
    pasteSelection: overrides.pasteSelection ?? (() => true),
    pasteSpecial: overrides.pasteSpecial ?? (() => true),
    recalculateFormula: () => true,
    removeHyperlink: () => true,
    removeDataValidation: overrides.removeDataValidation ?? (() => true),
    renameSheet: () => true,
    redo: () => false,
    setCellFormat,
    setFreezePanes: overrides.setFreezePanes ?? (() => true),
    setGridLines: () => true,
    setSelectedCellBorders: () => true,
    setSheetColor: () => true,
    setSelectedStructureHidden:
      overrides.setSelectedStructureHidden ?? (() => true),
    setTextOrientation: overrides.setTextOrientation ?? (() => true),
    setSelectedStructureSize: () => true,
    setSpreadsheetContent: () => true,
    setZoom: () => true,
    selectCellRange: () => true,
    sortSelectedCells,
    toggleAutoFilter: overrides.toggleAutoFilter ?? (() => true),
    updateTable: overrides.updateTable ?? (() => true),
    convertTableToRange: overrides.convertTableToRange ?? (() => true),
    undo: () => false,
  };
}
