import { expect, type Locator, type Page, test } from '@playwright/test';

test('Spreadsheet routes formatting and history shortcuts from the live grid', async ({
  page,
}) => {
  await openSpreadsheetFixture(page);

  await expectGridFocus(page);
  const nameBox = page.locator('.fortune-name-box');
  const selectedCell = await nameBox.textContent();
  expect(selectedCell).toBeTruthy();

  const shortcuts = [
    { button: page.getByRole('button', { name: '加粗' }), key: 'b' },
    { button: page.getByRole('button', { name: '斜体' }), key: 'i' },
    { button: page.getByRole('button', { name: '下划线' }), key: 'u' },
    { button: page.getByRole('button', { name: '删除线' }), key: '5' },
  ];
  for (const shortcut of shortcuts) await expect(shortcut.button).toBeEnabled();
  for (const modifier of ['Control', 'Meta']) {
    for (const shortcut of shortcuts) {
      const initialState = await shortcut.button.getAttribute('aria-pressed');
      expect(initialState === 'true' || initialState === 'false').toBe(true);
      const toggledState = initialState === 'true' ? 'false' : 'true';
      await page.keyboard.press(`${modifier}+${shortcut.key}`);
      await expect(shortcut.button).toHaveAttribute(
        'aria-pressed',
        toggledState,
      );
      await expectGridFocus(page);
      await page.keyboard.press(`${modifier}+${shortcut.key}`);
      await expect(shortcut.button).toHaveAttribute(
        'aria-pressed',
        initialState ?? '',
      );
    }
  }

  const bold = shortcuts[0].button;
  const initialBold = await bold.getAttribute('aria-pressed');
  expect(initialBold === 'true' || initialBold === 'false').toBe(true);
  const toggledBold = initialBold === 'true' ? 'false' : 'true';
  await page.keyboard.press('Control+b');
  await expect(bold).toHaveAttribute('aria-pressed', toggledBold);
  await page.keyboard.press('Control+z');
  await expect(bold).toHaveAttribute('aria-pressed', initialBold ?? '');
  await expect(nameBox).toHaveText(selectedCell ?? '');
  await page.keyboard.press('Control+y');
  await expect(bold).toHaveAttribute('aria-pressed', toggledBold);
  await expect(nameBox).toHaveText(selectedCell ?? '');
  await expectGridFocus(page);
});

test('Spreadsheet handles editor shortcuts before workbook document listeners can swallow them', async ({
  page,
}) => {
  await openSpreadsheetFixture(page);

  const grid = page.locator('.fortune-sheet-overlay');
  const bold = page.getByRole('button', { name: '加粗' });
  await grid.focus();
  const initialState = await bold.getAttribute('aria-pressed');
  expect(initialState === 'true' || initialState === 'false').toBe(true);
  const toggledState = initialState === 'true' ? 'false' : 'true';

  await page.evaluate(() => {
    const blockWorkbookShortcut = (event: KeyboardEvent) => {
      if (
        event.target instanceof Element &&
        event.target.closest('.fortune-container') &&
        (event.metaKey || event.ctrlKey) &&
        event.key.toLocaleLowerCase() === 'b'
      ) {
        event.stopPropagation();
        document.removeEventListener('keydown', blockWorkbookShortcut, true);
      }
    };
    document.addEventListener('keydown', blockWorkbookShortcut, true);
  });

  const modifier = process.platform === 'darwin' ? 'Meta' : 'Control';
  await page.keyboard.press(`${modifier}+b`);
  await expect(bold).toHaveAttribute('aria-pressed', toggledState);
  await expectGridFocus(page);
});

test('Spreadsheet owns find, match navigation, and focus restoration', async ({
  page,
}) => {
  await openSpreadsheetFixture(page);

  const grid = page.locator('.fortune-sheet-overlay');
  for (const modifier of ['Control', 'Meta']) {
    await grid.focus();
    await page.keyboard.press(`${modifier}+f`);
    const query = page.getByRole('textbox', { name: '查找当前工作表' });
    await expect(query).toBeVisible();
    await expect(query).toBeFocused();
    await query.press('Escape');
    await expect(query).toHaveCount(0);
    await expectGridFocus(page);
  }

  await page.getByRole('button', { name: '查找' }).click();
  const query = page.getByRole('textbox', { name: '查找当前工作表' });
  await expect(query).toBeFocused();
  const findGeometry = await page.evaluate(() => {
    const bar = document.querySelector('.work-spreadsheet-find-bar');
    const columnHeader = document.querySelector('.fortune-col-header');
    if (
      !(bar instanceof HTMLElement) ||
      !(columnHeader instanceof HTMLElement)
    ) {
      throw new Error('Spreadsheet Find geometry is unavailable.');
    }
    return {
      barTop: bar.getBoundingClientRect().top,
      columnHeaderBottom: columnHeader.getBoundingClientRect().bottom,
    };
  });
  expect(findGeometry.barTop).toBeGreaterThanOrEqual(
    findGeometry.columnHeaderBottom,
  );
  await query.fill('新版发布');
  await expect(page.getByText('1 个匹配', { exact: true })).toBeVisible();
  await query.press('Enter');
  await expect(page.locator('.fortune-name-box')).toHaveText('A5');
  await expect(page.getByText('1/1', { exact: true })).toBeVisible();
  await expect(query).toBeFocused();

  const modifier = process.platform === 'darwin' ? 'Meta' : 'Control';
  await page.keyboard.press(`${modifier}+f`);
  await expect(query).toBeFocused();
  await expect
    .poll(() =>
      query.evaluate(
        (input) =>
          input.selectionStart === 0 &&
          input.selectionEnd === input.value.length,
      ),
    )
    .toBe(true);

  await query.press('Escape');
  await expect(query).toHaveCount(0);
  await expectGridFocus(page);
});

test('Spreadsheet restores grid focus and navigates worksheets from the keyboard', async ({
  page,
}) => {
  await openSpreadsheetFixture(page);

  await page.locator('.fortune-sheet-overlay').click();
  await page.keyboard.press('Shift+F11');
  const createdSheet = page.getByRole('tab', { name: '工作表 2' });
  await expect(createdSheet).toHaveAttribute('aria-selected', 'true');
  await expectGridFocus(page);

  const firstSheet = page.getByRole('tab', { name: '执行看板' });
  await page.getByRole('button', { name: '上一个工作表' }).click();
  await expect(firstSheet).toHaveAttribute('aria-selected', 'true');
  await expectGridFocus(page);
  await page.getByRole('button', { name: '下一个工作表' }).click();
  await expect(createdSheet).toHaveAttribute('aria-selected', 'true');
  await expectGridFocus(page);

  await createdSheet.press('ArrowLeft');
  await expect(firstSheet).toHaveAttribute('aria-selected', 'true');
  await expect(firstSheet).toBeFocused();

  await firstSheet.press('End');
  await expect(createdSheet).toHaveAttribute('aria-selected', 'true');
  await expect(createdSheet).toBeFocused();

  await page.keyboard.press('Control+PageUp');
  await expect(firstSheet).toHaveAttribute('aria-selected', 'true');
  await expectGridFocus(page);

  await page.keyboard.press('Meta+PageDown');
  await expect(createdSheet).toHaveAttribute('aria-selected', 'true');
  await expectGridFocus(page);
});

test('Spreadsheet keeps worksheet switching available in read-only preview', async ({
  page,
}) => {
  await openSpreadsheetFixture(page);

  const grid = page.locator('.fortune-sheet-overlay');
  await grid.focus();
  await page.keyboard.press('Shift+F11');
  const firstSheet = page.getByRole('tab', { name: '执行看板' });
  const secondSheet = page.getByRole('tab', { name: '工作表 2' });
  await expect(secondSheet).toHaveAttribute('aria-selected', 'true');

  await page.getByRole('button', { name: '预览' }).click();
  await expect(page.locator('.work-spreadsheet-editor.preview')).toBeVisible();
  await firstSheet.click();
  await expect(firstSheet).toHaveAttribute('aria-selected', 'true');

  await grid.focus();
  await page.keyboard.press('Meta+PageDown');
  await expect(secondSheet).toHaveAttribute('aria-selected', 'true');
  await expectGridFocus(page);

  await page.keyboard.press('Shift+F11');
  await expect(
    page.getByRole('navigation', { name: '工作表' }).getByRole('tab'),
  ).toHaveCount(2);
});

test('Spreadsheet owns deterministic cell navigation, selection, editing, and clipboard shortcuts', async ({
  page,
}) => {
  await openSpreadsheetFixture(page);

  const grid = page.locator('.fortune-sheet-overlay');
  const nameBox = page.locator('.fortune-name-box');
  const formulaBar = page.locator('.fortune-fx-input');
  const selectedRange = page.getByRole('status', {
    name: '表格选区状态',
  });
  await grid.click();
  const initialCell = await nameBox.textContent();
  expect(initialCell).toMatch(/^[A-Z]+\d+$/);
  if (!initialCell) throw new Error('Spreadsheet selection is unavailable.');

  await page.keyboard.press('ArrowRight');
  await expect(nameBox).toHaveText(offsetCellReference(initialCell, 0, 1));
  await page.keyboard.press('ArrowDown');
  await expect(nameBox).toHaveText(offsetCellReference(initialCell, 1, 1));
  await page.keyboard.press('Tab');
  await expect(nameBox).toHaveText(offsetCellReference(initialCell, 1, 2));
  await page.keyboard.press('Shift+Tab');
  await expect(nameBox).toHaveText(offsetCellReference(initialCell, 1, 1));
  await page.keyboard.press('Enter');
  await expect(nameBox).toHaveText(offsetCellReference(initialCell, 2, 1));
  await page.keyboard.press('Shift+Enter');
  await expect(nameBox).toHaveText(offsetCellReference(initialCell, 1, 1));

  const modifier = process.platform === 'darwin' ? 'Meta' : 'Control';
  await page.keyboard.press('Home');
  const rowStart = offsetCellReference(
    initialCell,
    1,
    -cellColumn(initialCell),
  );
  await expect(nameBox).toHaveText(rowStart);
  await page.keyboard.press(`${modifier}+Home`);
  await expect(nameBox).toHaveText('A1');
  await page.keyboard.press(`${modifier}+End`);
  await expect(nameBox).toHaveText('G7');
  await page.keyboard.press('PageDown');
  await expect(nameBox).toHaveText('G27');
  await page.keyboard.press('PageUp');
  await expect(nameBox).toHaveText('G7');

  await page.keyboard.press('Shift+ArrowRight');
  await expect(selectedRange).toHaveText('G7:H7');
  await page.keyboard.press('Control+Space');
  await expect(selectedRange).toHaveText('H1:H40');
  await page.keyboard.press('Shift+Space');
  await expect(selectedRange).toHaveText('A7:L7');
  await page.keyboard.press(`${modifier}+a`);
  await expect(selectedRange).toHaveText('A1:L40');

  await page.keyboard.press('Shift+F11');
  await expect(nameBox).toHaveText('A1');
  await page.keyboard.press('F2');
  await expect
    .poll(() =>
      page.evaluate(() =>
        document.activeElement?.classList.contains('luckysheet-cell-input'),
      ),
    )
    .toBe(true);
  await page.keyboard.press('Escape');
  await expectGridFocus(page);

  await page.keyboard.press('A');
  await page.keyboard.press('3');
  await page.keyboard.press('S');
  await page.keyboard.press('Enter');
  await expect(nameBox).toHaveText('A2');
  await page.keyboard.press('ArrowUp');
  await expect(formulaBar).toHaveText('A3S');
  await page.keyboard.press(`${modifier}+z`);
  await expect(formulaBar).toHaveText('');
  await expectGridFocus(page);

  await page.context().grantPermissions(['clipboard-read', 'clipboard-write'], {
    origin: new URL(page.url()).origin,
  });
  const clipboardModifier = process.platform === 'darwin' ? 'Meta' : 'Control';
  await page.evaluate(() => navigator.clipboard.writeText('A3S\tOffice'));
  await page.keyboard.press(`${clipboardModifier}+v`);
  await expect(formulaBar).toHaveText('A3S');

  await page.keyboard.press(`${clipboardModifier}+c`);
  await expect
    .poll(() => page.evaluate(() => navigator.clipboard.readText()))
    .toBe('A3S\tOffice');
  await page.keyboard.press(`${clipboardModifier}+x`);
  await expect(formulaBar).toHaveText('');
  await page.keyboard.press(`${clipboardModifier}+z`);
  await expect(formulaBar).toHaveText('A3S');
  await expectGridFocus(page);

  await formulaBar.click();
  const formulaText = (await formulaBar.textContent()) ?? '';
  expect(formulaText).toBe('A3S');
  await page.keyboard.press(`${clipboardModifier}+a`);
  await expect
    .poll(() =>
      page.evaluate(() => ({
        activeClass: document.activeElement?.className,
        selection: window.getSelection()?.toString(),
      })),
    )
    .toMatchObject({
      activeClass: expect.stringContaining('fortune-fx-input'),
      selection: formulaText,
    });
  await page.keyboard.press(`${clipboardModifier}+c`);
  await expect
    .poll(() => page.evaluate(() => navigator.clipboard.readText()))
    .toBe(formulaText);
});

test('Spreadsheet keeps copy and paste working when browser clipboard access is blocked', async ({
  page,
}) => {
  await openSpreadsheetFixture(page);

  await page.locator('.fortune-sheet-overlay').focus();
  await page.keyboard.press('Shift+F11');
  const nameBox = page.locator('.fortune-name-box');
  const formulaBar = page.locator('.fortune-fx-input');
  await expect(nameBox).toHaveText('A1');
  await page.evaluate(() => {
    Object.defineProperty(navigator, 'clipboard', {
      configurable: true,
      value: {
        readText: async () => {
          throw new Error('Clipboard read blocked for regression coverage.');
        },
        writeText: async () => {
          throw new Error('Clipboard write blocked for regression coverage.');
        },
      },
    });
  });

  await page.keyboard.press('4');
  await page.keyboard.press('2');
  await page.keyboard.press('Enter');
  await expect(nameBox).toHaveText('A2');
  await page.keyboard.press('ArrowUp');
  const modifier = process.platform === 'darwin' ? 'Meta' : 'Control';
  await page.keyboard.press(`${modifier}+c`);
  await page.keyboard.press('ArrowRight');
  await expect(nameBox).toHaveText('B1');
  await page.keyboard.press(`${modifier}+v`);
  await expect(formulaBar).toHaveText('42');
  await expectGridFocus(page);
});

test('Spreadsheet returns keyboard control to the grid after ribbon menus', async ({
  page,
}) => {
  await openSpreadsheetFixture(page);

  await page.locator('.fortune-sheet-overlay').click();
  const nameBox = page.locator('.fortune-name-box');
  const initialCell = await nameBox.textContent();
  expect(initialCell).toMatch(/^[A-Z]+\d+$/);
  if (!initialCell) throw new Error('Spreadsheet selection is unavailable.');

  await page.getByRole('button', { name: '加粗' }).click();
  await expectGridFocus(page);
  await expect(nameBox).toHaveText(initialCell);

  await page.getByRole('combobox', { name: '字体' }).click();
  await page.getByRole('option', { name: 'Arial' }).click();
  await expectGridFocus(page);
  await expect(nameBox).toHaveText(initialCell);

  await page.keyboard.press('ArrowRight');
  await expect(nameBox).toHaveText(offsetCellReference(initialCell, 0, 1));
  await page.keyboard.press('F2');
  await expect
    .poll(() =>
      page.evaluate(() =>
        document.activeElement?.classList.contains('luckysheet-cell-input'),
      ),
    )
    .toBe(true);
});

test('Spreadsheet uses one accent across the grid and workbook footer', async ({
  page,
}) => {
  await openSpreadsheetFixture(page);

  await expect(page.getByLabel('普通表格视图')).toBeVisible();
  await expect(page.getByRole('button', { name: '普通表格视图' })).toHaveCount(
    0,
  );

  const appearance = await page.evaluate(() => {
    const selectedCell = document.querySelector('.luckysheet-cell-selected');
    const activeSheet = document.querySelector(
      '.work-spreadsheet-sheet-tab.active',
    );
    const zoomFill = document.querySelector('.work-office-slider-fill');
    const formulaBar = document.querySelector('.fortune-fx-editor');
    const columnHeader = document.querySelector('.fortune-col-header');
    const rowHeader = document.querySelector('.fortune-row-header');
    const footer = document.querySelector('.work-spreadsheet-footer');
    if (
      !selectedCell ||
      !activeSheet ||
      !zoomFill ||
      !formulaBar ||
      !columnHeader ||
      !rowHeader ||
      !footer
    ) {
      throw new Error('Spreadsheet accent surfaces are unavailable.');
    }
    return {
      accents: {
        grid: getComputedStyle(selectedCell).borderTopColor,
        sheet: getComputedStyle(activeSheet, '::after').backgroundColor,
        zoom: getComputedStyle(zoomFill).backgroundColor,
      },
      chrome: {
        columnHeader: getComputedStyle(columnHeader).backgroundColor,
        footer: getComputedStyle(footer).backgroundColor,
        formulaBar: getComputedStyle(formulaBar).backgroundColor,
        rowHeader: getComputedStyle(rowHeader).backgroundColor,
      },
    };
  });

  expect(new Set(Object.values(appearance.accents)).size).toBe(1);
  expect(appearance.chrome.columnHeader).toBe('rgba(0, 0, 0, 0)');
  expect(appearance.chrome.rowHeader).toBe('rgba(0, 0, 0, 0)');
  expect(appearance.chrome.formulaBar).toBe(appearance.chrome.footer);

  await page.getByRole('combobox', { name: '字体' }).click();
  const fontMenu = page.getByRole('listbox', { name: '字体' });
  await expect(fontMenu).toBeVisible();
  const selectedFont = fontMenu.getByRole('option', {
    name: 'Aptos',
    exact: true,
  });
  await expect(selectedFont).toHaveAttribute('aria-selected', 'true');
  const controlAppearance = await page.evaluate(() => {
    const menu = document.querySelector('[role="listbox"][aria-label="字体"]');
    const selected = menu?.querySelector(
      '[role="option"][aria-selected="true"]',
    );
    if (!(menu instanceof HTMLElement) || !(selected instanceof HTMLElement)) {
      throw new Error('Spreadsheet font menu is unavailable.');
    }
    return {
      accent: getComputedStyle(menu)
        .getPropertyValue('--work-office-control-accent')
        .trim(),
      selectedColor: getComputedStyle(selected).color,
    };
  });
  expect(controlAppearance).toEqual({
    accent: '#159469',
    selectedColor: 'rgb(21, 148, 105)',
  });
});

test('Spreadsheet keeps displayed cell-menu shortcuts executable', async ({
  page,
}) => {
  await openSpreadsheetFixture(page);

  const grid = page.locator('.fortune-sheet-overlay');
  const formulaBar = page.locator('.fortune-fx-input');
  const cellMenu = page.locator('.workspace-context-menu');
  const initialValue = (await formulaBar.textContent()) ?? '';
  expect(initialValue).not.toBe('');
  await page.context().grantPermissions(['clipboard-read', 'clipboard-write'], {
    origin: new URL(page.url()).origin,
  });

  await grid.focus();
  await grid.press('Shift+F10');
  await expect(cellMenu).toBeVisible();
  const modifier = process.platform === 'darwin' ? 'Meta' : 'Control';
  await page.keyboard.press(`${modifier}+c`);
  await expect(cellMenu).toBeHidden();
  await expect
    .poll(() => page.evaluate(() => navigator.clipboard.readText()))
    .not.toBe('');
  await expectGridFocus(page);

  await grid.press('Shift+F10');
  await expect(cellMenu).toBeVisible();
  await page.keyboard.press('Delete');
  await expect(cellMenu).toBeHidden();
  await expect(formulaBar).toHaveText('');
  await expectGridFocus(page);
  await page.keyboard.press(`${modifier}+z`);
  await expect(formulaBar).toHaveText(initialValue);
});

test('Spreadsheet uses one menu geometry for cells and worksheets', async ({
  page,
}) => {
  await openSpreadsheetFixture(page);

  await page.getByRole('button', { name: '执行看板选项' }).click();
  const worksheetMenu = page.locator('.work-spreadsheet-sheet-popover');
  await expect(worksheetMenu).toBeVisible();
  const worksheetGeometry = await officeMenuGeometry(worksheetMenu);
  await expect(
    worksheetMenu.getByRole('menuitem', { name: '重命名' }),
  ).toBeFocused();
  await page.keyboard.press('ArrowDown');
  await expect(
    worksheetMenu.getByRole('menuitem', { name: '复制工作表' }),
  ).toBeFocused();
  await expect(
    worksheetMenu.getByRole('menuitemradio', { name: '红色标签' }),
  ).toHaveAttribute('aria-checked', 'false');
  await page.keyboard.press('Tab');
  await expect(worksheetMenu).toBeHidden();
  await expect(page.getByRole('button', { name: '缩小表格' })).toBeFocused();

  const sheetTab = page.getByRole('tab', { name: '执行看板' });
  await sheetTab.focus();
  await sheetTab.press('Shift+F10');
  await expect(worksheetMenu).toBeVisible();
  await expect(
    worksheetMenu.getByRole('menuitem', { name: '重命名' }),
  ).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(
    page.getByRole('button', { name: '执行看板选项' }),
  ).toBeFocused();

  await page.locator('.fortune-sheet-overlay').click({ button: 'right' });
  const cellMenu = page.locator('.workspace-context-menu');
  await expect(cellMenu).toBeVisible();
  expect(await officeMenuGeometry(cellMenu)).toEqual(worksheetGeometry);
  await page.keyboard.press('Escape');
  await expect(cellMenu).toBeHidden();

  const grid = page.locator('.fortune-sheet-overlay');
  await grid.focus();
  await grid.press('Shift+F10');
  await expect(cellMenu).toBeVisible();
  await expect(cellMenu.getByRole('menuitem').first()).toBeFocused();
  const [selectionBounds, menuBounds] = await Promise.all([
    page.locator('.luckysheet-cell-selected').boundingBox(),
    cellMenu.boundingBox(),
  ]);
  if (!(selectionBounds && menuBounds)) {
    throw new Error('Spreadsheet keyboard menu geometry is unavailable.');
  }
  expect(menuBounds.x).toBeGreaterThanOrEqual(selectionBounds.x - 1);
  expect(menuBounds.x).toBeLessThanOrEqual(
    selectionBounds.x + selectionBounds.width + 1,
  );
  expect(menuBounds.y).toBeGreaterThanOrEqual(selectionBounds.y - 1);
  expect(menuBounds.y).toBeLessThanOrEqual(
    selectionBounds.y + selectionBounds.height + 1,
  );
  await page.keyboard.press('Escape');
  await expect(cellMenu).toBeHidden();
  await expect(grid).toBeFocused();
});

test('Spreadsheet header menus use A3S chrome, commands, and focus behavior', async ({
  page,
}) => {
  await openSpreadsheetFixture(page);

  const columnHeader = page.locator('.fortune-col-header');
  await columnHeader.hover({ position: { x: 150, y: 12 } });
  const headerTrigger = page.getByRole('button', { name: '列操作' });
  await expect(headerTrigger).toBeVisible();
  await expect(headerTrigger).toHaveAttribute('aria-haspopup', 'menu');
  await headerTrigger.click();

  const menu = page.getByRole('menu', { name: /^列 .* 操作$/ });
  await expect(menu).toBeVisible();
  await expect(menu).toHaveClass(/work-office-context-menu/);
  await expect(menu.getByRole('menuitem', { name: '升序排列' })).toBeVisible();
  await expect(
    menu.getByRole('menuitem', { name: '在左侧插入列' }),
  ).toBeVisible();
  await expect(menu.getByRole('menuitem', { name: '列宽…' })).toBeVisible();
  await expect(
    page.locator('.fortune-context-menu.luckysheet-cols-menu:visible'),
  ).toHaveCount(0);

  await page.keyboard.press('Escape');
  await expect(menu).toBeHidden();
  await expectGridFocus(page);

  await columnHeader.hover({ position: { x: 150, y: 12 } });
  await page.getByRole('button', { name: '列操作' }).click();
  await page.getByRole('menuitem', { name: '列宽…' }).click();
  const sizeDialog = page.getByRole('dialog', { name: '设置列宽' });
  const sizeInput = sizeDialog.getByRole('textbox', {
    name: '列宽（1–2038 像素）',
  });
  await expect(sizeInput).toBeFocused();
  await sizeInput.fill('120');
  await sizeDialog.getByRole('button', { name: '应用' }).click();
  await expect(sizeDialog).toBeHidden();
  await expectGridFocus(page);

  await expect(page.locator('.luckysheet-bottom-controll-row')).toHaveCSS(
    'display',
    'none',
  );
});

test('Spreadsheet explains invalid worksheet names and confirms deletion', async ({
  page,
}) => {
  await openSpreadsheetFixture(page);

  await page.locator('.fortune-sheet-overlay').focus();
  await page.keyboard.press('Shift+F11');
  const createdSheet = page.getByRole('tab', { name: '工作表 2' });
  await expect(createdSheet).toHaveAttribute('aria-selected', 'true');

  await createdSheet.dblclick();
  const renameInput = page.getByRole('textbox', { name: '重命名工作表 2' });
  await renameInput.fill('执行看板');
  await renameInput.press('Enter');
  await expect(renameInput).toHaveValue('执行看板');
  await expect(renameInput).toHaveAttribute('aria-invalid', 'true');
  const renameError = page.getByRole('alert').filter({
    hasText: '名称已存在',
  });
  await expect(renameError).toBeVisible();
  await expect(renameInput).toBeFocused();

  const invalidRenameGeometry = await renameInput.evaluate((input) => {
    const tab = input.closest<HTMLElement>('.work-spreadsheet-sheet-tab');
    const error = tab?.querySelector<HTMLElement>(
      '.work-spreadsheet-sheet-rename-error',
    );
    const tabs = tab?.parentElement;
    if (!(tab && error && tabs)) {
      throw new Error('Worksheet rename feedback is incomplete.');
    }
    const tabBounds = tab.getBoundingClientRect();
    const errorBounds = error.getBoundingClientRect();
    const tabsBounds = tabs.getBoundingClientRect();
    return {
      errorRight: errorBounds.right,
      tabRight: tabBounds.right,
      tabsRight: tabsBounds.right,
      viewportWidth: document.documentElement.clientWidth,
    };
  });
  expect(invalidRenameGeometry.errorRight).toBeLessThanOrEqual(
    invalidRenameGeometry.tabsRight + 1,
  );
  expect(invalidRenameGeometry.tabRight).toBeLessThanOrEqual(
    invalidRenameGeometry.viewportWidth + 1,
  );

  await renameInput.fill('季度看板');
  await expect(renameError).toHaveCount(0);
  await renameInput.press('Enter');
  await expect(page.getByRole('tab', { name: '季度看板' })).toBeFocused();

  const populatedSheet = page.getByRole('tab', { name: '执行看板' });
  const openDeleteDialog = async () => {
    await page.getByRole('button', { name: '执行看板选项' }).click();
    await page.getByRole('menuitem', { name: '删除工作表' }).click();
  };
  await openDeleteDialog();
  const dialog = page.getByRole('dialog', { name: '删除“执行看板”？' });
  await expect(dialog).toHaveAccessibleDescription(
    '工作表及其中的内容将被删除。',
  );
  await expect(dialog.getByRole('button', { name: '取消' })).toBeFocused();
  await expect(dialog.getByRole('button', { name: '删除' })).toHaveClass(
    /danger/,
  );
  await dialog.getByRole('button', { name: '取消' }).click();
  await expect(populatedSheet).toBeFocused();

  await openDeleteDialog();
  await page
    .getByRole('dialog', { name: '删除“执行看板”？' })
    .getByRole('button', { name: '删除' })
    .click();
  await expect(populatedSheet).toHaveCount(0);
  await expect(page.getByRole('tab', { name: '季度看板' })).toHaveAttribute(
    'aria-selected',
    'true',
  );
});

test('Spreadsheet keeps phone worksheet rename feedback readable above footer controls', async ({
  page,
}, testInfo) => {
  test.skip(
    testInfo.project.name !== 'desktop-1280',
    'The phone rename contract only needs one browser project.',
  );
  await page.setViewportSize({ width: 390, height: 700 });
  await openSpreadsheetFixture(page);

  await page.locator('.fortune-sheet-overlay').focus();
  await page.keyboard.press('Shift+F11');
  const createdSheet = page.getByRole('tab', { name: '工作表 2' });
  await createdSheet.dblclick();

  const renameInput = page.getByRole('textbox', { name: '重命名工作表 2' });
  await renameInput.fill('Bad/Name');
  await renameInput.press('Enter');
  const renameError = page.getByRole('alert').filter({
    hasText: '名称不能包含 \\ / ? * [ ] :',
  });
  await expect(renameError).toBeVisible();
  await expect(renameInput).toBeFocused();

  const geometry = await page
    .locator('.work-spreadsheet-footer')
    .evaluate((footer) => {
      const input = footer.querySelector<HTMLInputElement>(
        '.work-spreadsheet-sheet-tab.invalid > input',
      );
      const error = footer.querySelector<HTMLElement>(
        '.work-spreadsheet-sheet-rename-error',
      );
      const status = footer.querySelector<HTMLElement>(
        '.work-spreadsheet-status',
      );
      const tools = footer.querySelector<HTMLElement>(
        '.work-spreadsheet-sheet-tools',
      );
      const otherTab = footer.querySelector<HTMLElement>(
        '.work-spreadsheet-sheet-tab:not(.renaming)',
      );
      if (!(input && error && status && tools && otherTab)) {
        throw new Error('Phone worksheet rename feedback is incomplete.');
      }
      const footerBounds = footer.getBoundingClientRect();
      const inputBounds = input.getBoundingClientRect();
      const errorBounds = error.getBoundingClientRect();
      return {
        errorAlignSelf: getComputedStyle(error).alignSelf,
        errorBottom: errorBounds.bottom,
        errorClientWidth: error.clientWidth,
        errorFontSize: Number.parseFloat(getComputedStyle(error).fontSize),
        errorLeft: errorBounds.left,
        errorRight: errorBounds.right,
        errorScrollWidth: error.scrollWidth,
        errorTextAlign: getComputedStyle(error).textAlign,
        footerBottom: footerBounds.bottom,
        footerHeight: footerBounds.height,
        inputLeft: inputBounds.left,
        inputRight: inputBounds.right,
        otherTabDisplay: getComputedStyle(otherTab).display,
        statusDisplay: getComputedStyle(status).display,
        toolsDisplay: getComputedStyle(tools).display,
        viewportHeight: document.documentElement.clientHeight,
        viewportWidth: document.documentElement.clientWidth,
      };
    });

  expect(geometry.footerHeight).toBeGreaterThanOrEqual(48);
  expect(geometry.statusDisplay).toBe('none');
  expect(geometry.toolsDisplay).toBe('none');
  expect(geometry.otherTabDisplay).toBe('none');
  expect(geometry.errorAlignSelf).toBe('stretch');
  expect(geometry.errorFontSize).toBeGreaterThanOrEqual(11);
  expect(geometry.errorTextAlign).toBe('left');
  expect(geometry.inputLeft).toBeGreaterThanOrEqual(0);
  expect(geometry.inputRight).toBeLessThanOrEqual(geometry.viewportWidth);
  expect(geometry.errorLeft).toBeGreaterThanOrEqual(0);
  expect(geometry.errorRight).toBeLessThanOrEqual(geometry.viewportWidth);
  expect(geometry.errorBottom).toBeLessThanOrEqual(geometry.footerBottom);
  expect(geometry.footerBottom).toBeLessThanOrEqual(geometry.viewportHeight);
  expect(geometry.errorScrollWidth).toBeLessThanOrEqual(
    geometry.errorClientWidth + 1,
  );

  await renameInput.fill('移动端看板');
  await expect(renameError).toHaveCount(0);
  await renameInput.press('Enter');
  await expect(page.getByRole('tab', { name: '移动端看板' })).toBeFocused();
});

test('Spreadsheet keeps the active worksheet visible when the footer compacts to phone width', async ({
  page,
}) => {
  await page.setViewportSize({ width: 768, height: 800 });
  await openSpreadsheetFixture(page);

  await page.locator('.fortune-sheet-overlay').focus();
  for (let index = 0; index < 3; index += 1) {
    await page.keyboard.press('Shift+F11');
  }

  const activeTab = page.getByRole('tab', { name: '工作表 4' });
  await expect(activeTab).toHaveAttribute('aria-selected', 'true');
  await page.setViewportSize({ width: 390, height: 844 });

  await expect
    .poll(() =>
      activeTab.evaluate((tab) => {
        const tabList = tab.closest('[role="tablist"]');
        const tabContainer = tab.closest('.work-spreadsheet-sheet-tab');
        const options = tabContainer?.querySelector(
          '.work-spreadsheet-sheet-options',
        );
        if (
          !(tabList instanceof HTMLElement) ||
          !(tabContainer instanceof HTMLElement) ||
          !(options instanceof HTMLElement)
        ) {
          return false;
        }
        const tabBounds = tabContainer.getBoundingClientRect();
        const tabListBounds = tabList.getBoundingClientRect();
        const optionsBounds = options.getBoundingClientRect();
        const hit = document.elementFromPoint(
          optionsBounds.left + optionsBounds.width / 2,
          optionsBounds.top + optionsBounds.height / 2,
        );
        return (
          tabBounds.left >= tabListBounds.left - 1 &&
          tabBounds.right <= tabListBounds.right + 1 &&
          getComputedStyle(options).opacity === '1' &&
          options.contains(hit)
        );
      }),
    )
    .toBe(true);

  await page.getByRole('button', { name: '工作表 4选项' }).click();
  await expect(
    page.getByRole('menu', { name: '工作表 4工作表操作' }),
  ).toBeVisible();
});

async function openSpreadsheetFixture(page: Page) {
  await page.goto('/');
  await page
    .getByRole('button', {
      name: '季度执行计划 XLSX · 本次会话',
    })
    .click();
  await page.locator('.work-spreadsheet-canvas > .fortune-container').waitFor();
}

async function expectGridFocus(page: Page) {
  await expect
    .poll(() =>
      page.evaluate(() =>
        document.activeElement?.classList.contains('fortune-sheet-overlay'),
      ),
    )
    .toBe(true);
}

async function officeMenuGeometry(locator: Locator) {
  return locator.evaluate((menu) => {
    const button = menu.querySelector('button');
    if (!(button instanceof HTMLElement)) {
      throw new Error('Office menu action is unavailable.');
    }
    const menuStyle = getComputedStyle(menu);
    const buttonStyle = getComputedStyle(button);
    return {
      accent: menuStyle.getPropertyValue('--work-office-menu-accent').trim(),
      borderRadius: menuStyle.borderRadius,
      buttonHeight: button.getBoundingClientRect().height,
      buttonPadding: buttonStyle.padding,
      fontFamily: menuStyle.fontFamily,
      fontSize: menuStyle.fontSize,
      padding: menuStyle.padding,
    };
  });
}

function offsetCellReference(
  reference: string,
  rowOffset: number,
  columnOffset: number,
): string {
  const match = /^([A-Z]+)(\d+)$/.exec(reference);
  if (!match) throw new Error(`Invalid cell reference: ${reference}`);
  let column = 0;
  for (const character of match[1]) {
    column = column * 26 + character.charCodeAt(0) - 64;
  }
  column += columnOffset;
  let label = '';
  while (column > 0) {
    column -= 1;
    label = String.fromCharCode(65 + (column % 26)) + label;
    column = Math.floor(column / 26);
  }
  return `${label}${Number(match[2]) + rowOffset}`;
}

function cellColumn(reference: string): number {
  const match = /^([A-Z]+)\d+$/.exec(reference);
  if (!match) throw new Error(`Invalid cell reference: ${reference}`);
  let column = 0;
  for (const character of match[1]) {
    column = column * 26 + character.charCodeAt(0) - 64;
  }
  return column - 1;
}
