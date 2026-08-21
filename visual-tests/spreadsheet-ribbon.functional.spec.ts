import { expect, type Page, test } from '@playwright/test';
import { openSpreadsheetFixture } from './visual-test-support';

test('Spreadsheet follows the WPS ribbon information architecture', async ({
  page,
}) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));

  await openSpreadsheetFixture(page);

  const ribbon = page.locator('.work-spreadsheet-ribbon');
  const quickAccess = ribbon.getByRole('toolbar', {
    name: '快速访问工具栏',
  });
  await expect(quickAccess).toBeVisible();
  await expect(
    quickAccess.getByRole('button', { name: '撤销' }),
  ).toHaveAttribute('aria-keyshortcuts', 'Control+Z Meta+Z');
  await expect(
    quickAccess.getByRole('button', { name: '重做' }),
  ).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+Z Meta+Shift+Z Control+Y Meta+Y',
  );

  await expect(ribbon.getByRole('tab', { name: '开始' })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  await expect(ribbon.getByRole('button', { name: '条件格式' })).toBeVisible();

  await ribbon.getByRole('tab', { name: '插入' }).click();
  await expect(ribbon.getByRole('button', { name: '插入图表' })).toBeVisible();
  await expect(ribbon.getByRole('button', { name: /^条件格式/ })).toHaveCount(
    0,
  );

  await ribbon.getByRole('tab', { name: '数据' }).click();
  await expect(ribbon.getByRole('button', { name: '升序' })).toBeVisible();
  await expect(ribbon.getByRole('button', { name: '降序' })).toBeVisible();

  const formulasTab = ribbon.getByRole('tab', { name: '公式' });
  await formulasTab.click();
  const recalculate = ribbon.getByRole('button', { name: '重新计算工作簿' });
  await expect(recalculate).toHaveAttribute('aria-keyshortcuts', 'F9');
  await page.locator('.fortune-sheet-overlay').focus();
  await page.keyboard.press('F9');

  await formulasTab.dblclick();
  await expect(ribbon).toHaveAttribute('data-collapsed', 'true');
  await expect(ribbon.locator('.work-office-ribbon-panel')).toBeHidden();

  await ribbon.getByRole('tab', { name: '数据' }).click();
  await expect(ribbon.locator('.work-office-ribbon-panel')).toBeVisible();
  await expect(ribbon.getByRole('button', { name: '升序' })).toBeVisible();

  await page.locator('.fortune-sheet-overlay').click();
  await expect(ribbon.locator('.work-office-ribbon-panel')).toBeHidden();
  expect(browserErrors).toEqual([]);
});

test('Spreadsheet applies common WPS number formats without changing cell values', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await openSpreadsheetFixture(page);

  const grid = page.locator('.fortune-sheet-overlay');
  const nameBox = page.locator('.fortune-name-box');
  const formulaBar = page.locator('.fortune-fx-input');
  const ribbon = page.locator('.work-spreadsheet-ribbon');
  const numberGroup = ribbon.getByRole('region', { name: '数字' });
  const formatSelect = numberGroup.getByRole('combobox', {
    name: '数字格式',
  });
  const currency = numberGroup.getByRole('button', { name: '货币格式' });
  const percent = numberGroup.getByRole('button', { name: '百分比格式' });

  await grid.focus();
  await page.keyboard.press('Control+Home');
  for (let index = 0; index < 3; index += 1) {
    await page.keyboard.press('ArrowDown');
  }
  for (let index = 0; index < 2; index += 1) {
    await page.keyboard.press('ArrowRight');
  }
  await expect(nameBox).toHaveText('C4');
  await expect(formulaBar).toHaveText('100%');
  await expect(formatSelect).toContainText('百分比');
  await expect(currency).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+$ Meta+Shift+$',
  );
  await expect(percent).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+% Meta+Shift+%',
  );

  await formatSelect.click();
  const formatList = page.getByRole('listbox', { name: '数字格式' });
  await expect(formatList).toBeVisible();
  await expect(formatList.getByRole('option')).toHaveCount(11);
  await page.keyboard.press('End');
  await expect(formatList.getByRole('option', { name: '文本' })).toBeFocused();
  await page.keyboard.press('Home');
  await expect(formatList.getByRole('option', { name: '常规' })).toBeFocused();
  const listBounds = await formatList.boundingBox();
  const viewport = page.viewportSize();
  expect(listBounds).not.toBeNull();
  expect(viewport).not.toBeNull();
  expect((listBounds?.x ?? 0) + (listBounds?.width ?? 0)).toBeLessThanOrEqual(
    viewport?.width ?? 0,
  );
  expect((listBounds?.y ?? 0) + (listBounds?.height ?? 0)).toBeLessThanOrEqual(
    viewport?.height ?? 0,
  );
  await page.screenshot({
    path: testInfo.outputPath('spreadsheet-number-formats.png'),
    animations: 'disabled',
  });
  await formatList.getByRole('option', { name: '货币' }).click();
  await expect(formatSelect).toContainText('货币');
  await expect(formulaBar).toHaveText('1');
  await expect(grid).toBeFocused();

  for (const shortcut of [
    { key: 'Control+Shift+Backquote', label: '常规' },
    { key: 'Control+Shift+Digit1', label: '数字' },
    { key: 'Control+Shift+Digit4', label: '货币' },
    { key: 'Control+Shift+Digit5', label: '百分比' },
    { key: 'Control+Shift+Digit3', label: '短日期' },
    { key: 'Control+Shift+Digit2', label: '时间' },
    { key: 'Control+Shift+Digit6', label: '科学计数' },
  ]) {
    await page.keyboard.press(shortcut.key);
    await expect(formatSelect).toContainText(shortcut.label);
    await expect(grid).toBeFocused();
  }

  await page.keyboard.press('Control+z');
  await expect(formatSelect).toContainText('时间');
  await page.keyboard.press('Control+Shift+Backquote');
  await expect(formatSelect).toContainText('常规');
  await expect(formulaBar).toHaveText('1');
  await expect(grid).toBeFocused();
  expect(browserErrors).toEqual([]);
});

test('Spreadsheet adjusts mixed decimal formats without flattening their families', async ({
  page,
}) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await openSpreadsheetFixture(page);

  const grid = page.locator('.fortune-sheet-overlay');
  const nameBox = page.locator('.fortune-name-box');
  const numberGroup = page
    .locator('.work-spreadsheet-ribbon')
    .getByRole('region', { name: '数字' });
  const formatSelect = numberGroup.getByRole('combobox', {
    name: '数字格式',
  });

  await grid.focus();
  await page.keyboard.press('Shift+F11');
  await expect(nameBox).toHaveText('A1');
  await page.keyboard.type('1234.5');
  await page.keyboard.press('Enter');
  await page.keyboard.press('ArrowUp');
  await numberGroup.getByRole('button', { name: '货币格式' }).click();
  await expect(formatSelect).toContainText('货币');

  await page.keyboard.press('ArrowRight');
  await page.keyboard.type('0.25');
  await page.keyboard.press('Enter');
  await page.keyboard.press('ArrowUp');
  await numberGroup.getByRole('button', { name: '百分比格式' }).click();
  await expect(formatSelect).toContainText('百分比');

  await page.keyboard.press('ArrowRight');
  await page.keyboard.type('45292');
  await page.keyboard.press('Enter');
  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('Control+Shift+Digit3');
  await expect(formatSelect).toContainText('短日期');

  await page.keyboard.press('Control+Home');
  await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.press('Shift+ArrowRight');
  await numberGroup.getByRole('button', { name: '增加小数位' }).click();
  await expect(grid).toBeFocused();

  await page.keyboard.press('Control+Home');
  await expect(formatSelect).toContainText('货币');
  await page.keyboard.press('ArrowRight');
  await expect(formatSelect).toContainText('百分比');
  await page.keyboard.press('ArrowRight');
  await expect(formatSelect).toContainText('短日期');
  await page.keyboard.press('ArrowRight');
  await expect(formatSelect).toContainText('数字');

  await page.keyboard.press('Control+z');
  await expect(formatSelect).toContainText('常规');
  await page.keyboard.press('Control+Home');
  await expect(formatSelect).toContainText('货币');
  await page.keyboard.press('ArrowRight');
  await expect(formatSelect).toContainText('百分比');
  await page.keyboard.press('ArrowRight');
  await expect(formatSelect).toContainText('短日期');
  expect(browserErrors).toEqual([]);
});

test('Spreadsheet steps mixed WPS font sizes through one native batch', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await openSpreadsheetFixture(page);

  const grid = page.locator('.fortune-sheet-overlay');
  const nameBox = page.locator('.fortune-name-box');
  const fontGroup = page
    .locator('.work-spreadsheet-ribbon')
    .getByRole('region', { name: '字体' });
  const fontSize = fontGroup.getByRole('combobox', { name: '字号' });
  const growFont = fontGroup.getByRole('button', { name: '增大字号' });
  const shrinkFont = fontGroup.getByRole('button', { name: '减小字号' });

  await expect(growFont).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+. Meta+Shift+. Control+] Meta+]',
  );
  await expect(shrinkFont).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+, Meta+Shift+, Control+[ Meta+[',
  );

  await grid.focus();
  await page.keyboard.press('Shift+F11');
  await expect(nameBox).toHaveText('A1');
  await fontSize.click();
  await page.getByRole('option', { name: '9', exact: true }).click();
  await expect(grid).toBeFocused();

  await page.keyboard.press('Control+Home');
  await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.press('Shift+ArrowRight');
  await expect(nameBox).toHaveText('A1:C1');
  await growFont.click();
  await expect(grid).toBeFocused();

  await page.keyboard.press('Control+Home');
  await expect(fontSize).toHaveText('10');
  await page.keyboard.press('ArrowRight');
  await expect(fontSize).toHaveText('11');
  await page.keyboard.press('ArrowRight');
  await expect(fontSize).toHaveText('11');
  await page.screenshot({
    path: testInfo.outputPath('spreadsheet-mixed-font-size.png'),
    animations: 'disabled',
  });

  await page.keyboard.press('Control+z');
  await expect(fontSize).toHaveText('10');
  await page.keyboard.press('Control+Home');
  await expect(fontSize).toHaveText('9');
  await page.keyboard.press('ArrowRight');
  await expect(fontSize).toHaveText('10');

  await page.keyboard.press('Control+Home');
  await page.keyboard.press('Control+Shift+Period');
  await expect(fontSize).toHaveText('10');
  await page.keyboard.press('Control+]');
  await expect(fontSize).toHaveText('11');
  await page.keyboard.press('Control+Shift+Comma');
  await expect(fontSize).toHaveText('10');
  await page.keyboard.press('Control+[');
  await expect(fontSize).toHaveText('9');
  await expect(grid).toBeFocused();
  expect(browserErrors).toEqual([]);
});

test('Spreadsheet runs clipboard commands from the WPS Home group', async ({
  page,
}) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await openSpreadsheetFixture(page);
  await page.context().grantPermissions(['clipboard-read', 'clipboard-write'], {
    origin: new URL(page.url()).origin,
  });

  const grid = page.locator('.fortune-sheet-overlay');
  const nameBox = page.locator('.fortune-name-box');
  const formulaBar = page.locator('.fortune-fx-input');
  await grid.focus();
  await page.keyboard.press('Shift+F11');
  await expect(nameBox).toHaveText('A1');
  await page.keyboard.type('A3S');
  await page.keyboard.press('Enter');
  await page.keyboard.press('ArrowUp');
  await expect(formulaBar).toHaveText('A3S');

  const clipboard = page
    .locator('.work-spreadsheet-ribbon')
    .getByRole('region', { name: '剪贴板' });
  const paste = clipboard.getByRole('button', { name: '粘贴', exact: true });
  const cut = clipboard.getByRole('button', { name: '剪切' });
  const copy = clipboard.getByRole('button', { name: '复制' });
  await expect(paste).toHaveAttribute('aria-keyshortcuts', 'Control+V Meta+V');
  await expect(cut).toHaveAttribute('aria-keyshortcuts', 'Control+X Meta+X');
  await expect(copy).toHaveAttribute('aria-keyshortcuts', 'Control+C Meta+C');

  await copy.click();
  await expect
    .poll(() => page.evaluate(() => navigator.clipboard.readText()))
    .toBe('A3S');
  await expect(grid).toBeFocused();

  await page.keyboard.press('ArrowRight');
  await expect(nameBox).toHaveText('B1');
  await paste.click();
  await expect(formulaBar).toHaveText('A3S');
  await expect(grid).toBeFocused();

  await cut.click();
  await expect(formulaBar).toHaveText('');
  await expect(grid).toBeFocused();
  expect(browserErrors).toEqual([]);
});

test('Spreadsheet renders WPS strikethrough from the ribbon and Ctrl+5', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await openSpreadsheetFixture(page);

  const grid = page.locator('.fortune-sheet-overlay');
  const nameBox = page.locator('.fortune-name-box');
  const formulaBar = page.locator('.fortune-fx-input');
  const strike = page.getByRole('button', { name: '删除线' });

  await grid.focus();
  await page.keyboard.press('Shift+F11');
  await page.keyboard.type('WPS strike');
  await page.keyboard.press('Enter');
  await page.keyboard.press('ArrowUp');
  await expect(nameBox).toHaveText('A1');
  await expect(formulaBar).toHaveText('WPS strike');
  await expect(strike).toHaveAttribute('aria-keyshortcuts', 'Control+5 Meta+5');
  await expect(strike).toHaveAttribute('aria-pressed', 'false');

  await strike.click();
  await expect(strike).toHaveAttribute('aria-pressed', 'true');
  await expect(grid).toBeFocused();
  await page.screenshot({
    path: testInfo.outputPath('spreadsheet-strikethrough.png'),
    animations: 'disabled',
  });

  await page.keyboard.press('Control+5');
  await expect(strike).toHaveAttribute('aria-pressed', 'false');
  await page.keyboard.press('Control+5');
  await expect(strike).toHaveAttribute('aria-pressed', 'true');
  await page.keyboard.press('Control+z');
  await expect(strike).toHaveAttribute('aria-pressed', 'false');
  await expect(formulaBar).toHaveText('WPS strike');
  await expect(grid).toBeFocused();
  expect(browserErrors).toEqual([]);
});

test('Spreadsheet renders and undoes native WPS cell borders', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await openSpreadsheetFixture(page);

  const grid = page.locator('.fortune-sheet-overlay');
  const nameBox = page.locator('.fortune-name-box');
  const ribbon = page.locator('.work-spreadsheet-ribbon');
  await grid.focus();
  await page.keyboard.press('Shift+F11');
  await page.keyboard.press('Control+Home');
  await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.press('Shift+ArrowDown');
  await expect(nameBox).toHaveText('A1:C2');

  const borderGeometry = await spreadsheetSelectionCanvasGeometry(page);
  const moreBorders = ribbon.getByRole('button', { name: '更多框线' });
  await expect(moreBorders).toHaveAttribute('aria-haspopup', 'dialog');
  await moreBorders.click();
  const dialog = page.getByRole('dialog', { name: '框线设置' });
  await expect(dialog).toBeVisible();
  const borderMenu = dialog.getByRole('menu', { name: '框线位置' });
  const noBorder = borderMenu.getByRole('menuitemradio', { name: '无框线' });
  const outsideBorder = borderMenu.getByRole('menuitemradio', {
    name: '外侧框线',
  });
  await expect(
    borderMenu.getByRole('menuitemradio', { name: '上框线' }),
  ).toBeFocused();
  await expect(noBorder).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+_ Meta+Shift+_',
  );
  await expect(noBorder.locator('kbd')).toHaveText('Cmd/Ctrl+Shift+_');
  await expect(outsideBorder).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+& Meta+Shift+&',
  );
  await expect(outsideBorder.locator('kbd')).toHaveText('Cmd/Ctrl+Shift+&');

  const dialogBounds = await dialog.boundingBox();
  const viewport = page.viewportSize();
  expect(dialogBounds).not.toBeNull();
  expect(viewport).not.toBeNull();
  expect(
    (dialogBounds?.x ?? 0) + (dialogBounds?.width ?? 0),
  ).toBeLessThanOrEqual(viewport?.width ?? 0);
  expect(
    (dialogBounds?.y ?? 0) + (dialogBounds?.height ?? 0),
  ).toBeLessThanOrEqual(viewport?.height ?? 0);

  await dialog
    .getByRole('combobox', { name: '框线样式' })
    .selectOption('thick');
  const borderColor = dialog.getByLabel('框线颜色');
  await borderColor.fill('#b42318');
  await expect(borderColor).toHaveValue('#b42318');
  await expect(borderColor).toBeFocused();
  await page.keyboard.press('Shift+Tab');
  await expect(
    dialog.getByRole('combobox', { name: '框线样式' }),
  ).toBeFocused();
  await page.keyboard.press('Shift+Tab');
  await expect(
    borderMenu.getByRole('menuitemradio', { name: '上框线' }),
  ).toBeFocused();
  for (let index = 0; index < 6; index += 1) {
    await page.keyboard.press('ArrowDown');
  }
  await expect(
    borderMenu.getByRole('menuitemradio', { name: '外侧框线' }),
  ).toBeFocused();
  await page.keyboard.press('Enter');
  await expect(dialog).toHaveCount(0);
  await expect(grid).toBeFocused();

  await page.keyboard.press('Control+Home');
  await page.keyboard.press('ArrowRight');
  await page.keyboard.press('ArrowRight');
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowDown');
  await expect(nameBox).toHaveText('C3');
  await expect
    .poll(() => spreadsheetCanvasColorCount(page, borderGeometry, '#b42318'))
    .toBeGreaterThan(8);
  await page.screenshot({
    path: testInfo.outputPath('spreadsheet-cell-borders.png'),
    animations: 'disabled',
  });

  const undo = ribbon.getByRole('button', { name: '撤销' });
  await expect(undo).toBeEnabled();
  await undo.click();
  await expect(grid).toBeFocused();
  await expect
    .poll(() => spreadsheetCanvasColorCount(page, borderGeometry, '#b42318'))
    .toBe(0);

  await page.keyboard.press('Control+Home');
  await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.press('Shift+ArrowDown');
  const blackBaseline = await spreadsheetCanvasColorCount(
    page,
    borderGeometry,
    '#000000',
  );
  await page.keyboard.press('Control+Shift+Digit7');
  await expect
    .poll(() => spreadsheetCanvasColorCount(page, borderGeometry, '#000000'))
    .toBeGreaterThan(blackBaseline + 8);
  await page.screenshot({
    path: testInfo.outputPath('spreadsheet-border-shortcut.png'),
    animations: 'disabled',
  });

  await page.keyboard.press('Control+Shift+Minus');
  await expect
    .poll(() => spreadsheetCanvasColorCount(page, borderGeometry, '#000000'))
    .toBe(blackBaseline);
  await page.keyboard.press('Control+z');
  await expect
    .poll(() => spreadsheetCanvasColorCount(page, borderGeometry, '#000000'))
    .toBeGreaterThan(blackBaseline + 8);
  await page.keyboard.press('Control+z');
  await expect
    .poll(() => spreadsheetCanvasColorCount(page, borderGeometry, '#000000'))
    .toBe(blackBaseline);
  await expect(grid).toBeFocused();
  expect(browserErrors).toEqual([]);
});

test('Spreadsheet inserts and deletes rows from the WPS Home cells group', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await openSpreadsheetFixture(page);

  const grid = page.locator('.fortune-sheet-overlay');
  const nameBox = page.locator('.fortune-name-box');
  const formulaBar = page.locator('.fortune-fx-input');
  const ribbon = page.locator('.work-spreadsheet-ribbon');
  const cells = ribbon.getByRole('region', { name: '单元格' });
  const trigger = cells.getByRole('button', { name: '行和列' });

  await grid.focus();
  await page.keyboard.press('Control+Home');
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowDown');
  await expect(nameBox).toHaveText('A4');
  await expect(formulaBar).toHaveText('客户洞察报告');

  await expect(trigger).toHaveAttribute('aria-haspopup', 'menu');
  await trigger.click();
  const menu = page.getByRole('menu', { name: '行和列选项' });
  await expect(menu).toBeVisible();
  const insertAbove = menu.getByRole('menuitem', { name: '在上方插入行' });
  await expect(insertAbove).toBeFocused();
  await expect(menu.getByRole('menuitem')).toHaveCount(10);

  const menuBounds = await menu.boundingBox();
  const viewport = page.viewportSize();
  expect(menuBounds).not.toBeNull();
  expect(viewport).not.toBeNull();
  expect((menuBounds?.x ?? 0) + (menuBounds?.width ?? 0)).toBeLessThanOrEqual(
    viewport?.width ?? 0,
  );
  expect((menuBounds?.y ?? 0) + (menuBounds?.height ?? 0)).toBeLessThanOrEqual(
    viewport?.height ?? 0,
  );
  await page.screenshot({
    path: testInfo.outputPath('rows-and-columns-menu.png'),
    animations: 'disabled',
  });

  await insertAbove.press('Enter');
  await expect(grid).toBeFocused();
  await expect(nameBox).toHaveText('A4:L4');
  await expect(formulaBar).toHaveText('');
  await page.keyboard.press('ArrowDown');
  await expect(nameBox).toHaveText('A5');
  await expect(formulaBar).toHaveText('客户洞察报告');

  await page.keyboard.press('ArrowUp');
  await trigger.click();
  const restoreMenu = page.getByRole('menu', { name: '行和列选项' });
  await restoreMenu
    .getByRole('menuitem', { name: '删除所选行' })
    .press('Enter');
  await expect(grid).toBeFocused();
  await expect(nameBox).toHaveText('A4');
  await expect(formulaBar).toHaveText('客户洞察报告');
  expect(browserErrors).toEqual([]);
});

test('Spreadsheet applies accessible WPS text orientations as one undoable intent', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await openSpreadsheetFixture(page);

  const grid = page.locator('.fortune-sheet-overlay');
  const nameBox = page.locator('.fortune-name-box');
  const ribbon = page.locator('.work-spreadsheet-ribbon');
  const alignment = ribbon.getByRole('region', { name: '对齐' });
  const trigger = alignment.getByRole('button', { name: '文字方向' });

  await grid.focus();
  await page.keyboard.press('Shift+F11');
  await page.keyboard.type('A3S orientation');
  await page.keyboard.press('Enter');
  await page.keyboard.press('ArrowUp');
  await expect(nameBox).toHaveText('A1');
  await expect(trigger).toHaveAttribute('title', '文字方向（当前：横排文字）');
  await expect(trigger).toHaveAttribute('aria-pressed', 'false');

  await trigger.click();
  const menu = page.getByRole('menu', { name: '文字方向选项' });
  const horizontal = menu.getByRole('menuitemradio', { name: '横排文字' });
  const clockwise = menu.getByRole('menuitemradio', { name: '顺时针倾斜' });
  const rotateDown = menu.getByRole('menuitemradio', {
    name: '向下旋转文字',
  });
  await expect(menu.getByRole('menuitemradio')).toHaveCount(6);
  await expect(horizontal).toHaveAttribute('aria-checked', 'true');
  await expect(horizontal).toBeFocused();
  await menu.press('End');
  await expect(rotateDown).toBeFocused();
  await menu.press('Home');
  await expect(horizontal).toBeFocused();

  const menuBounds = await menu.boundingBox();
  const viewport = page.viewportSize();
  expect(menuBounds).not.toBeNull();
  expect(viewport).not.toBeNull();
  expect((menuBounds?.x ?? 0) + (menuBounds?.width ?? 0)).toBeLessThanOrEqual(
    viewport?.width ?? 0,
  );
  expect((menuBounds?.y ?? 0) + (menuBounds?.height ?? 0)).toBeLessThanOrEqual(
    viewport?.height ?? 0,
  );

  await clockwise.click();
  await expect(grid).toBeFocused();
  await expect(trigger).toHaveAttribute('aria-pressed', 'true');
  await expect(trigger).toHaveAttribute(
    'title',
    '文字方向（当前：顺时针倾斜）',
  );
  await page.screenshot({
    path: testInfo.outputPath('spreadsheet-text-orientation.png'),
    animations: 'disabled',
  });

  await trigger.click();
  await expect(
    page
      .getByRole('menu', { name: '文字方向选项' })
      .getByRole('menuitemradio', { name: '顺时针倾斜' }),
  ).toHaveAttribute('aria-checked', 'true');
  await page.keyboard.press('Escape');
  await grid.focus();
  await page.keyboard.press('Control+z');
  await expect(trigger).toHaveAttribute('aria-pressed', 'false');
  await expect(trigger).toHaveAttribute('title', '文字方向（当前：横排文字）');
  await expect(grid).toBeFocused();
  expect(browserErrors).toEqual([]);
});

test('Spreadsheet hides and unhides selected rows with WPS menu and grid shortcuts', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await openSpreadsheetFixture(page);

  const grid = page.locator('.fortune-sheet-overlay');
  const ribbon = page.locator('.work-spreadsheet-ribbon');
  const trigger = ribbon
    .getByRole('region', { name: '单元格' })
    .getByRole('button', { name: '行和列' });

  await grid.focus();
  await page.keyboard.press('Control+Home');
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowDown');
  await expect(page.locator('.fortune-name-box')).toHaveText('A4');
  const visibleCanvasDigest = await spreadsheetCanvasDigest(page);

  await trigger.click();
  const menu = page.getByRole('menu', { name: '行和列选项' });
  const hideRows = menu.getByRole('menuitem', {
    name: '隐藏所选行',
    exact: true,
  });
  const hideColumns = menu.getByRole('menuitem', {
    name: '隐藏所选列',
    exact: true,
  });
  const unhideRows = menu.getByRole('menuitem', {
    name: '取消隐藏所选行',
    exact: true,
  });
  const unhideColumns = menu.getByRole('menuitem', {
    name: '取消隐藏所选列',
    exact: true,
  });
  await expect(hideRows).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+9 Meta+9',
  );
  await expect(hideColumns).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+0 Meta+0',
  );
  await expect(unhideRows).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+9 Meta+Shift+9',
  );
  await expect(unhideColumns).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+0 Meta+Shift+0',
  );
  await expect(hideRows.locator('kbd')).toHaveText('Cmd/Ctrl+9');
  await expect(unhideRows.locator('kbd')).toHaveText('Cmd/Ctrl+Shift+9');
  await page.screenshot({
    path: testInfo.outputPath('spreadsheet-row-column-visibility.png'),
    animations: 'disabled',
  });

  await page.keyboard.press('Escape');
  await grid.focus();
  await page.keyboard.press('Control+9');
  await expect(grid).toBeFocused();
  await expect
    .poll(() => spreadsheetCanvasDigest(page))
    .not.toBe(visibleCanvasDigest);
  await page.keyboard.press('Control+Shift+9');
  await expect(grid).toBeFocused();
  await expect
    .poll(() => spreadsheetCanvasDigest(page))
    .toBe(visibleCanvasDigest);
  await page.keyboard.press('Control+0');
  await expect
    .poll(() => spreadsheetCanvasDigest(page))
    .not.toBe(visibleCanvasDigest);
  await page.keyboard.press('Control+Shift+0');
  await expect
    .poll(() => spreadsheetCanvasDigest(page))
    .toBe(visibleCanvasDigest);
  await expect(grid).toBeFocused();
  expect(browserErrors).toEqual([]);
});

test('Spreadsheet merges and fills cells from the WPS Home alignment group', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await openSpreadsheetFixture(page);

  const grid = page.locator('.fortune-sheet-overlay');
  const nameBox = page.locator('.fortune-name-box');
  const formulaBar = page.locator('.fortune-fx-input');
  const alignment = page
    .locator('.work-spreadsheet-ribbon')
    .getByRole('region', { name: '对齐' });
  const mergeAndCenter = alignment.getByRole('button', {
    name: '合并居中',
  });
  const moreMergeOptions = alignment.getByRole('button', {
    name: '更多合并方式',
  });

  await grid.focus();
  await page.keyboard.press('Shift+F11');
  await page.keyboard.type('North');
  await page.keyboard.press('Enter');
  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.press('Shift+ArrowDown');
  await expect(nameBox).toHaveText('A1:C2');

  await expect(mergeAndCenter).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+M',
  );
  await mergeAndCenter.click();
  await expect(grid).toBeFocused();

  await moreMergeOptions.click();
  const menu = page.getByRole('menu', { name: '合并选项' });
  await expect(menu).toBeVisible();
  await expect(menu.getByRole('menuitem')).toHaveCount(5);
  await expect(menu.getByRole('menuitem', { name: '合并居中' })).toBeDisabled();
  const unmergeAndFill = menu.getByRole('menuitem', {
    name: '取消合并并填充',
  });
  await expect(unmergeAndFill).toBeEnabled();

  const menuBounds = await menu.boundingBox();
  const viewport = page.viewportSize();
  expect(menuBounds).not.toBeNull();
  expect(viewport).not.toBeNull();
  expect((menuBounds?.x ?? 0) + (menuBounds?.width ?? 0)).toBeLessThanOrEqual(
    viewport?.width ?? 0,
  );
  expect((menuBounds?.y ?? 0) + (menuBounds?.height ?? 0)).toBeLessThanOrEqual(
    viewport?.height ?? 0,
  );
  await page.screenshot({
    path: testInfo.outputPath('merge-menu.png'),
    animations: 'disabled',
  });

  await menu.press('End');
  await expect(unmergeAndFill).toBeFocused();
  await unmergeAndFill.press('Enter');
  await expect(grid).toBeFocused();
  await page.keyboard.press('Control+Home');
  await page.keyboard.press('ArrowRight');
  await expect(nameBox).toHaveText('B1');
  await expect(formulaBar).toHaveText('North');

  await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.press('Shift+ArrowDown');
  await page.keyboard.press('Control+M');
  await expect(grid).toBeFocused();
  await moreMergeOptions.click();
  await expect(
    page.getByRole('menu', { name: '合并选项' }).getByRole('menuitem', {
      name: '取消合并单元格',
    }),
  ).toBeEnabled();
  await page.keyboard.press('Escape');
  await expect(moreMergeOptions).toBeFocused();
  expect(browserErrors).toEqual([]);
});

test('Spreadsheet clears cell state from the WPS Home editing group', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await openSpreadsheetFixture(page);

  const grid = page.locator('.fortune-sheet-overlay');
  const nameBox = page.locator('.fortune-name-box');
  const formulaBar = page.locator('.fortune-fx-input');
  const ribbon = page.locator('.work-spreadsheet-ribbon');
  const bold = ribbon.getByRole('button', { name: '加粗' });
  const clear = ribbon.getByRole('button', { name: '清除', exact: true });

  await grid.focus();
  await page.keyboard.press('Shift+F11');
  await page.keyboard.type('Keep style');
  await page.keyboard.press('Enter');
  await page.keyboard.press('ArrowUp');
  await expect(nameBox).toHaveText('A1');
  await bold.click();
  await expect(grid).toBeFocused();
  await expect(bold).toHaveAttribute('aria-pressed', 'true');

  await expect(clear).toHaveAttribute('aria-haspopup', 'menu');
  await clear.click();
  const menu = page.getByRole('menu', { name: '清除选项' });
  await expect(menu).toBeVisible();
  await expect(menu.getByRole('menuitem')).toHaveCount(5);
  const clearAll = menu.getByRole('menuitem', { name: '清除全部' });
  const clearContents = menu.getByRole('menuitem', { name: '清除内容' });
  const clearHyperlinks = menu.getByRole('menuitem', { name: '清除超链接' });
  await expect(clearAll).toBeFocused();
  await expect(clearContents).toHaveAttribute(
    'aria-keyshortcuts',
    'Delete Backspace',
  );

  const menuBounds = await menu.boundingBox();
  const viewport = page.viewportSize();
  expect(menuBounds).not.toBeNull();
  expect(viewport).not.toBeNull();
  expect((menuBounds?.x ?? 0) + (menuBounds?.width ?? 0)).toBeLessThanOrEqual(
    viewport?.width ?? 0,
  );
  expect((menuBounds?.y ?? 0) + (menuBounds?.height ?? 0)).toBeLessThanOrEqual(
    viewport?.height ?? 0,
  );
  await page.screenshot({
    path: testInfo.outputPath('clear-menu.png'),
    animations: 'disabled',
  });

  await menu.press('End');
  await expect(clearHyperlinks).toBeFocused();
  await menu.press('Home');
  await expect(clearAll).toBeFocused();
  await clearContents.click();
  await expect(grid).toBeFocused();
  await expect(formulaBar).toHaveText('');
  await expect(bold).toHaveAttribute('aria-pressed', 'true');

  await page.keyboard.type('Keep value');
  await page.keyboard.press('Enter');
  await page.keyboard.press('ArrowUp');
  await clear.click();
  await page
    .getByRole('menu', { name: '清除选项' })
    .getByRole('menuitem', { name: '清除格式' })
    .press('Enter');
  await expect(grid).toBeFocused();
  await expect(formulaBar).toHaveText('Keep value');
  await expect(bold).toHaveAttribute('aria-pressed', 'false');

  await page.keyboard.type('Backspace path');
  await page.keyboard.press('Enter');
  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('Backspace');
  await expect(formulaBar).toHaveText('');
  await page.keyboard.type('Delete path');
  await page.keyboard.press('Enter');
  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('Delete');
  await expect(formulaBar).toHaveText('');
  expect(browserErrors).toEqual([]);
});

test('Spreadsheet fills formulas and styles with WPS shortcuts and one-step history', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await openSpreadsheetFixture(page);

  const grid = page.locator('.fortune-sheet-overlay');
  const nameBox = page.locator('.fortune-name-box');
  const formulaBar = page.locator('.fortune-fx-input');
  const ribbon = page.locator('.work-spreadsheet-ribbon');
  const bold = ribbon.getByRole('button', { name: '加粗' });
  const fill = ribbon.getByRole('button', { name: '填充', exact: true });

  await grid.focus();
  await page.keyboard.press('Shift+F11');
  await expect(nameBox).toHaveText('A1');
  await page.keyboard.type('2');
  await page.keyboard.press('Enter');
  await page.keyboard.type('4');
  await page.keyboard.press('Enter');
  await page.keyboard.press('Control+Home');
  await page.keyboard.press('ArrowRight');
  await page.keyboard.type('=A1*3');
  await page.keyboard.press('Enter');
  await page.keyboard.press('ArrowUp');
  await expect(nameBox).toHaveText('B1');
  await expect(formulaBar).toHaveText('=A1*3');
  await bold.click();
  await expect(grid).toBeFocused();

  await page.keyboard.press('Shift+ArrowDown');
  await page.keyboard.press('Shift+ArrowDown');
  await expect(nameBox).toHaveText('B1:B3');
  await expect(fill).toBeEnabled();
  await expect(fill).toHaveAttribute('aria-haspopup', 'menu');
  await fill.click();
  const fillMenu = page.getByRole('menu', { name: '填充选项' });
  const fillDown = fillMenu.getByRole('menuitem', { name: '向下填充' });
  const fillUp = fillMenu.getByRole('menuitem', { name: '向上填充' });
  await expect(fillMenu.getByRole('menuitem')).toHaveCount(4);
  await expect(fillDown).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+D Meta+D',
  );
  await expect(
    fillMenu.getByRole('menuitem', { name: '向右填充' }),
  ).toHaveAttribute('aria-keyshortcuts', 'Control+R Meta+R');
  await expect(fillDown).toBeFocused();
  await fillMenu.press('End');
  await expect(fillUp).toBeFocused();
  await fillMenu.press('Home');
  await expect(fillDown).toBeFocused();
  await page.screenshot({
    path: testInfo.outputPath('spreadsheet-fill-menu.png'),
    animations: 'disabled',
  });
  await page.keyboard.press('Escape');
  await expect(fill).toBeFocused();

  await grid.focus();
  await page.keyboard.press('Control+d');
  await expect(grid).toBeFocused();
  await page.keyboard.press('Control+Home');
  await page.keyboard.press('ArrowRight');
  await page.keyboard.press('ArrowDown');
  await expect(nameBox).toHaveText('B2');
  await expect(formulaBar).toHaveText('=A2*3');
  await expect(bold).toHaveAttribute('aria-pressed', 'true');
  await page.keyboard.press('ArrowDown');
  await expect(nameBox).toHaveText('B3');
  await expect(formulaBar).toHaveText('=A3*3');
  await expect(bold).toHaveAttribute('aria-pressed', 'true');

  await page.keyboard.press('Control+z');
  await expect(formulaBar).toHaveText('');
  await expect(bold).toHaveAttribute('aria-pressed', 'false');
  await page.keyboard.press('ArrowUp');
  await expect(nameBox).toHaveText('B2');
  await expect(formulaBar).toHaveText('');
  await page.keyboard.press('Control+y');
  await expect(formulaBar).toHaveText('=A2*3');
  await expect(bold).toHaveAttribute('aria-pressed', 'true');

  await page.keyboard.press('Control+Home');
  for (let index = 0; index < 4; index += 1) {
    await page.keyboard.press('ArrowDown');
  }
  await expect(nameBox).toHaveText('A5');
  await page.keyboard.type('Right edge');
  await page.keyboard.press('Enter');
  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.press('Shift+ArrowRight');
  await expect(nameBox).toHaveText('A5:C5');
  await page.keyboard.press('Control+r');
  await expect(grid).toBeFocused();
  await page.keyboard.press('Control+Home');
  for (let index = 0; index < 4; index += 1) {
    await page.keyboard.press('ArrowDown');
  }
  await page.keyboard.press('ArrowRight');
  await expect(nameBox).toHaveText('B5');
  await expect(formulaBar).toHaveText('Right edge');
  await page.keyboard.press('ArrowRight');
  await expect(formulaBar).toHaveText('Right edge');
  await page.keyboard.press('Control+z');
  await expect(formulaBar).toHaveText('');
  await page.keyboard.press('ArrowLeft');
  await expect(formulaBar).toHaveText('');

  await page.keyboard.press('Control+Home');
  await page.evaluate(() => {
    (
      window as unknown as { __a3sFillShortcutPageState?: string }
    ).__a3sFillShortcutPageState = 'retained';
  });
  await page.keyboard.press('Control+r');
  await expect(nameBox).toHaveText('A1');
  expect(
    await page.evaluate(
      () =>
        (window as unknown as { __a3sFillShortcutPageState?: string })
          .__a3sFillShortcutPageState,
    ),
  ).toBe('retained');
  expect(browserErrors).toEqual([]);
});

test('Spreadsheet applies one-shot and locked WPS format-painter patterns', async ({
  page,
}) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await openSpreadsheetFixture(page);

  const editor = page.locator('.work-spreadsheet-editor');
  const grid = page.locator('.fortune-sheet-overlay');
  const nameBox = page.locator('.fortune-name-box');
  const ribbon = page.locator('.work-spreadsheet-ribbon');
  const bold = ribbon.getByRole('button', { name: '加粗' });
  const formatPainter = ribbon.getByRole('button', { name: '格式刷' });

  await grid.focus();
  await page.keyboard.press('Shift+F11');
  await expect(nameBox).toHaveText('A1');
  await page.keyboard.type('Source');
  await page.keyboard.press('Enter');
  await page.keyboard.press('ArrowUp');
  await bold.click();
  await expect(grid).toBeFocused();
  await expect(bold).toHaveAttribute('aria-pressed', 'true');

  await page.keyboard.press('ArrowRight');
  await expect(nameBox).toHaveText('B1');
  await page.keyboard.press('ArrowRight');
  await page.keyboard.press('Shift+ArrowRight');
  await bold.click();
  await expect(grid).toBeFocused();

  await page.keyboard.press('Control+Home');
  await page.keyboard.press('Shift+ArrowRight');
  await formatPainter.click();
  await expect(formatPainter).toHaveAttribute('aria-pressed', 'true');
  await expect(editor).toHaveAttribute('data-format-painter', 'once');
  await expect(grid).toHaveCSS('cursor', 'copy');
  await expect(grid).toBeFocused();

  await page.keyboard.press('ArrowRight');
  await expect(nameBox).toHaveText('C1');
  await expect(formatPainter).toHaveAttribute('aria-pressed', 'false');
  await expect(editor).not.toHaveAttribute('data-format-painter');
  await expect(bold).toHaveAttribute('aria-pressed', 'true');
  await page.keyboard.press('ArrowRight');
  await expect(nameBox).toHaveText('D1');
  await expect(bold).toHaveAttribute('aria-pressed', 'false');

  await page.keyboard.press('ArrowLeft');
  await formatPainter.dblclick();
  await expect(formatPainter).toHaveAttribute('aria-pressed', 'true');
  await expect(formatPainter).toContainText('连续');
  await expect(editor).toHaveAttribute('data-format-painter', 'locked');
  await page.keyboard.press('ArrowDown');
  await expect(nameBox).toHaveText('C2');
  await expect(formatPainter).toHaveAttribute('aria-pressed', 'true');
  await expect(bold).toHaveAttribute('aria-pressed', 'true');
  await page.keyboard.press('ArrowDown');
  await expect(nameBox).toHaveText('C3');
  await expect(formatPainter).toHaveAttribute('aria-pressed', 'true');

  await page.keyboard.press('Escape');
  await expect(formatPainter).toHaveAttribute('aria-pressed', 'false');
  await expect(editor).not.toHaveAttribute('data-format-painter');
  await expect(grid).toBeFocused();
  expect(browserErrors).toEqual([]);
});

test('Spreadsheet follows WPS AutoFilter range and keyboard habits', async ({
  page,
}) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await openSpreadsheetFixture(page);

  const editor = page.locator('.work-spreadsheet-editor');
  const canvas = page.locator('.work-spreadsheet-canvas');
  const grid = page.locator('.fortune-sheet-overlay');
  const nameBox = page.locator('.fortune-name-box');
  const ribbon = page.locator('.work-spreadsheet-ribbon');
  await grid.focus();
  await page.keyboard.press('Control+Home');
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowDown');
  await expect(nameBox).toHaveText('A3');

  await ribbon.getByRole('tab', { name: '数据' }).click();
  const filter = ribbon.getByRole('button', { name: '自动筛选' });
  await expect(filter).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+L Meta+Shift+L',
  );
  await filter.click();
  await expect(editor).toHaveAttribute('data-auto-filter', 'active');
  await expect(filter).toHaveAttribute('aria-pressed', 'true');
  await expect(canvas.locator('.luckysheet-filter-options')).toHaveCount(7);
  await expect(canvas.getByRole('button', { name: '状态 筛选' })).toBeVisible();

  await expect(grid).toBeFocused();
  await page.keyboard.press('Control+Shift+L');
  await expect(editor).not.toHaveAttribute('data-auto-filter');
  await expect(canvas.locator('.luckysheet-filter-options')).toHaveCount(0);
  await expect(grid).toBeFocused();
  await page.keyboard.press('Control+Shift+L');
  await expect(editor).toHaveAttribute('data-auto-filter', 'active');
  await expect(canvas.locator('.luckysheet-filter-options')).toHaveCount(7);

  for (let index = 0; index < 6; index += 1) {
    await page.keyboard.press('ArrowRight');
  }
  await expect(nameBox).toHaveText('G3');
  await page.keyboard.press('Alt+ArrowDown');
  const dialog = canvas.getByRole('dialog', { name: '状态 筛选' });
  await expect(dialog).toBeVisible();
  await expect(
    dialog.getByRole('textbox', { name: '搜索筛选值' }),
  ).toBeVisible();
  const dialogBounds = await dialog.boundingBox();
  const viewport = page.viewportSize();
  expect(dialogBounds).not.toBeNull();
  expect(viewport).not.toBeNull();
  expect(
    (dialogBounds?.x ?? 0) + (dialogBounds?.width ?? 0),
  ).toBeLessThanOrEqual(viewport?.width ?? 0);
  expect(
    (dialogBounds?.y ?? 0) + (dialogBounds?.height ?? 0),
  ).toBeLessThanOrEqual(viewport?.height ?? 0);

  await dialog
    .getByRole('button', { name: '清除', exact: true })
    .press('Enter');
  const atRisk = dialog.getByRole('checkbox', { name: '显示 有风险' });
  await atRisk.focus();
  await page.keyboard.press('Space');
  await expect(atRisk).toBeChecked();
  await dialog.getByRole('button', { name: /确\s*认/ }).press('Enter');
  await expect(dialog).toHaveCount(0);
  await expect(canvas.getByRole('button', { name: '状态 筛选' })).toHaveClass(
    /luckysheet-filter-options-active/,
  );

  await grid.focus();
  await page.keyboard.press('Control+Shift+L');
  await expect(editor).not.toHaveAttribute('data-auto-filter');
  await page.keyboard.press('Control+Shift+L');
  await expect(editor).toHaveAttribute('data-auto-filter', 'active');
  await page.keyboard.press('Alt+ArrowDown');
  const resetDialog = canvas.getByRole('dialog', { name: '状态 筛选' });
  await expect(resetDialog).toBeVisible();
  for (const checkbox of await resetDialog.getByRole('checkbox').all()) {
    await expect(checkbox).toBeChecked();
  }
  await page.keyboard.press('Escape');
  await expect(resetDialog).toHaveCount(0);
  await expect(canvas.getByRole('button', { name: '状态 筛选' })).toBeFocused();
  expect(browserErrors).toEqual([]);
});

test('Spreadsheet freezes panes from the WPS View window group', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await openSpreadsheetFixture(page);

  const editor = page.locator('.work-spreadsheet-editor');
  const grid = page.locator('.fortune-sheet-overlay');
  const nameBox = page.locator('.fortune-name-box');
  const ribbon = page.locator('.work-spreadsheet-ribbon');
  await grid.focus();
  await page.keyboard.press('Control+Home');
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowRight');
  await expect(nameBox).toHaveText('B3');

  await ribbon.getByRole('tab', { name: '视图' }).click();
  const trigger = ribbon.getByRole('button', { name: '冻结窗格' });
  await expect(trigger).toHaveAttribute('aria-haspopup', 'menu');
  await trigger.click();
  const menu = page.getByRole('menu', { name: '冻结窗格选项' });
  await expect(menu).toBeVisible();
  const customFreeze = menu.getByRole('menuitem', {
    name: '冻结至第 2 行、A 列',
  });
  await expect(customFreeze).toBeFocused();

  const menuBounds = await menu.boundingBox();
  const viewport = page.viewportSize();
  expect(menuBounds).not.toBeNull();
  expect(viewport).not.toBeNull();
  expect((menuBounds?.x ?? 0) + (menuBounds?.width ?? 0)).toBeLessThanOrEqual(
    viewport?.width ?? 0,
  );
  expect((menuBounds?.y ?? 0) + (menuBounds?.height ?? 0)).toBeLessThanOrEqual(
    viewport?.height ?? 0,
  );

  await customFreeze.press('Enter');
  await expect(editor).toHaveAttribute('data-freeze-panes', 'active');
  await expect(trigger).toHaveAttribute('aria-pressed', 'true');
  await expect(grid).toBeFocused();
  expect(
    await page
      .locator('.fortune-cols-freeze-handle')
      .evaluate((element) => Number.parseFloat(element.style.left)),
  ).toBeGreaterThan(0);
  expect(
    await page
      .locator('.fortune-rows-freeze-handle')
      .evaluate((element) => Number.parseFloat(element.style.top)),
  ).toBeGreaterThan(0);

  await trigger.click();
  const activeMenu = page.getByRole('menu', { name: '冻结窗格选项' });
  const unfreeze = activeMenu.getByRole('menuitem', { name: '取消冻结窗格' });
  const firstColumn = activeMenu.getByRole('menuitem', { name: '冻结首列' });
  await expect(unfreeze).toBeFocused();
  await page.screenshot({
    path: testInfo.outputPath('freeze-panes-active-menu.png'),
    animations: 'disabled',
  });
  await activeMenu.press('End');
  await expect(firstColumn).toBeFocused();
  await activeMenu.press('Home');
  await expect(unfreeze).toBeFocused();
  await page.keyboard.press('Enter');
  await expect(editor).not.toHaveAttribute('data-freeze-panes');
  await expect(trigger).toHaveAttribute('aria-pressed', 'false');
  await expect(grid).toBeFocused();
  expect(browserErrors).toEqual([]);
});

interface SpreadsheetCanvasGeometry {
  bottom: number;
  left: number;
  right: number;
  top: number;
}

async function spreadsheetSelectionCanvasGeometry(
  page: Page,
): Promise<SpreadsheetCanvasGeometry> {
  return page
    .locator('.luckysheet-cell-selected')
    .last()
    .evaluate((selection) => {
      const canvas = document.querySelector<HTMLCanvasElement>(
        '.fortune-sheet-canvas',
      );
      if (!canvas) throw new Error('Spreadsheet canvas is unavailable.');
      const selectionBounds = selection.getBoundingClientRect();
      const canvasBounds = canvas.getBoundingClientRect();
      return {
        bottom: selectionBounds.bottom - canvasBounds.top,
        left: selectionBounds.left - canvasBounds.left,
        right: selectionBounds.right - canvasBounds.left,
        top: selectionBounds.top - canvasBounds.top,
      };
    });
}

async function spreadsheetCanvasColorCount(
  page: Page,
  geometry: SpreadsheetCanvasGeometry,
  color: string,
): Promise<number> {
  return page.evaluate(
    ({ bounds, expected }) => {
      const canvas = document.querySelector<HTMLCanvasElement>(
        '.fortune-sheet-canvas',
      );
      if (!canvas) return 0;
      const context = canvas.getContext('2d');
      const canvasBounds = canvas.getBoundingClientRect();
      if (!context || !canvasBounds.width || !canvasBounds.height) return 0;
      const scaleX = canvas.width / canvasBounds.width;
      const scaleY = canvas.height / canvasBounds.height;
      const red = Number.parseInt(expected.slice(1, 3), 16);
      const green = Number.parseInt(expected.slice(3, 5), 16);
      const blue = Number.parseInt(expected.slice(5, 7), 16);
      const points: Array<[number, number]> = [];
      for (let x = bounds.left; x <= bounds.right; x += 2) {
        points.push([x, bounds.top], [x, bounds.bottom]);
      }
      for (let y = bounds.top; y <= bounds.bottom; y += 2) {
        points.push([bounds.left, y], [bounds.right, y]);
      }
      let matches = 0;
      for (const [x, y] of points) {
        for (let offsetX = -2; offsetX <= 2; offsetX += 1) {
          for (let offsetY = -2; offsetY <= 2; offsetY += 1) {
            const pixel = context.getImageData(
              Math.max(0, Math.round((x + offsetX) * scaleX)),
              Math.max(0, Math.round((y + offsetY) * scaleY)),
              1,
              1,
            ).data;
            if (
              Math.abs(pixel[0] - red) <= 8 &&
              Math.abs(pixel[1] - green) <= 8 &&
              Math.abs(pixel[2] - blue) <= 8 &&
              pixel[3] > 200
            ) {
              matches += 1;
              break;
            }
          }
        }
      }
      return matches;
    },
    { bounds: geometry, expected: color },
  );
}

async function spreadsheetCanvasDigest(page: Page): Promise<number> {
  return page.evaluate(() => {
    const canvas = document.querySelector<HTMLCanvasElement>(
      '.fortune-sheet-canvas',
    );
    const context = canvas?.getContext('2d');
    if (!canvas || !context) return 0;
    const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data;
    let digest = 2_166_136_261;
    for (let index = 0; index < pixels.length; index += 16) {
      digest ^=
        (pixels[index] ?? 0) ^
        ((pixels[index + 1] ?? 0) << 8) ^
        ((pixels[index + 2] ?? 0) << 16);
      digest = Math.imul(digest, 16_777_619);
    }
    return digest >>> 0;
  });
}
