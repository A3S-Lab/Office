import { expect, type Page, test } from '@playwright/test';

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
  const paste = clipboard.getByRole('button', { name: '粘贴' });
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

async function openSpreadsheetFixture(page: Page): Promise<void> {
  await page.goto('/');
  await page
    .getByRole('button', {
      name: '季度执行计划 XLSX · 本次会话',
    })
    .click();
  await page.locator('.work-spreadsheet-canvas > .fortune-container').waitFor();
}
