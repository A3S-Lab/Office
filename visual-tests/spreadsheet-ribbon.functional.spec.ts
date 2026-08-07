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

async function openSpreadsheetFixture(page: Page): Promise<void> {
  await page.goto('/');
  await page
    .getByRole('button', {
      name: '季度执行计划 XLSX · 本次会话',
    })
    .click();
  await page.locator('.work-spreadsheet-canvas > .fortune-container').waitFor();
}
