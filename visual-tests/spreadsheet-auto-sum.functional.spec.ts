import { expect, test } from '@playwright/test';
import { openSpreadsheetFixture } from './visual-test-support';

test('Spreadsheet applies WPS AutoSum formulas from the shortcut and split menu', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await openSpreadsheetFixture(page);

  const ribbon = page.locator('.work-spreadsheet-ribbon');
  const toolbar = ribbon.getByRole('toolbar', { name: '开始工具栏' });
  const editing = ribbon.getByRole('region', { name: '编辑' });
  const primary = editing.getByRole('button', { name: '自动求和' });
  const disclosure = editing.getByRole('button', {
    name: '更多自动计算方式',
  });
  const grid = page.locator('.fortune-sheet-overlay');
  const nameBox = page.locator('.fortune-name-box');
  const formulaBar = page.locator('.fortune-fx-input');
  const bold = ribbon.getByRole('button', { name: '加粗' });

  await expect(primary).toBeVisible();
  await expect(primary).toHaveAttribute('aria-keyshortcuts', 'Alt+=');
  await expect(primary).toHaveAttribute('title', '自动求和（Alt+=）');
  await expect(disclosure).toHaveAttribute('aria-haspopup', 'menu');
  await expect(toolbar).toHaveAttribute(
    'data-density',
    testInfo.project.name === 'compact-768' ? 'compact-low' : 'comfortable',
  );

  await grid.focus();
  await page.keyboard.press('Control+Home');
  for (let row = 0; row < 7; row += 1) {
    await page.keyboard.press('ArrowDown');
  }
  await page.keyboard.press('ArrowRight');
  await page.keyboard.press('ArrowRight');
  await expect(nameBox).toHaveText('C8');
  await expect(formulaBar).toHaveText('');

  await bold.click();
  await expect(grid).toBeFocused();
  await expect(bold).toHaveAttribute('aria-pressed', 'true');
  await page.keyboard.press('Alt+Equal');
  await expect(nameBox).toHaveText('C8');
  await expect(formulaBar).toHaveText('=SUM(C4:C7)');
  await expect(bold).toHaveAttribute('aria-pressed', 'true');
  await expect(grid).toBeFocused();

  await page.keyboard.press('Control+z');
  await expect(formulaBar).toHaveText('');
  await expect(bold).toHaveAttribute('aria-pressed', 'true');
  await expect(grid).toBeFocused();

  await disclosure.click();
  const menu = page.getByRole('menu', { name: '自动计算选项' });
  const sum = menu.getByRole('menuitem', { name: '自动求和' });
  const average = menu.getByRole('menuitem', { name: '平均值' });
  await expect(menu.getByRole('menuitem')).toHaveCount(5);
  await expect(sum).toHaveAttribute('aria-keyshortcuts', 'Alt+=');
  await expect(sum).toBeFocused();
  await menu.press('ArrowDown');
  await expect(average).toBeFocused();

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
    path: testInfo.outputPath('spreadsheet-auto-sum-menu.png'),
    animations: 'disabled',
  });

  await page.keyboard.press('Enter');
  await expect(menu).toHaveCount(0);
  await expect(formulaBar).toHaveText('=AVERAGE(C4:C7)');
  await expect(bold).toHaveAttribute('aria-pressed', 'true');
  await expect(grid).toBeFocused();
  expect(browserErrors).toEqual([]);
});
