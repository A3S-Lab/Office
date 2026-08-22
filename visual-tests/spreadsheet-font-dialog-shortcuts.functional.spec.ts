import { expect, test } from '@playwright/test';
import { openSpreadsheetFixture } from './visual-test-support';

test('Spreadsheet opens focused font controls through Traditional Office shortcuts', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });
  await openSpreadsheetFixture(page);

  const grid = page.locator('.fortune-sheet-overlay');
  const formulaBar = page.locator('.fortune-fx-input');
  await grid.focus();
  await page.keyboard.press('Control+Shift+F');

  let dialog = page.getByRole('dialog', { name: '设置单元格格式' });
  await expect(dialog).toBeVisible();
  await expect(dialog.getByRole('tab', { name: '字体' })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  const family = dialog.getByRole('combobox', { name: '单元格字体' });
  await expect(family).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+F Meta+Shift+F',
  );
  await expect(family).toBeFocused();
  await page.screenshot({
    path: testInfo.outputPath('spreadsheet-font-family-shortcut.png'),
    animations: 'disabled',
  });

  await page.keyboard.press('Escape');
  await expect(dialog).toBeHidden();
  await expect(grid).toBeFocused();
  await page.keyboard.press('Control+Shift+P');

  dialog = page.getByRole('dialog', { name: '设置单元格格式' });
  await expect(dialog).toBeVisible();
  const size = dialog.getByRole('combobox', { name: '单元格字号' });
  await expect(size).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+P Meta+Shift+P',
  );
  await expect(size).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(grid).toBeFocused();

  await formulaBar.click();
  await page.keyboard.press('Control+Shift+F');
  await expect(
    page.getByRole('dialog', { name: '设置单元格格式' }),
  ).toHaveCount(0);
  expect(browserErrors).toEqual([]);
});
