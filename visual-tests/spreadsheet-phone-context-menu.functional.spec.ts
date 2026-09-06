import { expect, type Page, test } from '@playwright/test';

test('Spreadsheet phone context menu restores the grid focus', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });
  page.on('pageerror', (error) => browserErrors.push(error.message));

  await page.setViewportSize({ width: 390, height: 700 });
  await openSpreadsheetFixture(page);

  const grid = page.locator('.fortune-sheet-overlay');
  await grid.focus();
  await expect(grid).toBeFocused();
  await grid.click({ button: 'right' });

  const menu = page.locator('.workspace-context-menu');
  await expect(menu).toBeVisible();
  await expect(menu.getByRole('menuitem').first()).toBeFocused();
  await page.screenshot({
    path: testInfo.outputPath('spreadsheet-phone-context-menu.png'),
    animations: 'disabled',
  });

  await page.keyboard.press('Escape');
  await expect(menu).toBeHidden();
  await expect(grid).toBeFocused();
  expect(browserErrors).toEqual([]);
});

async function openSpreadsheetFixture(page: Page): Promise<void> {
  await page.goto('/playground/');
  await page
    .getByRole('button', {
      name: '季度执行计划 XLSX · 本次会话',
    })
    .click();
  await page.locator('.work-spreadsheet-canvas > .fortune-container').waitFor();
}
