import { expect, test } from '@playwright/test';

test('Spreadsheet renders native rich text and formats every run through one controlled update', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await page.goto('/?e2e=spreadsheet-rich-text');

  const grid = page.locator('.fortune-sheet-overlay');
  const status = page.getByTestId('spreadsheet-rich-text-status');
  await expect(grid).toBeVisible();
  await expect(grid).toBeFocused();
  await expect(status).toHaveAttribute('data-revision', '1');
  await expect(status).toHaveAttribute('data-run-count', '3');
  await expect(status).toHaveAttribute('data-run-bold', '1,0,0');
  await expect(status).toHaveAttribute('data-run-text', 'Native rich text');

  await page.keyboard.press('Control+Shift+F');
  const dialog = page.getByRole('dialog', { name: '设置单元格格式' });
  await expect(dialog).toBeVisible();
  await dialog.getByRole('checkbox', { name: '加粗' }).click();
  await dialog.getByRole('button', { name: '应用' }).click();

  await expect(status).toHaveAttribute('data-revision', '2');
  await expect(status).toHaveAttribute('data-run-count', '3');
  await expect(status).toHaveAttribute('data-run-bold', '1,1,1');
  await expect(status).toHaveAttribute('data-run-text', 'Native rich text');
  await expect(grid).toBeFocused();
  await page.screenshot({
    path: testInfo.outputPath('spreadsheet-rich-text.png'),
    animations: 'disabled',
  });
  expect(browserErrors).toEqual([]);
});
