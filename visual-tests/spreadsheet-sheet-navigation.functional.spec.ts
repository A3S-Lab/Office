import { expect, test } from '@playwright/test';
import { openSpreadsheetFixture } from './visual-test-support';

test('Spreadsheet switches worksheets with WPS Ctrl+PageUp/PageDown', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });

  await openSpreadsheetFixture(page);

  const sheetBar = page.locator('.work-spreadsheet-sheet-bar');
  await expect(
    sheetBar.getByRole('button', { name: '下一个工作表' }),
  ).toHaveAttribute('aria-keyshortcuts', 'Control+PageDown Meta+PageDown');
  await expect(sheetBar.getByRole('tab', { name: '执行看板' })).toHaveAttribute(
    'aria-selected',
    'true',
  );

  await sheetBar.getByRole('button', { name: '新建工作表' }).click();
  await expect(sheetBar.getByRole('tab', { name: '工作表 2' })).toHaveAttribute(
    'aria-selected',
    'true',
  );

  const grid = page.locator('.fortune-sheet-overlay');
  await grid.focus();
  await page.keyboard.press('Control+PageUp');
  await expect(sheetBar.getByRole('tab', { name: '执行看板' })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  await expect(grid).toBeFocused();

  await page.keyboard.press('Control+PageDown');
  await expect(sheetBar.getByRole('tab', { name: '工作表 2' })).toHaveAttribute(
    'aria-selected',
    'true',
  );

  await page.screenshot({
    path: testInfo.outputPath(
      `spreadsheet-sheet-navigation-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  expect(browserErrors).toEqual([]);
});
