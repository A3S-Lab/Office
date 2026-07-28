import { expect, test, type Page } from '@playwright/test';

test('Spreadsheet closes a ribbon-opened workbook pane with Escape', async ({
  page,
}) => {
  await openSpreadsheetFixture(page);

  await page.getByRole('tab', { name: '插入', exact: true }).click();
  await page.getByRole('button', { name: '插入图表' }).click();
  const chartPane = page.getByRole('region', { name: '图表管理器' });
  await expect(chartPane).toBeVisible();

  await page.keyboard.press('Escape');
  await expect(chartPane).toBeHidden();
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
