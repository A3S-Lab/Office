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

async function openSpreadsheetFixture(page: Page): Promise<void> {
  await page.goto('/');
  await page
    .getByRole('button', {
      name: '季度执行计划 XLSX · 本次会话',
    })
    .click();
  await page.locator('.work-spreadsheet-canvas > .fortune-container').waitFor();
}
