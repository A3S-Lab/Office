import { expect, type Page, test } from '@playwright/test';

test('Spreadsheet phone task pane contains focus and restores the ribbon invoker', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });
  page.on('pageerror', (error) => browserErrors.push(error.message));

  await page.setViewportSize({ width: 390, height: 700 });
  await openSpreadsheetFixture(page);
  await page.getByRole('tab', { name: '数据', exact: true }).click();

  const trigger = page.getByRole('button', { name: '数据透视表' });
  await trigger.click();

  const pane = page.getByRole('dialog', { name: '数据透视表管理器' });
  const close = pane.getByRole('button', { name: '关闭数据透视表' });
  const create = pane.getByRole('button', { name: '根据当前选区新建' });
  await expect(pane).toHaveAttribute('aria-modal', 'true');
  await expect(page.locator('.work-spreadsheet-ribbon')).toHaveAttribute(
    'inert',
    '',
  );
  await expect(page.locator('.work-spreadsheet-canvas')).toHaveAttribute(
    'inert',
    '',
  );
  await expect(page.locator('.work-spreadsheet-footer')).toHaveAttribute(
    'inert',
    '',
  );
  await expect(close).toBeFocused();

  await close.press('Tab');
  await expect(create).toBeFocused();
  await create.press('Shift+Tab');
  await expect(close).toBeFocused();

  await page.screenshot({
    path: testInfo.outputPath('spreadsheet-phone-task-pane.png'),
    animations: 'disabled',
  });

  await close.press('Escape');
  await expect(pane).toBeHidden();
  await expect(trigger).toBeFocused();
  await expect(page.locator('.work-spreadsheet-ribbon')).not.toHaveAttribute(
    'inert',
    '',
  );
  await expect(page.locator('.work-spreadsheet-canvas')).not.toHaveAttribute(
    'inert',
    '',
  );
  await expect(page.locator('.work-spreadsheet-footer')).not.toHaveAttribute(
    'inert',
    '',
  );
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
