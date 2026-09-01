import { expect, test } from '@playwright/test';

test('Spreadsheet totals-row controls stay discoverable and usable', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });
  page.on('pageerror', (error) => browserErrors.push(error.message));

  await page.goto('/playground/');
  await page
    .locator("button[data-template-id='structured-references']")
    .click();
  await page.locator('.work-spreadsheet-canvas > .fortune-container').waitFor();
  await expect(page.getByRole('textbox', { name: '文件名' })).toHaveValue(
    '结构化引用示例',
  );

  const grid = page.locator('.fortune-sheet-overlay');
  await grid.focus();
  await page.keyboard.press('Control+Home');
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowDown');

  const ribbon = page.locator('.work-spreadsheet-ribbon');
  await expect(ribbon.getByRole('tab', { name: '表格设计' })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  await ribbon.getByRole('button', { name: '汇总行' }).click();

  const dialog = page.getByRole('dialog', { name: '表格汇总行设置' });
  await expect(dialog).toBeVisible();
  await expect(
    dialog.getByRole('checkbox', { name: '启用汇总行' }),
  ).toBeChecked();
  await expect(
    dialog.getByRole('combobox', { name: 'Units 汇总函数' }),
  ).toHaveValue('sum');
  await dialog.screenshot({
    path: testInfo.outputPath('spreadsheet-table-totals-dialog.png'),
    animations: 'disabled',
  });
  await dialog.getByRole('button', { name: '完成' }).click();
  await expect(ribbon.getByRole('button', { name: '汇总行' })).toHaveAttribute(
    'aria-pressed',
    'true',
  );
  expect(browserErrors).toEqual([]);
});
