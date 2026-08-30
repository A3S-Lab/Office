import { expect, test } from '@playwright/test';
import { openSpreadsheetFixture } from './visual-test-support';

test('Spreadsheet authors, reuses, applies, and undoes a custom-list sort', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await openSpreadsheetFixture(page);

  const grid = page.locator('.fortune-sheet-overlay');
  const nameBox = page.locator('.fortune-name-box');
  const formulaBar = page.locator('.fortune-fx-input');
  const ribbon = page.locator('.work-spreadsheet-ribbon');
  const customSort = ribbon.getByRole('button', { name: '自定义排序' });
  await grid.focus();
  await page.keyboard.press('Control+Home');
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowDown');
  for (let index = 0; index < 6; index += 1) {
    await page.keyboard.press('Shift+ArrowRight');
  }
  for (let index = 0; index < 4; index += 1) {
    await page.keyboard.press('Shift+ArrowDown');
  }
  await expect(nameBox).toHaveText('A3:G7');

  await ribbon.getByRole('tab', { name: '数据' }).click();
  await customSort.click();
  let dialog = page.getByRole('dialog', { name: '自定义排序' });
  await expect(dialog).toContainText('执行看板!A3:G7');
  await dialog
    .getByRole('combobox', { name: '排序条件 1 列' })
    .selectOption('6');
  const order = dialog.getByRole('combobox', { name: '排序条件 1 次序' });
  await order.selectOption({ label: '新建自定义序列…' });
  await dialog
    .getByRole('textbox', { name: '排序条件 1 自定义序列' })
    .fill('有风险\n进行中\n正常\n已完成');
  await dialog.getByRole('button', { name: '使用序列' }).click();
  await expect(order).toHaveValue('custom-list:7');
  await expect(
    order.getByRole('option', { name: '有风险 → 进行中 → 正常 → …' }),
  ).toBeAttached();
  await dialog.getByRole('button', { name: '添加条件' }).click();
  await expect(
    dialog.getByRole('combobox', { name: '排序条件 2 列' }),
  ).toHaveValue('0');
  await dialog.screenshot({
    path: testInfo.outputPath('spreadsheet-custom-list-sort-dialog.png'),
    animations: 'disabled',
  });
  await dialog.getByRole('button', { name: '确定' }).click();
  await expect(dialog).toHaveCount(0);
  await expect(customSort).toBeFocused();

  await customSort.click();
  dialog = page.getByRole('dialog', { name: '自定义排序' });
  await expect(
    dialog
      .getByRole('combobox', { name: '排序条件 1 次序' })
      .getByRole('option', { name: '有风险 → 进行中 → 正常 → …' }),
  ).toBeAttached();
  await dialog.getByRole('button', { name: '取消' }).click();
  await expect(customSort).toBeFocused();

  await grid.focus();
  await page.keyboard.press('Control+Home');
  for (let index = 0; index < 3; index += 1) {
    await page.keyboard.press('ArrowDown');
  }
  await expect(nameBox).toHaveText('A4');
  await expect(formulaBar).toHaveText('新版发布');
  for (let index = 0; index < 5; index += 1) {
    await page.keyboard.press('ArrowRight');
  }
  await expect(nameBox).toHaveText('F4');
  await expect(formulaBar).toHaveText('=AVERAGE(C4:E4)');

  await page.keyboard.press('Control+z');
  await expect(formulaBar).toHaveText('=SUM(C4:E4)/3');
  for (let index = 0; index < 5; index += 1) {
    await page.keyboard.press('ArrowLeft');
  }
  await expect(nameBox).toHaveText('A4');
  await expect(formulaBar).toHaveText('客户洞察报告');
  expect(browserErrors).toEqual([]);
});
