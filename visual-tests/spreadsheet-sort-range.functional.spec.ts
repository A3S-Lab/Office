import { expect, test } from '@playwright/test';
import { openSpreadsheetFixture } from './visual-test-support';

test('Spreadsheet expands a partial WPS sort range without breaking rows', async ({
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
  for (let index = 0; index < 3; index += 1) {
    await page.keyboard.press('ArrowDown');
  }
  for (let index = 0; index < 5; index += 1) {
    await page.keyboard.press('ArrowRight');
  }
  for (let index = 0; index < 3; index += 1) {
    await page.keyboard.press('Shift+ArrowDown');
  }
  await expect(nameBox).toHaveText('F4:F7');

  await ribbon.getByRole('tab', { name: '数据' }).click();
  await customSort.click();
  let rangeDialog = page.getByRole('dialog', { name: '排序提醒' });
  await expect(rangeDialog).toContainText('A3:G7');
  await expect(rangeDialog).toContainText('F4:F7');
  await expect(
    rangeDialog.getByRole('radio', { name: /扩展选定区域/ }),
  ).toBeChecked();
  await rangeDialog.getByRole('radio', { name: /以当前选定区域排序/ }).check();
  await rangeDialog.getByRole('button', { name: '排序' }).click();

  let sortDialog = page.getByRole('dialog', { name: '自定义排序' });
  await expect(sortDialog).toContainText('执行看板!F4:F7');
  await expect(
    sortDialog.getByRole('checkbox', { name: '数据包含标题' }),
  ).not.toBeChecked();
  await sortDialog.getByRole('button', { name: '取消' }).click();
  await expect(customSort).toBeFocused();

  await customSort.click();
  rangeDialog = page.getByRole('dialog', { name: '排序提醒' });
  await expect(
    rangeDialog.getByRole('radio', { name: /扩展选定区域/ }),
  ).toBeChecked();
  await rangeDialog.screenshot({
    path: testInfo.outputPath('spreadsheet-sort-range-warning.png'),
    animations: 'disabled',
  });
  await rangeDialog.getByRole('button', { name: '排序' }).click();

  sortDialog = page.getByRole('dialog', { name: '自定义排序' });
  await expect(sortDialog).toContainText('执行看板!A3:G7');
  await expect(
    sortDialog.getByRole('checkbox', { name: '数据包含标题' }),
  ).toBeChecked();
  await expect(
    sortDialog.getByRole('combobox', { name: '排序条件 1 列' }),
  ).toHaveValue('5');
  await sortDialog
    .getByRole('combobox', { name: '排序条件 1 次序' })
    .selectOption('descending');
  await sortDialog.getByRole('button', { name: '确定' }).click();
  await expect(customSort).toBeFocused();

  await grid.focus();
  await page.keyboard.press('Control+Home');
  for (let index = 0; index < 3; index += 1) {
    await page.keyboard.press('ArrowDown');
  }
  for (let index = 0; index < 5; index += 1) {
    await page.keyboard.press('ArrowRight');
  }
  await expect(nameBox).toHaveText('F4');
  await expect(formulaBar).toHaveText('=AVERAGE(C4:E4)');

  await page.keyboard.press('Control+z');
  await expect(formulaBar).toHaveText('=SUM(C4:E4)/3');
  expect(browserErrors).toEqual([]);
});
