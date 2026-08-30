import { expect, test } from '@playwright/test';
import { openSpreadsheetFixture } from './visual-test-support';

test('Spreadsheet applies and undoes a formula-safe multi-key sort', async ({
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
  await ribbon.getByRole('button', { name: '自定义排序' }).click();
  const dialog = page.getByRole('dialog', { name: '自定义排序' });
  await expect(dialog).toContainText('执行看板!A3:G7');
  await expect(
    dialog.getByRole('checkbox', { name: '数据包含标题' }),
  ).toBeChecked();
  await dialog
    .getByRole('combobox', { name: '排序条件 1 列' })
    .selectOption('5');
  await dialog
    .getByRole('combobox', { name: '排序条件 1 次序' })
    .selectOption('descending');
  await dialog.getByRole('button', { name: '添加条件' }).click();
  await dialog
    .getByRole('combobox', { name: '排序条件 2 列' })
    .selectOption('1');
  await dialog.screenshot({
    path: testInfo.outputPath('spreadsheet-custom-sort-dialog.png'),
    animations: 'disabled',
  });
  await dialog.getByRole('button', { name: '确定' }).click();
  await expect(dialog).toHaveCount(0);
  await expect(
    ribbon.getByRole('button', { name: '自定义排序' }),
  ).toBeFocused();
  await grid.focus();

  await page.keyboard.press('Control+Home');
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('ArrowDown');
  await expect(nameBox).toHaveText('A4');
  await expect(formulaBar).toHaveText('团队能力建设');
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
