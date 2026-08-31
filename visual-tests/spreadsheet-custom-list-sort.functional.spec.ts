import { expect, test } from '@playwright/test';
import { openSpreadsheetFixture } from './visual-test-support';

test('Spreadsheet persists, applies, and undoes a custom-list sort', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await page.goto('/playground/');
  await page.evaluate(() => {
    localStorage.removeItem('a3s-office.spreadsheet-sort-custom-lists.v1');
  });
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
  await expect(order.locator('optgroup[label="已保存的序列"]')).toBeAttached();
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
  const reopenedOrder = dialog.getByRole('combobox', {
    name: '排序条件 1 次序',
  });
  await expect(
    reopenedOrder.getByRole('option', {
      name: '有风险 → 进行中 → 正常 → …',
    }),
  ).toBeAttached();
  const managerButton = dialog.getByRole('button', {
    name: '管理自定义序列',
  });
  const usesNarrowManagerLayout = testInfo.project.name === 'compact-768';
  if (usesNarrowManagerLayout) {
    await page.setViewportSize({ width: 680, height: 900 });
  }
  await managerButton.click();
  const manager = page.getByRole('dialog', { name: '自定义序列' });
  const list = manager.getByRole('listbox', { name: '自定义序列列表' });
  const entries = manager.getByRole('textbox', { name: '自定义序列项目' });
  await expect(list).toBeFocused();
  await expect(entries).toHaveAttribute('readonly', '');

  await list.selectOption('user:0');
  await entries.fill('紧急\n普通');
  await manager.getByRole('button', { name: '保存更改' }).click();
  await manager.getByRole('button', { name: '新建序列' }).click();
  await entries.fill('待处理\n处理中\n已完成');
  await manager.getByRole('button', { name: '添加序列' }).click();
  await manager.getByRole('button', { name: '新建序列' }).click();
  await entries.fill('北区\n中区\n南区');
  await manager.getByRole('button', { name: '添加序列' }).click();
  await manager.getByRole('button', { name: '上移序列' }).click();
  await manager.getByRole('button', { name: '上移序列' }).click();
  await list.selectOption('user:1');
  await manager.getByRole('button', { name: '删除序列' }).click();
  await expect(list.locator('optgroup[label="用户序列"] option')).toHaveText([
    '北区 → 中区 → 南区',
    '紧急 → 普通',
  ]);
  const managerBounds = await manager.boundingBox();
  const managerViewport = page.viewportSize();
  expect(managerBounds).not.toBeNull();
  expect(managerViewport).not.toBeNull();
  expect(
    (managerBounds?.x ?? 0) + (managerBounds?.width ?? 0),
  ).toBeLessThanOrEqual(managerViewport?.width ?? 0);
  expect(
    (managerBounds?.y ?? 0) + (managerBounds?.height ?? 0),
  ).toBeLessThanOrEqual(managerViewport?.height ?? 0);
  await manager.screenshot({
    path: testInfo.outputPath('spreadsheet-custom-list-manager-dialog.png'),
    animations: 'disabled',
  });
  await manager.getByRole('button', { name: '确定' }).click();
  await expect(manager).toHaveCount(0);
  await expect(managerButton).toBeFocused();
  await expect(
    reopenedOrder.locator('optgroup[label="已保存的序列"] option'),
  ).toHaveText(['北区 → 中区 → 南区', '紧急 → 普通']);
  if (usesNarrowManagerLayout) {
    await page.setViewportSize({ width: 768, height: 800 });
  }
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

  await openSpreadsheetFixture(page);
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
  dialog = page.getByRole('dialog', { name: '自定义排序' });
  const remountedOrder = dialog.getByRole('combobox', {
    name: '排序条件 1 次序',
  });
  await expect(
    remountedOrder.locator('optgroup[label="已保存的序列"]'),
  ).toBeAttached();
  await expect(
    remountedOrder.locator('optgroup[label="已保存的序列"] option'),
  ).toHaveText(['北区 → 中区 → 南区', '紧急 → 普通']);
  await expect(
    remountedOrder.getByRole('option', {
      name: '有风险 → 进行中 → 正常 → …',
    }),
  ).toHaveCount(0);
  await dialog.getByRole('button', { name: '取消' }).click();
  await expect(customSort).toBeFocused();
  expect(browserErrors).toEqual([]);
});
