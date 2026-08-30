import { expect, type Locator, type Page, test } from '@playwright/test';

test('Spreadsheet applies and undoes effective appearance sorting', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await page.goto('/playground/?e2e=spreadsheet-appearance-sort');

  const grid = page.locator('.fortune-sheet-overlay');
  const nameBox = page.locator('.fortune-name-box');
  const formulaBar = page.locator('.fortune-fx-input');
  const ribbon = page.locator('.work-spreadsheet-ribbon');
  const customSort = ribbon.getByRole('button', { name: '自定义排序' });
  await expect(grid).toBeVisible();
  await expect(nameBox).toHaveText('A1:D6');

  await ribbon.getByRole('tab', { name: '数据' }).click();
  await customSort.click();
  const dialog = page.getByRole('dialog', { name: '自定义排序' });
  await expect(dialog).toContainText('Visual priority!A1:D6');
  await expect(
    dialog.getByRole('checkbox', { name: '数据包含标题' }),
  ).toBeChecked();
  await expect(
    dialog.getByRole('combobox', { name: '排序条件 1 列' }),
  ).toHaveValue('1');

  await dialog
    .getByRole('combobox', { name: '排序条件 1 排序依据' })
    .selectOption('cell-color');
  await dialog
    .getByRole('combobox', { name: '排序条件 1 目标外观' })
    .selectOption('cell-color:#fce8e6');
  await expect(dialog).toContainText('单元格颜色 #FCE8E6，置于顶端');

  await dialog.getByRole('button', { name: '添加条件' }).click();
  await dialog
    .getByRole('combobox', { name: '排序条件 2 列' })
    .selectOption('1');
  await dialog
    .getByRole('combobox', { name: '排序条件 2 排序依据' })
    .selectOption('cell-color');
  await dialog
    .getByRole('combobox', { name: '排序条件 2 目标外观' })
    .selectOption('cell-color:#fff2cc');
  await expect(dialog).toContainText('单元格颜色 #FFF2CC，置于顶端');

  await dialog.getByRole('button', { name: '添加条件' }).click();
  await dialog
    .getByRole('combobox', { name: '排序条件 3 排序依据' })
    .selectOption('font-color');
  await dialog
    .getByRole('combobox', { name: '排序条件 3 目标外观' })
    .selectOption('font-color:#d84b4f');
  await expect(dialog).toContainText('字体颜色 #D84B4F，置于顶端');

  await dialog.getByRole('button', { name: '添加条件' }).click();
  await expect(
    dialog.getByRole('combobox', { name: '排序条件 4 列' }),
  ).toHaveValue('2');
  await dialog
    .getByRole('combobox', { name: '排序条件 4 排序依据' })
    .selectOption('icon');
  await dialog
    .getByRole('combobox', { name: '排序条件 4 目标外观' })
    .selectOption('icon:3TrafficLights1:2');
  await expect(dialog).toContainText('三色交通灯（实心） 3/3，置于顶端');
  await expectDialogInsideViewport(page, dialog);
  await dialog.screenshot({
    path: testInfo.outputPath(
      `spreadsheet-appearance-sort-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });

  await dialog.getByRole('button', { name: '确定' }).click();
  await expect(dialog).toHaveCount(0);
  await expect(customSort).toBeFocused();

  await grid.focus();
  await page.keyboard.press('Control+Home');
  await page.keyboard.press('ArrowDown');
  await expect(nameBox).toHaveText('A2');
  await expect(formulaBar).toHaveText('Zulu');
  await page.keyboard.press('ArrowDown');
  await expect(formulaBar).toHaveText('Beta');
  await page.keyboard.press('ArrowDown');
  await expect(formulaBar).toHaveText('Gamma');
  await page.keyboard.press('Control+Home');
  await page.keyboard.press('ArrowDown');
  for (let index = 0; index < 3; index += 1) {
    await page.keyboard.press('ArrowRight');
  }
  await expect(nameBox).toHaveText('D2');
  await expect(formulaBar).toHaveText('=C2+90');

  await page.keyboard.press('Control+z');
  await expect(formulaBar).toHaveText('=C2*2');
  await expect(grid).toBeFocused();
  for (let index = 0; index < 3; index += 1) {
    await page.keyboard.press('ArrowLeft');
  }
  await expect(nameBox).toHaveText('A2');
  await expect(formulaBar).toHaveText('Alpha');
  expect(browserErrors).toEqual([]);
});

async function expectDialogInsideViewport(
  page: Page,
  dialog: Locator,
): Promise<void> {
  const bounds = await dialog.boundingBox();
  const viewport = page.viewportSize();
  expect(bounds).not.toBeNull();
  expect(viewport).not.toBeNull();
  expect(bounds?.x ?? -1).toBeGreaterThanOrEqual(0);
  expect(bounds?.y ?? -1).toBeGreaterThanOrEqual(0);
  expect((bounds?.x ?? 0) + (bounds?.width ?? 0)).toBeLessThanOrEqual(
    viewport?.width ?? 0,
  );
  expect((bounds?.y ?? 0) + (bounds?.height ?? 0)).toBeLessThanOrEqual(
    viewport?.height ?? 0,
  );
}
