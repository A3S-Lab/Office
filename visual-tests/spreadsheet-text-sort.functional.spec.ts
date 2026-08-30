import { expect, type Locator, type Page, test } from '@playwright/test';

test('Spreadsheet applies and undoes WPS Chinese text sorting', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await page.goto('/playground/?e2e=spreadsheet-text-sort');

  const grid = page.locator('.fortune-sheet-overlay');
  const nameBox = page.locator('.fortune-name-box');
  const formulaBar = page.locator('.fortune-fx-input');
  const ribbon = page.locator('.work-spreadsheet-ribbon');
  const customSort = ribbon.getByRole('button', { name: '自定义排序' });
  await expect(grid).toBeVisible();
  await expect(nameBox).toHaveText('A1:C7');

  await ribbon.getByRole('tab', { name: '数据' }).click();
  await customSort.click();
  const dialog = page.getByRole('dialog', { name: '自定义排序' });
  await expect(dialog).toContainText('Chinese names!A1:C7');
  await expect(
    dialog.getByRole('checkbox', { name: '数据包含标题' }),
  ).toBeChecked();

  await dialog.getByRole('button', { name: '选项…' }).click();
  const options = page.getByRole('dialog', { name: '排序选项' });
  await expect(options.getByRole('radio', { name: '拼音排序' })).toBeChecked();
  await expect(
    options.getByRole('checkbox', { name: '区分大小写' }),
  ).not.toBeChecked();
  await options.getByRole('radio', { name: '笔画排序' }).click();
  await options.getByRole('checkbox', { name: '区分大小写' }).click();
  await expectDialogInsideViewport(page, options);
  await options.screenshot({
    path: testInfo.outputPath(
      `spreadsheet-text-sort-options-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  await options.getByRole('button', { name: '确定' }).click();
  await expect(options).toHaveCount(0);

  await dialog.getByRole('button', { name: '添加条件' }).click();
  await expect(
    dialog.getByRole('combobox', { name: '排序条件 2 列' }),
  ).toHaveValue('1');
  await expectDialogInsideViewport(page, dialog);
  await dialog.screenshot({
    path: testInfo.outputPath(
      `spreadsheet-text-sort-${testInfo.project.name}.png`,
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
  await expect(formulaBar).toHaveText('丁');
  await page.keyboard.press('ArrowDown');
  await expect(formulaBar).toHaveText('王');
  await page.keyboard.press('ArrowRight');
  await expect(nameBox).toHaveText('B3');
  await expect(formulaBar).toHaveText('a');
  await page.keyboard.press('Control+Home');
  await page.keyboard.press('ArrowDown');
  await page.keyboard.press('Control+ArrowRight');
  await expect(nameBox).toHaveText('C2');
  await expect(formulaBar).toHaveText('=B2&"-D"');

  await page.keyboard.press('Control+z');
  await expect(formulaBar).toHaveText('=B2&"-Z"');
  await expect(grid).toBeFocused();
  await page.keyboard.press('ArrowLeft');
  await page.keyboard.press('ArrowLeft');
  await expect(nameBox).toHaveText('A2');
  await expect(formulaBar).toHaveText('赵');
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
