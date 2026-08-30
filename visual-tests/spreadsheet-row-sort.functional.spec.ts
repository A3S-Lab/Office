import { expect, type Locator, type Page, test } from '@playwright/test';

test('Spreadsheet applies and undoes WPS left-to-right sorting', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await page.goto('/playground/?e2e=spreadsheet-row-sort');

  const grid = page.locator('.fortune-sheet-overlay');
  const nameBox = page.locator('.fortune-name-box');
  const formulaBar = page.locator('.fortune-fx-input');
  const ribbon = page.locator('.work-spreadsheet-ribbon');
  const customSort = ribbon.getByRole('button', { name: '自定义排序' });
  await expect(grid).toBeVisible();
  await expect(nameBox).toHaveText('A1:D3');

  await ribbon.getByRole('tab', { name: '数据' }).click();
  await customSort.click();
  const dialog = page.getByRole('dialog', { name: '自定义排序' });
  await expect(dialog).toContainText('Horizontal plan!A1:D3');
  await expect(
    dialog.getByRole('checkbox', { name: '数据包含标题' }),
  ).not.toBeChecked();

  await dialog.getByRole('button', { name: '选项…' }).click();
  const options = page.getByRole('dialog', { name: '排序选项' });
  await expect(options.getByRole('radio', { name: /按列排序/ })).toBeChecked();
  await options.getByRole('radio', { name: /按行排序/ }).click();
  await expectDialogInsideViewport(page, options);
  await options.screenshot({
    path: testInfo.outputPath(
      `spreadsheet-row-sort-options-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  await options.getByRole('button', { name: '确定' }).click();
  await expect(options).toHaveCount(0);

  const header = dialog.getByRole('checkbox', { name: '数据包含标题' });
  await expect(header).not.toBeChecked();
  await expect(header).toBeDisabled();
  await expect(
    dialog.getByRole('combobox', { name: '排序条件 1 行' }),
  ).toHaveValue('0');
  await dialog.getByRole('button', { name: '添加条件' }).click();
  await expect(
    dialog.getByRole('combobox', { name: '排序条件 2 行' }),
  ).toHaveValue('1');
  await dialog
    .getByRole('combobox', { name: '排序条件 2 次序' })
    .selectOption('descending');
  await expectDialogInsideViewport(page, dialog);
  await dialog.screenshot({
    path: testInfo.outputPath(
      `spreadsheet-row-sort-${testInfo.project.name}.png`,
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
  await expect(formulaBar).toHaveText('Beta');
  await page.keyboard.press('ArrowDown');
  await expect(nameBox).toHaveText('A3');
  await expect(formulaBar).toHaveText('=A2&"-B"');

  await page.keyboard.press('Control+z');
  await expect(formulaBar).toHaveText('=A2&"-G"');
  await expect(grid).toBeFocused();
  await page.keyboard.press('ArrowUp');
  await expect(nameBox).toHaveText('A2');
  await expect(formulaBar).toHaveText('Gamma');
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
