import { expect, test } from '@playwright/test';
import { openSpreadsheetFixture } from './visual-test-support';

test('Spreadsheet creates, designs, converts, and undoes a native table', async ({
  page,
  context,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await openSpreadsheetFixture(page);

  const grid = page.locator('.fortune-sheet-overlay');
  const nameBox = page.locator('.fortune-name-box');
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

  const pageCount = context.pages().length;
  await page.keyboard.press(
    process.platform === 'darwin' ? 'Meta+t' : 'Control+t',
  );
  await expect(context.pages()).toHaveLength(pageCount);
  const dialog = page.getByRole('dialog', { name: '创建表格' });
  await expect(dialog).toBeVisible();
  const range = dialog.getByRole('textbox', { name: '表格区域' });
  await expect(range).toHaveValue('A3:G7');
  await expect(
    dialog.getByRole('checkbox', { name: '表包含标题' }),
  ).toBeChecked();
  await range.fill('A3:A3');
  await expect(dialog.getByRole('button', { name: '确定' })).toBeDisabled();
  await range.fill('A3:G7');
  await expect(dialog.getByRole('button', { name: '确定' })).toBeEnabled();
  await expectInsideViewport(page, dialog);
  await dialog.getByRole('button', { name: '确定' }).click();
  await expect(dialog).toHaveCount(0);

  const tableDesign = ribbon.getByRole('tab', { name: '表格设计' });
  await expect(tableDesign).toHaveAttribute('aria-selected', 'true');
  const undo = ribbon.getByRole('button', { name: '撤销' });
  const redo = ribbon.getByRole('button', { name: '重做' });
  await expect(undo).toBeEnabled();
  await undo.click();
  await expect(tableDesign).toHaveCount(0);
  await redo.click();
  await expect(tableDesign).toHaveAttribute('aria-selected', 'true');

  const styleTrigger = ribbon.getByRole('button', { name: '表格样式' });
  await styleTrigger.click();
  const gallery = page.getByRole('menu', { name: '表格样式库' });
  await expect(gallery).toBeVisible();
  await expect(gallery.getByRole('menuitemradio')).toHaveCount(60);
  await expect(gallery.getByRole('group', { name: '浅色' })).toBeVisible();
  await expect(gallery.getByRole('group', { name: '中等' })).toBeVisible();
  await expect(gallery.getByRole('group', { name: '深色' })).toBeVisible();
  const light1 = gallery.getByRole('menuitemradio', {
    name: '应用表格样式：浅色 1',
    exact: true,
  });
  const light2 = gallery.getByRole('menuitemradio', {
    name: '应用表格样式：浅色 2',
    exact: true,
  });
  const light9 = gallery.getByRole('menuitemradio', {
    name: '应用表格样式：浅色 9',
    exact: true,
  });
  await expect(light1).toBeFocused();
  await gallery.press('ArrowRight');
  await expect(light2).toBeFocused();
  await gallery.press('ArrowDown');
  await expect(light9).toBeFocused();
  await expectInsideViewport(page, gallery);
  await page.screenshot({
    path: testInfo.outputPath('spreadsheet-table-style-gallery.png'),
    animations: 'disabled',
  });
  await gallery
    .getByRole('menuitemradio', { name: '应用表格样式：深色 11' })
    .click();
  await expect(gallery).toHaveCount(0);
  await expect(styleTrigger).toBeFocused();

  const tableName = ribbon.getByRole('textbox', { name: '表格名称' });
  await tableName.fill('Quarterly_2026');
  await tableName.press('Enter');
  await expect(tableName).toHaveValue('Quarterly_2026');
  const firstColumn = ribbon.getByRole('button', { name: '首列' });
  await firstColumn.click();
  await expect(firstColumn).toHaveAttribute('aria-pressed', 'true');
  const rowStripes = ribbon.getByRole('button', { name: '行条纹' });
  await rowStripes.click();
  await expect(rowStripes).toHaveAttribute('aria-pressed', 'false');

  await ribbon.getByRole('button', { name: '转换为区域' }).click();
  await expect(tableDesign).toHaveCount(0);
  await expect(ribbon.getByRole('tab', { name: '开始' })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  await undo.click();
  await expect(tableDesign).toHaveAttribute('aria-selected', 'true');
  expect(browserErrors).toEqual([]);
});

async function expectInsideViewport(
  page: import('@playwright/test').Page,
  element: import('@playwright/test').Locator,
) {
  const bounds = await element.boundingBox();
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
