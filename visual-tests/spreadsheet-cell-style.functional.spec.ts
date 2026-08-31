import { expect, test } from '@playwright/test';
import { openSpreadsheetFixture } from './visual-test-support';

test('Spreadsheet applies and undoes WPS cell styles from the grouped gallery', async ({
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
  const ribbon = page.locator('.work-spreadsheet-ribbon');
  const trigger = ribbon.getByRole('button', { name: '单元格样式' });

  await grid.focus();
  await page.keyboard.press('Shift+F11');
  await expect(page.getByRole('tab', { name: '工作表 2' })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  await expect(grid).toBeFocused();
  await page.keyboard.type('Ready');
  await page.keyboard.press('Enter');
  await page.keyboard.press('ArrowUp');
  await page.keyboard.press('Shift+ArrowRight');
  await page.keyboard.press('Shift+ArrowDown');
  await expect(nameBox).toHaveText('A1:B2');

  await expect(trigger).toHaveAttribute('aria-haspopup', 'menu');
  await expect(trigger).toHaveAttribute('title', '单元格样式（当前：常规）');
  await trigger.click();
  const menu = page.getByRole('menu', { name: '单元格样式库' });
  await expect(menu).toBeVisible();
  await expect(menu.getByRole('group', { name: '常用' })).toBeVisible();
  await expect(menu.getByRole('group', { name: '数据和模型' })).toBeVisible();
  await expect(menu.getByRole('group', { name: '标题和汇总' })).toBeVisible();
  await expect(menu.getByRole('menuitemradio')).toHaveCount(17);
  const normal = menu.getByRole('menuitemradio', {
    name: '应用单元格样式：常规',
  });
  const good = menu.getByRole('menuitemradio', {
    name: '应用单元格样式：好',
  });
  const total = menu.getByRole('menuitemradio', {
    name: '应用单元格样式：总计',
  });
  const neutral = menu.getByRole('menuitemradio', {
    name: '应用单元格样式：适中',
  });
  await expect(normal).toBeFocused();
  await expect(normal).toHaveAttribute('aria-checked', 'true');
  await expect(good.locator('.work-spreadsheet-cell-style-preview')).toHaveCSS(
    'background-color',
    'rgb(198, 239, 206)',
  );

  const bounds = await menu.boundingBox();
  const viewport = page.viewportSize();
  expect(bounds).not.toBeNull();
  expect(viewport).not.toBeNull();
  expect((bounds?.x ?? 0) + (bounds?.width ?? 0)).toBeLessThanOrEqual(
    viewport?.width ?? 0,
  );
  expect((bounds?.y ?? 0) + (bounds?.height ?? 0)).toBeLessThanOrEqual(
    viewport?.height ?? 0,
  );
  await page.screenshot({
    path: testInfo.outputPath('spreadsheet-cell-style-gallery.png'),
    animations: 'disabled',
  });

  await menu.press('ArrowRight');
  await expect(good).toBeFocused();
  await menu.press('ArrowDown');
  await expect(neutral).toBeFocused();
  await menu.press('End');
  await expect(total).toBeFocused();
  await menu.press('Home');
  await expect(normal).toBeFocused();
  await good.click();
  await expect(menu).toHaveCount(0);
  await expect(trigger).toBeFocused();
  await expect(trigger).toHaveAttribute('title', '单元格样式（当前：好）');

  await grid.focus();
  await page.keyboard.press('Control+Home');
  await page.keyboard.press('ArrowRight');
  await page.keyboard.press('ArrowDown');
  await expect(nameBox).toHaveText('B2');
  await expect(trigger).toHaveAttribute('title', '单元格样式（当前：好）');
  await page.screenshot({
    path: testInfo.outputPath('spreadsheet-cell-style-applied.png'),
    animations: 'disabled',
  });

  await ribbon.getByRole('button', { name: '撤销' }).click();
  await expect(trigger).toHaveAttribute('title', '单元格样式（当前：常规）');

  await trigger.click();
  await page
    .getByRole('menu', { name: '单元格样式库' })
    .getByRole('menuitemradio', { name: '应用单元格样式：总计' })
    .click();
  await expect(trigger).toHaveAttribute('title', '单元格样式（当前：总计）');
  await ribbon.getByRole('button', { name: '撤销' }).click();
  await expect(trigger).toHaveAttribute('title', '单元格样式（当前：常规）');
  expect(browserErrors).toEqual([]);
});
