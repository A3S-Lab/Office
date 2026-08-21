import { expect, test } from '@playwright/test';
import { openSpreadsheetFixture } from './visual-test-support';

test('Spreadsheet preserves advanced underline styles across ribbon, dialog, shortcut, and undo', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await openSpreadsheetFixture(page);

  const grid = page.locator('.fortune-sheet-overlay');
  const ribbon = page.locator('.work-spreadsheet-ribbon');
  const fontGroup = ribbon.getByRole('region', { name: '字体' });
  const underline = fontGroup.getByRole('button', {
    name: '下划线',
    exact: true,
  });
  const moreUnderline = fontGroup.getByRole('button', {
    name: '更多下划线',
  });

  await grid.focus();
  await page.keyboard.press('Shift+F11');
  await page.keyboard.type('A3S underline');
  await page.keyboard.press('Enter');
  await page.keyboard.press('ArrowUp');

  await expect(underline).toHaveAttribute('aria-pressed', 'false');
  await expect(underline).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+U Meta+U',
  );
  await moreUnderline.click();
  const menu = page.getByRole('menu', { name: '下划线样式' });
  await expect(menu).toBeVisible();
  await expect(menu.getByRole('menuitemradio')).toHaveCount(5);
  await expect(
    menu.getByRole('menuitemradio', { name: '无下划线' }),
  ).toBeFocused();
  await expect(
    menu.getByRole('menuitemradio', { name: '无下划线' }),
  ).toHaveAttribute('aria-checked', 'true');
  await page.screenshot({
    path: testInfo.outputPath('spreadsheet-underline-menu.png'),
    animations: 'disabled',
  });
  await menu.getByRole('menuitemradio', { name: '双会计用下划线' }).click();

  await expect(grid).toBeFocused();
  await expect(underline).toHaveAttribute('aria-pressed', 'true');
  await expect(underline).toHaveAttribute(
    'title',
    '下划线（双会计用下划线；Cmd/Ctrl+U）',
  );

  await page.keyboard.press('Control+1');
  const dialog = page.getByRole('dialog', { name: '设置单元格格式' });
  await expect(dialog).toBeVisible();
  await dialog.getByRole('tab', { name: '字体' }).click();
  await dialog.getByRole('combobox', { name: '下划线样式' }).click();
  const underlineList = page.getByRole('listbox', { name: '下划线样式' });
  await expect(
    underlineList.getByRole('option', { name: '双会计用下划线' }),
  ).toHaveAttribute('aria-selected', 'true');
  await underlineList.getByRole('option', { name: '单会计用下划线' }).click();
  await dialog.getByRole('button', { name: '应用' }).click();

  await expect(grid).toBeFocused();
  await expect(underline).toHaveAttribute(
    'title',
    '下划线（单会计用下划线；Cmd/Ctrl+U）',
  );
  await page.screenshot({
    path: testInfo.outputPath('spreadsheet-accounting-underline.png'),
    animations: 'disabled',
  });

  await page.keyboard.press('Control+u');
  await expect(underline).toHaveAttribute('aria-pressed', 'false');
  await expect(underline).toHaveAttribute(
    'title',
    '下划线（无下划线；Cmd/Ctrl+U）',
  );
  await page.keyboard.press('Control+z');
  await expect(underline).toHaveAttribute(
    'title',
    '下划线（单会计用下划线；Cmd/Ctrl+U）',
  );
  expect(browserErrors).toEqual([]);
});
