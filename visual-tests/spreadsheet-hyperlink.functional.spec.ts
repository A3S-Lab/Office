import { expect, test } from '@playwright/test';

test('Spreadsheet hyperlinks stay atomic and accessible at every layout', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await page.goto('/playground/?e2e=spreadsheet-hyperlink');
  await page.locator('.work-spreadsheet-canvas > .fortune-container').waitFor();

  const grid = page.locator('.fortune-sheet-overlay');
  const formulaBar = page.locator('.fortune-fx-input');
  const ribbon = page.locator('.work-spreadsheet-ribbon');
  const undo = ribbon.getByRole('button', { name: '撤销' });
  await grid.focus();
  await page.keyboard.press('Control+Home');
  await expect(page.locator('.fortune-name-box')).toHaveText('A1');
  await expect(formulaBar).toHaveText('A3S Office');

  await ribbon.getByRole('tab', { name: '插入' }).click();
  const launcher = ribbon.getByRole('button', { name: '超链接' });
  await expect(launcher).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+K Meta+K',
  );
  await launcher.focus();
  await launcher.click();
  let dialog = page.getByRole('dialog', { name: '插入超链接' });
  await expect(dialog).toBeVisible();
  await expectDialogInsideViewport(page, dialog);
  await page.keyboard.press('Escape');
  await expect(dialog).toHaveCount(0);
  await expect(launcher).toBeFocused();

  await grid.focus();
  await page.keyboard.press('Control+k');
  dialog = page.getByRole('dialog', { name: '插入超链接' });
  await expect(dialog).toBeVisible();
  await expect(dialog.getByRole('radio')).toHaveCount(3);
  await expect(dialog.getByRole('radio', { name: '网页' })).toBeChecked();
  await dialog.getByRole('radio', { name: '单元格区域' }).click();
  const cellRange = dialog.getByRole('textbox', { name: '单元格或区域' });
  await cellRange.fill("'Hidden Archive'!A1");
  await expect(dialog.getByRole('alert')).toHaveText('不能链接到隐藏工作表。');

  await dialog.getByRole('radio', { name: '工作表' }).click();
  const sheet = dialog.getByRole('combobox', { name: '工作表' });
  await expect(sheet).toHaveValue('Archive 2025');
  await expect(sheet.getByRole('option')).toHaveCount(2);
  await expect(
    sheet.getByRole('option', { name: 'Hidden Archive' }),
  ).toHaveCount(0);

  await dialog.getByRole('radio', { name: '网页' }).click();
  const address = dialog.getByRole('textbox', { name: '地址' });
  await address.fill('javascript:alert(1)');
  await expect(dialog.getByRole('alert')).toHaveText(
    '请输入有效的 HTTP 或 HTTPS 地址。',
  );
  await expect(dialog.getByRole('button', { name: '确定' })).toBeDisabled();
  await address.fill('a3s.dev/office');
  await expect(dialog.getByRole('button', { name: '确定' })).toBeEnabled();
  await page.screenshot({
    path: testInfo.outputPath('spreadsheet-hyperlink-dialog.png'),
    animations: 'disabled',
  });
  await dialog.getByRole('button', { name: '确定' }).click();
  await expect(dialog).toHaveCount(0);
  await expect(grid).toBeFocused();
  await expect(formulaBar).toHaveText('A3S Office');
  await expect(undo).toBeEnabled();

  await page.keyboard.press('Control+k');
  dialog = page.getByRole('dialog', { name: '编辑超链接' });
  await expect(dialog).toBeVisible();
  await expect(dialog.getByRole('textbox', { name: '地址' })).toHaveValue(
    'https://a3s.dev/office',
  );
  await dialog.getByRole('button', { name: '移除超链接' }).click();
  await expect(dialog).toHaveCount(0);
  await expect(grid).toBeFocused();
  await expect(formulaBar).toHaveText('A3S Office');

  await page.keyboard.press('Control+k');
  dialog = page.getByRole('dialog', { name: '插入超链接' });
  await expect(dialog).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(grid).toBeFocused();

  await page.keyboard.press('Control+z');
  await expect(grid).toBeFocused();
  await page.keyboard.press('Control+k');
  dialog = page.getByRole('dialog', { name: '编辑超链接' });
  await expect(dialog).toBeVisible();
  await expect(dialog.getByRole('textbox', { name: '地址' })).toHaveValue(
    'https://a3s.dev/office',
  );
  await page.keyboard.press('Escape');
  await expect(grid).toBeFocused();

  await page.keyboard.press('Control+z');
  await expect(grid).toBeFocused();
  await page.keyboard.press('Control+k');
  dialog = page.getByRole('dialog', { name: '插入超链接' });
  await expect(dialog).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(grid).toBeFocused();
  await expect(undo).toBeDisabled();
  expect(browserErrors).toEqual([]);
});

async function expectDialogInsideViewport(
  page: import('@playwright/test').Page,
  dialog: import('@playwright/test').Locator,
) {
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
