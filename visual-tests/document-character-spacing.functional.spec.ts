import { expect, test } from '@playwright/test';

test('Writer authors native character spacing through the focused font dialog', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  const modifier = process.platform === 'darwin' ? 'Meta' : 'Control';
  page.on('pageerror', (error) => browserErrors.push(error.message));
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });
  await page.goto('/');
  await page
    .getByRole('button', {
      name: '空白文字 从一张干净的 A4 页面开始',
    })
    .click();

  const editor = page.getByRole('textbox', { name: '文档正文' });
  await expect(editor).toHaveAttribute('data-pagination-state', 'ready');
  await editor.fill('Native character spacing');
  await editor.click();
  await page.keyboard.press(`${modifier}+a`);
  await page.keyboard.press(`${modifier}+d`);

  let dialog = page.getByRole('dialog', { name: '字体高级设置' });
  await expect(dialog).toBeVisible();
  await expect(
    dialog.getByRole('combobox', { name: '拉丁文字字体' }),
  ).toBeFocused();
  const spacing = dialog.getByRole('combobox', { name: '字符间距' });
  await spacing.click();
  await page.getByRole('option', { name: '加宽' }).click();
  await dialog.getByRole('textbox', { name: '间距值（磅）' }).fill('2');
  await page.screenshot({
    path: testInfo.outputPath('writer-character-spacing-dialog.png'),
    animations: 'disabled',
  });
  await dialog.getByRole('button', { name: '应用' }).click();

  const expanded = editor.locator(
    'span[data-office-character-spacing-twips="40"]',
  );
  await expect(expanded).toHaveText('Native character spacing');
  await expect(expanded).toHaveAttribute('style', /letter-spacing:\s*2pt/);
  await expect(editor).toBeFocused();
  await page.keyboard.press(`${modifier}+z`);
  await expect(expanded).toHaveCount(0);

  await page.keyboard.press(`${modifier}+d`);
  dialog = page.getByRole('dialog', { name: '字体高级设置' });
  await expect(dialog).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(dialog).toBeHidden();
  await expect(editor).toBeFocused();
  expect(browserErrors).toEqual([]);
});
