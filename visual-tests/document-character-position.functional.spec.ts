import { expect, test } from '@playwright/test';

test('Writer authors native character position through the font dialog', async ({
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
  await editor.fill('Native character position');
  await editor.click();
  await page.keyboard.press(`${modifier}+a`);
  await page.keyboard.press(`${modifier}+d`);

  const dialog = page.getByRole('dialog', { name: '字体高级设置' });
  await expect(dialog).toBeVisible();
  await dialog.getByRole('combobox', { name: '字符位置' }).click();
  await page.getByRole('option', { name: '提升' }).click();
  await dialog.getByRole('textbox', { name: '位置值（磅）' }).fill('2');
  await expect(
    dialog.getByLabel('字符间距和位置预览').locator('output > span'),
  ).toHaveAttribute('style', /vertical-align:\s*2pt/);
  await page.screenshot({
    path: testInfo.outputPath('writer-character-position-dialog.png'),
    animations: 'disabled',
  });
  await dialog.getByRole('button', { name: '应用' }).click();

  const raised = editor.locator(
    'span[data-office-character-position-half-points="4"]',
  );
  await expect(raised).toHaveText('Native character position');
  await expect(raised).toHaveAttribute(
    'style',
    /--work-document-character-position:\s*2pt/,
  );
  await expect(editor).toBeFocused();
  await page.keyboard.press(`${modifier}+z`);
  await expect(raised).toHaveCount(0);
  expect(browserErrors).toEqual([]);
});

test('Writer keeps character-position controls usable in the compact dialog', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });
  await page.setViewportSize({ width: 560, height: 820 });
  await page.goto('/');
  await page
    .getByRole('button', {
      name: '空白文字 从一张干净的 A4 页面开始',
    })
    .click();
  const editor = page.getByRole('textbox', { name: '文档正文' });
  await editor.fill('Compact position');
  await editor.click();
  const modifier = process.platform === 'darwin' ? 'Meta' : 'Control';
  await page.keyboard.press(`${modifier}+a`);
  await page.keyboard.press(`${modifier}+d`);

  const dialog = page.getByRole('dialog', { name: '字体高级设置' });
  await expect(
    dialog.getByRole('combobox', { name: '字符位置' }),
  ).toBeVisible();
  await dialog.getByRole('combobox', { name: '字符位置' }).click();
  await page.getByRole('option', { name: '降低' }).click();
  const amount = dialog.getByRole('textbox', { name: '位置值（磅）' });
  await expect(amount).toBeVisible();
  await amount.fill('1.5');
  await expect(dialog.getByRole('button', { name: '应用' })).toBeEnabled();
  await page.screenshot({
    path: testInfo.outputPath('writer-character-position-compact.png'),
    animations: 'disabled',
  });
  await dialog.getByRole('button', { name: '应用' }).click();
  const positioned = editor.locator(
    'span[data-office-character-position-half-points="-3"]',
  );
  await expect(positioned).toHaveText('Compact position');
  await expect
    .poll(() => page.evaluate(() => window.getSelection()?.toString()))
    .toBe('Compact position');
  await page.keyboard.press(`${modifier}+Shift+=`);
  await expect(editor.locator('sup')).toHaveText('Compact position');
  await expect
    .poll(() =>
      positioned.evaluate((element) => getComputedStyle(element).verticalAlign),
    )
    .toBe('baseline');
  await page.keyboard.press(`${modifier}+Shift+=`);
  await expect(editor.locator('sup')).toHaveCount(0);
  await expect
    .poll(() =>
      positioned.evaluate((element) => getComputedStyle(element).verticalAlign),
    )
    .not.toBe('baseline');
  expect(browserErrors).toEqual([]);
});
