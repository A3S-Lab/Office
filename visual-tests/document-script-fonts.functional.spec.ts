import { expect, test } from '@playwright/test';

test('Writer renders independent Latin, East Asian, and complex-script fonts', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  const modifier = process.platform === 'darwin' ? 'Meta' : 'Control';
  page.on('pageerror', (error) => browserErrors.push(error.message));
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });

  await page.goto('/playground/');
  await page
    .getByRole('button', {
      name: '空白文字 从一张干净的 A4 页面开始',
    })
    .click();

  const editor = page.getByRole('textbox', { name: '文档正文' });
  await expect(editor).toHaveAttribute('data-pagination-state', 'ready');
  await editor.fill('Latin 中文 العربية');
  await editor.click();
  await page.keyboard.press(`${modifier}+a`);
  await page.keyboard.press(`${modifier}+d`);

  const dialog = page.getByRole('dialog', { name: '字体高级设置' });
  await expect(dialog).toBeVisible();
  const latinFont = dialog.getByRole('combobox', { name: '拉丁文字字体' });
  await expect(latinFont).toBeFocused();
  await latinFont.click();
  await page.getByRole('option', { name: 'Times New Roman' }).click();
  await dialog.getByRole('combobox', { name: '东亚文字字体' }).click();
  await page.getByRole('option', { name: '思源黑体' }).click();
  await dialog.getByRole('combobox', { name: '复杂文字字体' }).click();
  await page.getByRole('option', { name: 'Noto Naskh Arabic' }).click();

  const preview = dialog.getByLabel('字符高级格式预览');
  await expect(preview.locator('span[style*="Times New Roman"]')).toContainText(
    'Latin',
  );
  await expect(
    preview.locator('span[style*="A3S Office Noto Sans Hans"]'),
  ).toContainText('中文');
  await expect(
    preview.locator('span[style*="A3S Office Noto Naskh Arabic"]'),
  ).toContainText('العربية');
  await page.screenshot({
    path: testInfo.outputPath(
      `writer-script-fonts-dialog-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  await dialog.getByRole('button', { name: '应用' }).click();

  const latin = editor.locator('span[data-office-script-font-slot="ascii"]');
  const eastAsian = editor.locator(
    'span[data-office-script-font-slot="eastAsia"]',
  );
  const complexScript = editor.locator(
    'span[data-office-script-font-slot="complexScript"]',
  );
  await expect(latin).toContainText('Latin');
  await expect(eastAsian).toContainText('中文');
  await expect(complexScript).toContainText('العربية');
  await expect
    .poll(() =>
      latin.first().evaluate((node) => getComputedStyle(node).fontFamily),
    )
    .toContain('Times New Roman');
  await expect
    .poll(() => eastAsian.evaluate((node) => getComputedStyle(node).fontFamily))
    .toContain('A3S Office Noto Sans Hans');
  await expect
    .poll(() =>
      complexScript.evaluate((node) => getComputedStyle(node).fontFamily),
    )
    .toContain('A3S Office Noto Naskh Arabic');
  await expect(editor).toBeFocused();

  await page.keyboard.press(`${modifier}+z`);
  await expect(editor.locator('span[data-office-script-fonts]')).toHaveCount(0);
  expect(browserErrors).toEqual([]);
});
