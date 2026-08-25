import { expect, test } from '@playwright/test';

test('Writer authors native character scale through the responsive font dialog', async ({
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
  await editor.fill('Native character scale');
  await editor.click();
  await page.keyboard.press(`${modifier}+a`);
  await page.keyboard.press(`${modifier}+d`);

  const dialog = page.getByRole('dialog', { name: '字体高级设置' });
  await expect(dialog).toBeVisible();
  const scale = dialog.getByRole('textbox', {
    name: '字符缩放比例（%）',
  });
  await expect(scale).toBeVisible();
  await scale.fill('80');
  await expect(
    dialog.getByLabel('字符高级格式预览').locator('output'),
  ).toHaveAttribute('style', /font-stretch:\s*80%/);
  await expect(dialog.getByRole('button', { name: '应用' })).toBeEnabled();
  await page.screenshot({
    path: testInfo.outputPath(
      `writer-character-scale-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  await dialog.getByRole('button', { name: '应用' }).click();

  const scaled = editor.locator(
    'span[data-office-character-scale-percent="80"]',
  );
  await expect(scaled).toHaveText('Native character scale');
  await expect(scaled).toHaveAttribute('style', /font-stretch:\s*80%/);
  await expect(editor).toBeFocused();
  await expect
    .poll(() => page.evaluate(() => window.getSelection()?.toString()))
    .toBe('Native character scale');

  await page.keyboard.press(`${modifier}+z`);
  await expect(scaled).toHaveCount(0);
  expect(browserErrors).toEqual([]);
});
