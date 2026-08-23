import { expect, test } from '@playwright/test';

test('Writer authors and clears a native kerning threshold in the responsive font dialog', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  const modifier = process.platform === 'darwin' ? 'Meta' : 'Control';
  page.on('pageerror', (error) => browserErrors.push(error.message));
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });
  if (testInfo.project.name === 'compact-768') {
    await page.setViewportSize({ width: 390, height: 844 });
  }

  await page.goto('/');
  await page
    .getByRole('button', {
      name: '空白文字 从一张干净的 A4 页面开始',
    })
    .click();

  const editor = page.getByRole('textbox', { name: '文档正文' });
  await expect(editor).toHaveAttribute('data-pagination-state', 'ready');
  await editor.fill('Native kerning threshold');
  await editor.click();
  await page.keyboard.press(`${modifier}+a`);
  await page.keyboard.press(`${modifier}+d`);

  let dialog = page.getByRole('dialog', { name: '字体高级设置' });
  await expect(dialog).toBeVisible();
  await expect(
    dialog.getByRole('group', {
      name: '字符缩放、间距、字距调整、位置、着重号与隐藏文字',
    }),
  ).toBeVisible();
  const kerning = dialog.getByRole('checkbox', {
    name: '为字号达到以下值的字体调整字距',
  });
  const threshold = dialog.getByRole('textbox', {
    name: '字距调整阈值（磅）',
  });
  await expect(kerning).not.toBeChecked();
  await expect(threshold).toBeDisabled();
  await kerning.click();
  await expect(kerning).toBeChecked();
  await expect(threshold).toBeEnabled();
  await threshold.fill('10');
  await expect(
    dialog.getByLabel('字符高级格式预览').locator('output'),
  ).toHaveAttribute('style', /font-kerning:\s*normal/);
  await expect(dialog.getByRole('button', { name: '应用' })).toBeEnabled();
  await page.screenshot({
    path: testInfo.outputPath(
      `writer-kerning-dialog-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  await dialog.getByRole('button', { name: '应用' }).click();

  const kerned = editor.locator(
    'span[data-office-kerning-threshold-half-points="20"]',
  );
  await expect(kerned).toHaveText('Native kerning threshold');
  await expect(kerned).toHaveAttribute('style', /font-kerning:\s*normal/);
  await expect
    .poll(() =>
      kerned.evaluate((element) => getComputedStyle(element).fontKerning),
    )
    .toBe('normal');
  await expect(editor).toBeFocused();
  await expect
    .poll(() => page.evaluate(() => window.getSelection()?.toString()))
    .toBe('Native kerning threshold');

  await page.keyboard.press(`${modifier}+d`);
  dialog = page.getByRole('dialog', { name: '字体高级设置' });
  await expect(dialog).toBeVisible();
  const enabledKerning = dialog.getByRole('checkbox', {
    name: '为字号达到以下值的字体调整字距',
  });
  await expect(enabledKerning).toBeChecked();
  await enabledKerning.click();
  await expect(
    dialog.getByRole('textbox', { name: '字距调整阈值（磅）' }),
  ).toBeDisabled();
  await expect(
    dialog.getByLabel('字符高级格式预览').locator('output'),
  ).toHaveAttribute('style', /font-kerning:\s*none/);
  await dialog.getByRole('button', { name: '应用' }).click();
  await expect(kerned).toHaveCount(0);
  await expect(editor).toBeFocused();

  await page.keyboard.press(`${modifier}+z`);
  await expect(kerned).toHaveText('Native kerning threshold');
  await expect(kerned).toHaveAttribute('style', /font-kerning:\s*normal/);
  expect(browserErrors).toEqual([]);
});
