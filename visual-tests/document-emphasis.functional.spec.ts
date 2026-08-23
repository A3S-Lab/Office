import { expect, test } from '@playwright/test';

test('Writer authors and explicitly clears a native emphasis mark in the responsive font dialog', async ({
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
  await editor.fill('Native emphasis mark');
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
  await dialog.getByRole('combobox', { name: '着重号' }).click();
  await page.getByRole('option', { name: '上方圆圈', exact: true }).click();

  const preview = dialog
    .getByLabel('字符高级格式预览')
    .locator('output > span');
  await expect(preview).toHaveAttribute(
    'style',
    /text-emphasis-style:\s*open circle/,
  );
  await expect(preview).toHaveAttribute(
    'style',
    /text-emphasis-position:\s*over right/,
  );
  await expect(dialog.getByRole('button', { name: '应用' })).toBeEnabled();
  await page.screenshot({
    path: testInfo.outputPath(
      `writer-emphasis-dialog-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  await dialog.getByRole('button', { name: '应用' }).click();

  const emphasized = editor.locator('span[data-office-emphasis-mark="circle"]');
  await expect(emphasized).toHaveText('Native emphasis mark');
  await expect(emphasized).toHaveAttribute(
    'style',
    /text-emphasis-style:\s*open circle/,
  );
  await expect
    .poll(() =>
      emphasized.evaluate((element) =>
        getComputedStyle(element)
          .getPropertyValue('text-emphasis-style')
          .trim(),
      ),
    )
    .toBe('open circle');
  await expect(editor).toBeFocused();
  await expect
    .poll(() => page.evaluate(() => window.getSelection()?.toString()))
    .toBe('Native emphasis mark');

  await page.keyboard.press(`${modifier}+d`);
  dialog = page.getByRole('dialog', { name: '字体高级设置' });
  await expect(dialog).toBeVisible();
  await dialog.getByRole('combobox', { name: '着重号' }).click();
  await page.getByRole('option', { name: '无', exact: true }).click();
  await expect(
    dialog.getByLabel('字符高级格式预览').locator('output > span'),
  ).toHaveAttribute('style', /text-emphasis-style:\s*none/);
  await dialog.getByRole('button', { name: '应用' }).click();

  const explicitNone = editor.locator('span[data-office-emphasis-mark="none"]');
  await expect(emphasized).toHaveCount(0);
  await expect(explicitNone).toHaveText('Native emphasis mark');
  await expect(explicitNone).toHaveAttribute(
    'style',
    /text-emphasis-style:\s*none/,
  );
  await expect(editor).toBeFocused();

  await page.keyboard.press(`${modifier}+z`);
  await expect(explicitNone).toHaveCount(0);
  await expect(emphasized).toHaveText('Native emphasis mark');
  expect(browserErrors).toEqual([]);
});
