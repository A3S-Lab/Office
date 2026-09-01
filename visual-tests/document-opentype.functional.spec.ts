import { expect, test, type Locator, type Page } from '@playwright/test';

test('Writer authors native OpenType typography in the responsive font dialog', async ({
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

  await page.goto('/playground/');
  await page
    .getByRole('button', {
      name: '空白文字 从一张干净的 A4 页面开始',
    })
    .click();

  const editor = page.getByRole('textbox', { name: '文档正文' });
  await expect(editor).toHaveAttribute('data-pagination-state', 'ready');
  await editor.fill('Office 0123 typography');
  await editor.click();
  await page.keyboard.press(`${modifier}+a`);
  await page.keyboard.press(`${modifier}+d`);

  const dialog = page.getByRole('dialog', { name: '字体高级设置' });
  await expect(dialog).toBeVisible();
  await expect(
    dialog.getByRole('group', { name: 'OpenType 排版' }),
  ).toBeVisible();
  await chooseOption(dialog, page, 'OpenType 连字', '全部');
  await chooseOption(dialog, page, 'OpenType 数字字形', '旧式数字');
  await chooseOption(dialog, page, 'OpenType 数字间距', '等宽');
  await chooseOption(dialog, page, 'OpenType 样式集', '样式集 4');
  await chooseOption(dialog, page, 'OpenType 上下文替代', '禁用');

  const preview = dialog.getByLabel('字符高级格式预览').locator('output');
  await expect(preview).toHaveAttribute(
    'style',
    /font-feature-settings:\s*"liga", "clig", "hlig", "dlig", "ss04"/,
  );
  await expect(preview).toHaveAttribute(
    'style',
    /font-variant-numeric:\s*oldstyle-nums tabular-nums/,
  );
  await expect(preview).toHaveAttribute(
    'style',
    /font-variant-ligatures:\s*no-contextual/,
  );
  await page.screenshot({
    path: testInfo.outputPath(
      `writer-opentype-dialog-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  await dialog.getByRole('button', { name: '应用' }).click();

  const typography = editor.locator('span[data-office-opentype-features]');
  await expect(typography).toHaveText('Office 0123 typography');
  await expect(typography).toHaveAttribute(
    'data-office-opentype-features',
    /"ligatures":"all".*"numberForm":"oldStyle".*"numberSpacing":"tabular".*"stylisticSets":\[4\].*"contextualAlternates":false/,
  );
  await expect(typography).toHaveAttribute(
    'style',
    /font-feature-settings:\s*"liga"/,
  );
  await expect
    .poll(() =>
      typography.evaluate((element) => ({
        ligatures: getComputedStyle(element).fontVariantLigatures,
        numeric: getComputedStyle(element).fontVariantNumeric,
      })),
    )
    .toEqual({
      ligatures: 'no-contextual',
      numeric: 'oldstyle-nums tabular-nums',
    });
  await expect(editor).toBeFocused();
  await expect
    .poll(() => page.evaluate(() => window.getSelection()?.toString()))
    .toBe('Office 0123 typography');

  await page.keyboard.press(`${modifier}+z`);
  await expect(typography).toHaveCount(0);
  expect(browserErrors).toEqual([]);
});

async function chooseOption(
  dialog: Locator,
  page: Page,
  label: string,
  option: string,
): Promise<void> {
  await dialog.getByRole('combobox', { name: label }).click();
  await page.getByRole('option', { name: option }).click();
}
