import { expect, test } from '@playwright/test';

test('Writer authors native proofing languages in the responsive Review dialog', async ({
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
      name: '校对语言 拉丁、东亚、双向文字与校对排除',
    })
    .click();

  const editor = page.getByRole('textbox', { name: '文档正文' });
  await expect(editor).toHaveAttribute('data-pagination-state', 'ready');
  const latin = editor.getByText('English proofing language', { exact: true });
  const eastAsian = editor.getByText('简体中文校对语言', { exact: true });
  const bidi = editor.getByText('لغة التدقيق العربية', { exact: true });
  const excluded = editor.getByText(
    'A3S-API-v2: this product identifier is excluded from proofing.',
    { exact: true },
  );
  const originalLanguages =
    '{"latin":"en-US","eastAsia":"zh-CN","bidi":"ar-SA"}';
  await expect(latin).toHaveAttribute(
    'data-office-proofing-languages',
    originalLanguages,
  );
  await expect(latin).toHaveAttribute('lang', 'en-US');
  await expect(latin).toHaveAttribute('data-office-no-proof', 'false');
  await expect(eastAsian).toHaveAttribute('lang', 'zh-CN');
  await expect(bidi).toHaveAttribute('lang', 'ar-SA');
  await expect(excluded).toHaveAttribute(
    'data-office-proofing-languages',
    '{"latin":"x-none"}',
  );
  await expect(excluded).toHaveAttribute('data-office-no-proof', 'true');
  await expect(excluded).toHaveAttribute('spellcheck', 'false');

  await latin.click();
  if (process.platform === 'darwin') {
    await page.keyboard.press('Meta+ArrowLeft');
    await page.keyboard.press('Meta+Shift+ArrowRight');
  } else {
    await page.keyboard.press('Home');
    await page.keyboard.press('Shift+End');
  }
  await expect
    .poll(() => page.evaluate(() => window.getSelection()?.toString()))
    .toBe('English proofing language');
  await page.getByRole('tab', { name: '审阅' }).click();
  await page.getByRole('button', { name: '设置校对语言' }).click();

  const dialog = page.getByRole('dialog', { name: '设置校对语言' });
  await expect(dialog).toBeVisible();
  const latinLanguage = dialog.getByRole('combobox', {
    name: '拉丁文字校对语言',
  });
  await expect(latinLanguage).toBeFocused();
  await expect(latinLanguage).toHaveValue('en-US');
  await expect(
    dialog.getByRole('combobox', { name: '东亚文字校对语言' }),
  ).toHaveValue('zh-CN');
  await expect(
    dialog.getByRole('combobox', { name: '双向文字校对语言' }),
  ).toHaveValue('ar-SA');
  await expect(dialog.getByRole('combobox', { name: '校对行为' })).toHaveText(
    '检查拼写和语法',
  );
  await expectDialogInsideViewport(dialog);

  await latinLanguage.fill('fr-FR');
  await dialog.getByRole('combobox', { name: '校对行为' }).click();
  await page
    .getByRole('option', { name: '不检查拼写或语法', exact: true })
    .click();
  await page.screenshot({
    path: testInfo.outputPath(
      `writer-proofing-languages-dialog-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  await dialog.getByRole('button', { name: '应用' }).click();

  const updatedLanguages =
    '{"latin":"fr-FR","eastAsia":"zh-CN","bidi":"ar-SA"}';
  await expect(latin).toHaveAttribute(
    'data-office-proofing-languages',
    updatedLanguages,
  );
  await expect(latin).toHaveAttribute('data-office-no-proof', 'true');
  await expect(latin).toHaveAttribute('lang', 'fr-FR');
  await expect(latin).toHaveAttribute('spellcheck', 'false');
  await expect(editor).toBeFocused();
  await expect
    .poll(() => page.evaluate(() => window.getSelection()?.toString()))
    .toBe('English proofing language');

  await page.keyboard.press(`${modifier}+z`);
  await expect(latin).toHaveAttribute(
    'data-office-proofing-languages',
    originalLanguages,
  );
  await expect(latin).toHaveAttribute('data-office-no-proof', 'false');
  await expect(latin).toHaveAttribute('lang', 'en-US');
  expect(browserErrors).toEqual([]);
});

async function expectDialogInsideViewport(
  dialog: import('@playwright/test').Locator,
): Promise<void> {
  const geometry = await dialog.evaluate((element) => {
    const bounds = element.getBoundingClientRect();
    return {
      bottom: bounds.bottom,
      centerX: bounds.left + bounds.width / 2,
      left: bounds.left,
      right: bounds.right,
      top: bounds.top,
      viewportHeight: document.documentElement.clientHeight,
      viewportWidth: document.documentElement.clientWidth,
    };
  });
  expect(geometry.left).toBeGreaterThanOrEqual(8);
  expect(geometry.right).toBeLessThanOrEqual(geometry.viewportWidth - 8);
  expect(geometry.top).toBeGreaterThanOrEqual(8);
  expect(geometry.bottom).toBeLessThanOrEqual(geometry.viewportHeight - 8);
  expect(geometry.centerX).toBeCloseTo(geometry.viewportWidth / 2, 0);
}
