import { expect, test } from '@playwright/test';

test('Writer previews, resolves conflicts, and undoes native text effects', async ({
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
      name: '文字效果 空心、阴影、阳文与阴文',
    })
    .click();

  const editor = page.getByRole('textbox', { name: '文档正文' });
  await expect(editor).toHaveAttribute('data-pagination-state', 'ready');
  const outlineShadow = editor.locator(
    'span[data-office-legacy-text-outline="true"][data-office-legacy-text-shadow="true"]',
  );
  const emboss = editor.locator('span[data-office-legacy-text-emboss="true"]');
  const imprint = editor.locator(
    'span[data-office-legacy-text-imprint="true"]',
  );
  await expect(outlineShadow).toHaveText('空心 + 阴影');
  await expect(emboss).toHaveText('阳文');
  await expect(imprint).toHaveText('阴文');
  await expect
    .poll(() =>
      outlineShadow.evaluate((element) => {
        const style = getComputedStyle(element);
        return {
          fill: style.getPropertyValue('-webkit-text-fill-color'),
          shadow: style.textShadow,
          stroke: style.getPropertyValue('-webkit-text-stroke-width'),
        };
      }),
    )
    .toEqual(
      expect.objectContaining({
        fill: 'rgba(0, 0, 0, 0)',
        stroke: expect.not.stringMatching(/^0(?:px)?$/),
        shadow: expect.not.stringMatching(/^none$/),
      }),
    );

  await outlineShadow.click();
  if (process.platform === 'darwin') {
    await page.keyboard.press('Meta+ArrowLeft');
    await page.keyboard.press('Meta+Shift+ArrowRight');
  } else {
    await page.keyboard.press('Home');
    await page.keyboard.press('Shift+End');
  }
  await expect
    .poll(() => page.evaluate(() => window.getSelection()?.toString()))
    .toBe('空心 + 阴影');
  await expect(
    page.getByRole('toolbar', { name: '文本快捷工具栏' }),
  ).toBeVisible();
  await page.keyboard.press(`${modifier}+d`);

  const dialog = page.getByRole('dialog', { name: '字体高级设置' });
  await expect(dialog).toBeVisible();
  const effects = dialog.getByRole('group', { name: '文字效果' });
  const outlineCheckbox = effects.getByRole('checkbox', {
    name: '空心',
    exact: true,
  });
  const shadowCheckbox = effects.getByRole('checkbox', {
    name: '阴影',
    exact: true,
  });
  const embossCheckbox = effects.getByRole('checkbox', {
    name: '阳文',
    exact: true,
  });
  const imprintCheckbox = effects.getByRole('checkbox', {
    name: '阴文',
    exact: true,
  });
  await expect(outlineCheckbox).toBeChecked();
  await expect(shadowCheckbox).toBeChecked();
  await expect(embossCheckbox).not.toBeChecked();
  await expect(imprintCheckbox).not.toBeChecked();

  await embossCheckbox.click();
  await expect(outlineCheckbox).not.toBeChecked();
  await expect(shadowCheckbox).not.toBeChecked();
  await expect(embossCheckbox).toBeChecked();
  await expect(imprintCheckbox).not.toBeChecked();
  await expect(
    dialog.getByLabel('字符高级格式预览').locator('output > span'),
  ).toHaveAttribute('style', /text-shadow/);
  await page.screenshot({
    path: testInfo.outputPath(
      `writer-text-effects-dialog-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  await dialog.getByRole('button', { name: '应用' }).click();

  const resolved = editor.locator(
    'span[data-office-legacy-text-outline="false"][data-office-legacy-text-shadow="false"][data-office-legacy-text-emboss="true"][data-office-legacy-text-imprint="false"]',
  );
  await expect(resolved).toHaveText('空心 + 阴影');
  await expect(editor).toHaveAttribute('data-pagination-text-engine', 'wasm');
  await expect(editor).toBeFocused();

  await page.keyboard.press(`${modifier}+z`);
  await expect(resolved).toHaveCount(0);
  await expect(outlineShadow).toHaveText('空心 + 阴影');
  expect(browserErrors).toEqual([]);
});
