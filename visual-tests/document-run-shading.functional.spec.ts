import { expect, test, type Locator } from '@playwright/test';

test('Writer authors and undoes exact native character shading', async ({
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
  const latestCapabilities = page.getByRole('region', { name: '最新能力' });
  const characterShadingEntry = latestCapabilities.getByRole('button', {
    name: '打开最新能力：字符底纹',
  });
  await expect(latestCapabilities).toBeInViewport();
  await expect(characterShadingEntry).toBeInViewport();
  await characterShadingEntry.click();

  const editor = page.getByRole('textbox', { name: '文档正文' });
  await expect(editor).toHaveAttribute('data-pagination-state', 'ready');
  const sample = editor.locator('span[data-office-run-shading*="pct25"]');
  await expect(sample).toHaveText('25% 字符底纹');
  await expect
    .poll(() => shadingPresentation(sample))
    .toEqual({
      backgroundColor: 'rgb(221, 235, 247)',
      backgroundImage: expect.stringContaining('radial-gradient'),
      backgroundSize: '5px 5px',
      boxDecorationBreak: 'clone',
      shading: {
        color: { value: '#4472c4' },
        fill: { value: '#ddebf7' },
        pattern: 'pct25',
      },
    });

  await sample.click();
  if (process.platform === 'darwin') {
    await page.keyboard.press('Meta+ArrowLeft');
    await page.keyboard.press('Meta+Shift+ArrowRight');
  } else {
    await page.keyboard.press('Home');
    await page.keyboard.press('Shift+End');
  }
  await expect
    .poll(() => page.evaluate(() => window.getSelection()?.toString()))
    .toBe('25% 字符底纹');
  await expect(
    page.getByRole('toolbar', { name: '文本快捷工具栏' }),
  ).toBeVisible();
  await page.keyboard.press(`${modifier}+d`);

  const dialog = page.getByRole('dialog', { name: '字体高级设置' });
  await expect(dialog).toBeVisible();
  await expect(
    dialog.getByRole('combobox', { name: '字符底纹', exact: true }),
  ).toHaveText('底纹');
  await expect(
    dialog.getByRole('combobox', { name: '字符底纹图案' }),
  ).toHaveText('25%');
  await dialog.getByRole('combobox', { name: '字符底纹图案' }).click();
  await page.getByRole('option', { name: '对角交叉', exact: true }).click();
  await dialog.getByRole('button', { name: '字符底纹前景色' }).click();
  await page.getByRole('option', { name: '颜色 #0070c0' }).click();
  await dialog.getByRole('button', { name: '字符底纹背景色' }).click();
  await page.getByRole('option', { name: '颜色 #d9ead3' }).click();

  const preview = dialog
    .getByLabel('字符高级格式预览')
    .locator('output > span');
  await expect(preview).toHaveCSS('background-color', 'rgb(217, 234, 211)');
  await expect(preview).toHaveCSS(
    'background-image',
    /repeating-linear-gradient/,
  );
  await page.screenshot({
    path: testInfo.outputPath(
      `writer-run-shading-dialog-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  await dialog.getByRole('button', { name: '应用' }).click();

  const updated = editor.getByText('25% 字符底纹', { exact: true });
  const appliedMarker = editor.locator(
    'span[data-office-run-shading*="diagCross"]' +
      '[data-office-run-shading*="#0070c0"]' +
      '[data-office-run-shading*="#d9ead3"]',
  );
  await expect(updated).toHaveText('25% 字符底纹');
  await expect
    .poll(() => shadingPresentation(updated))
    .toEqual({
      backgroundColor: 'rgb(217, 234, 211)',
      backgroundImage: expect.stringContaining('repeating-linear-gradient'),
      backgroundSize: 'auto',
      boxDecorationBreak: 'clone',
      shading: {
        color: { value: '#0070c0' },
        fill: { value: '#d9ead3' },
        pattern: 'diagCross',
      },
    });
  await expect(editor).toBeFocused();

  await page.keyboard.press(`${modifier}+z`);
  await expect(appliedMarker).toHaveCount(0);
  await expect(sample).toHaveText('25% 字符底纹');
  await expect
    .poll(() => shadingPresentation(sample))
    .toMatchObject({
      backgroundColor: 'rgb(221, 235, 247)',
      backgroundImage: expect.stringContaining('radial-gradient'),
      shading: { pattern: 'pct25' },
    });
  expect(browserErrors).toEqual([]);
});

async function shadingPresentation(locator: Locator): Promise<{
  backgroundColor: string;
  backgroundImage: string;
  backgroundSize: string;
  boxDecorationBreak: string;
  shading: unknown;
}> {
  return locator.evaluate((element) => {
    const style = getComputedStyle(element);
    const serialized = element.getAttribute('data-office-run-shading');
    return {
      backgroundColor: style.backgroundColor,
      backgroundImage: style.backgroundImage,
      backgroundSize: style.backgroundSize,
      boxDecorationBreak: style.boxDecorationBreak,
      shading: serialized ? JSON.parse(serialized) : null,
    };
  });
}
