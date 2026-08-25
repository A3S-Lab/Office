import { expect, test } from '@playwright/test';

test('Writer authors and undoes an exact native character border', async ({
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
      name: '字符边框 原生线型、颜色、宽度、间距与阴影',
    })
    .click();

  const editor = page.getByRole('textbox', { name: '文档正文' });
  await expect(editor).toHaveAttribute('data-pagination-state', 'ready');
  const sample = editor.locator('span[data-office-run-border*="single"]');
  await expect(sample).toHaveText('单实线字符边框');
  await expect
    .poll(() =>
      sample.evaluate((element) => {
        const style = getComputedStyle(element);
        return {
          color: style.borderColor,
          declaredWidth: (element as HTMLElement).style.borderWidth,
          style: style.borderStyle,
        };
      }),
    )
    .toEqual({
      color: 'rgb(68, 114, 196)',
      declaredWidth: '0.667px',
      style: 'solid',
    });
  const renderedBorderWidth = await sample.evaluate((element) => {
    const stage = element.closest<HTMLElement>(
      '[data-testid="document-page-stage"]',
    );
    const pageZoom = stage
      ? Number.parseFloat(getComputedStyle(stage).zoom || '1')
      : 1;
    return Number.parseFloat(getComputedStyle(element).borderWidth) * pageZoom;
  });
  expect(renderedBorderWidth).toBeGreaterThanOrEqual(2 / 3);
  expect(renderedBorderWidth).toBeLessThanOrEqual(1.001);

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
    .toBe('单实线字符边框');
  await page.keyboard.press(`${modifier}+d`);

  const dialog = page.getByRole('dialog', { name: '字体高级设置' });
  await expect(dialog).toBeVisible();
  await expect(
    dialog.getByRole('combobox', { name: '字符边框', exact: true }),
  ).toHaveText('边框');
  await dialog.getByRole('combobox', { name: '字符边框线型' }).click();
  await page.getByRole('option', { name: '双波浪线' }).click();
  await dialog.getByRole('textbox', { name: '字符边框宽度（磅）' }).fill('1.5');
  await dialog.getByRole('textbox', { name: '字符边框间距（磅）' }).fill('3');
  await dialog.getByRole('button', { name: '字符边框颜色' }).click();
  await page.getByRole('option', { name: '颜色 #0070c0' }).click();
  await dialog.getByRole('checkbox', { name: '字符边框阴影' }).check();
  await dialog.getByRole('checkbox', { name: '字符边框框架' }).check();
  await expect(
    dialog.getByLabel('字符高级格式预览').locator('output > span'),
  ).toHaveAttribute('style', /border:\s*2px double rgb\(0, 112, 192\)/);
  await page.screenshot({
    path: testInfo.outputPath(
      `writer-run-border-dialog-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  await dialog.getByRole('button', { name: '应用' }).click();

  const updated = editor.locator('span[data-office-run-border*="doubleWave"]');
  await expect(updated).toHaveText('单实线字符边框');
  await expect(updated).toHaveCSS('border-style', 'double');
  await expect(updated).toHaveCSS('border-color', 'rgb(0, 112, 192)');
  await expect(editor).toBeFocused();
  await page.keyboard.press(`${modifier}+z`);
  await expect(updated).toHaveCount(0);
  await expect(sample).toHaveText('单实线字符边框');
  expect(browserErrors).toEqual([]);
});
