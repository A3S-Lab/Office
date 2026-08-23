import { expect, test } from '@playwright/test';

test('Writer authors, reveals, toggles, and undoes native hidden text', async ({
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
  const editorRoot = page.locator('.work-document-editor');
  await expect(editor).toHaveAttribute('data-pagination-state', 'ready');
  await editor.fill('Native hidden text');
  await editor.click();
  await page.keyboard.press(`${modifier}+a`);
  await page.keyboard.press(`${modifier}+d`);

  const dialog = page.getByRole('dialog', { name: '字体高级设置' });
  await expect(dialog).toBeVisible();
  await expect(
    dialog.getByRole('group', {
      name: '字符缩放、间距、字距调整、位置与文字效果',
    }),
  ).toBeVisible();
  const hiddenTextCheckbox = dialog.getByRole('checkbox', {
    name: '隐藏文字',
  });
  await expect(hiddenTextCheckbox).not.toBeChecked();
  await hiddenTextCheckbox.click();
  await expect(hiddenTextCheckbox).toBeChecked();

  const preview = dialog
    .getByLabel('字符高级格式预览')
    .locator('output > span');
  await expect(preview).toHaveAttribute(
    'style',
    /text-decoration-line:\s*underline/,
  );
  await expect(preview).toHaveAttribute(
    'style',
    /text-decoration-style:\s*dotted/,
  );
  await page.screenshot({
    path: testInfo.outputPath(
      `writer-hidden-text-dialog-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  await dialog.getByRole('button', { name: '应用' }).click();

  const hidden = editor.locator('span[data-office-hidden-text="true"]');
  await expect(hidden).toHaveText('Native hidden text');
  await expect(editorRoot).not.toHaveClass(/show-hidden-text/);
  await expect
    .poll(() => hidden.evaluate((element) => getComputedStyle(element).display))
    .toBe('none');
  await expect(editor).toBeFocused();

  await page.getByRole('tab', { name: '视图' }).click();
  const showHiddenText = page.getByRole('button', { name: '显示隐藏文字' });
  await expect(showHiddenText).toHaveAttribute('aria-pressed', 'false');
  await showHiddenText.click();
  await expect(showHiddenText).toHaveAttribute('aria-pressed', 'true');
  await expect(editorRoot).toHaveClass(/show-hidden-text/);
  await expect
    .poll(() =>
      hidden.evaluate((element) => {
        const style = getComputedStyle(element);
        return {
          display: style.display,
          line: style.textDecorationLine,
          style: style.textDecorationStyle,
        };
      }),
    )
    .toEqual({ display: 'inline', line: 'underline', style: 'dotted' });
  await expect(editor).toBeFocused();

  await editor.click();
  await page.keyboard.press(`${modifier}+a`);
  await page.keyboard.press(`${modifier}+Shift+h`);
  const visibleReset = editor.locator('span[data-office-hidden-text="false"]');
  await expect(hidden).toHaveCount(0);
  await expect(visibleReset).toHaveText('Native hidden text');

  await page.keyboard.press(`${modifier}+z`);
  await expect(visibleReset).toHaveCount(0);
  await expect(hidden).toHaveText('Native hidden text');
  await expect
    .poll(() =>
      hidden.evaluate(
        (element) => getComputedStyle(element).textDecorationStyle,
      ),
    )
    .toBe('dotted');
  expect(browserErrors).toEqual([]);
});
