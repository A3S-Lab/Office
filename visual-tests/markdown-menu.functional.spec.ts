import { expect, type Locator, test } from '@playwright/test';

test('Markdown insert dialogs return focus to the invoking editing surface', async ({
  page,
}) => {
  await page.goto('/');
  await page.getByRole('button', { name: '# 产品说明 MD · 本次会话' }).click();

  const source = page.getByRole('textbox', { name: 'Markdown 源码' });
  const preview = page.getByRole('document', { name: 'Markdown 预览' });
  await expect(source).toBeVisible();
  await expect(preview).toBeVisible();
  await expect(preview).toHaveAttribute('contenteditable', 'false');

  await source.focus();
  await page.getByRole('tab', { name: '插入' }).click();
  await page.getByRole('button', { name: '添加链接' }).click();
  await expect(page.getByRole('dialog', { name: '添加链接' })).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(source).toBeFocused();

  await page.getByRole('tab', { name: '视图' }).click();
  await page
    .getByRole('region', { name: '编辑方式' })
    .getByRole('button', { name: '编辑' })
    .click();
  const visual = page.getByRole('textbox', { name: 'Markdown 编辑区' });
  await visual.focus();
  await page.getByRole('tab', { name: '插入' }).click();
  await page.getByRole('button', { name: '添加链接' }).click();
  await expect(page.getByRole('dialog', { name: '添加链接' })).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(visual).toBeFocused();
});

test('Markdown split mode keeps source editable and preview read-only', async ({
  page,
}) => {
  await page.goto('/');
  await page.getByRole('button', { name: '# 产品说明 MD · 本次会话' }).click();

  const sourcePane = page.getByRole('region', {
    name: 'Markdown 源码窗格',
  });
  const previewPane = page.getByRole('region', {
    name: 'Markdown 预览窗格',
  });
  await expect(sourcePane.getByText('编辑', { exact: true })).toBeVisible();
  await expect(previewPane.getByText('预览', { exact: true })).toBeVisible();

  const source = page.getByRole('textbox', { name: 'Markdown 源码' });
  const preview = page.getByRole('document', { name: 'Markdown 预览' });
  await expect(preview).toHaveAttribute('aria-readonly', 'true');
  await expect(preview).toHaveAttribute('contenteditable', 'false');

  await source.fill('# Single editing source');
  await expect(
    preview.getByRole('heading', { name: 'Single editing source' }),
  ).toBeVisible();

  await preview.press('x');
  await expect(source).toHaveValue('# Single editing source');
});

test('Markdown edits and removes the exact link selected in source mode', async ({
  page,
}) => {
  await page.goto('/');
  await page.getByRole('button', { name: '# 产品说明 MD · 本次会话' }).click();

  const source = page.getByRole('textbox', { name: 'Markdown 源码' });
  const markdown = '[Office](https://a3s.dev) after';
  await source.fill(markdown);
  await source.evaluate((element: HTMLTextAreaElement) => {
    element.focus();
    element.setSelectionRange(3, 3, 'none');
    element.dispatchEvent(new Event('select', { bubbles: true }));
  });
  await page.getByRole('tab', { name: '插入' }).click();

  const editLink = page.getByRole('button', { name: '编辑链接' });
  const removeLink = page.getByRole('button', { name: '移除链接' });
  await expect(editLink).toHaveAttribute('aria-pressed', 'true');
  await expect(removeLink).toBeEnabled();

  await editLink.click();
  const dialog = page.getByRole('dialog', { name: '编辑链接' });
  await expect(dialog.getByRole('textbox', { name: '显示文字' })).toHaveValue(
    'Office',
  );
  await expect(dialog.getByRole('textbox', { name: '链接地址' })).toHaveValue(
    'https://a3s.dev',
  );
  await dialog
    .getByRole('textbox', { name: '链接地址' })
    .fill('https://a3s.dev/cancelled');
  await page.keyboard.press('Escape');
  await expect(dialog).toBeHidden();
  await expect(source).toHaveValue(markdown);
  await expect(source).toBeFocused();
  await expect
    .poll(() =>
      source.evaluate((element) => ({
        end: element.selectionEnd,
        start: element.selectionStart,
      })),
    )
    .toEqual({ end: 3, start: 3 });

  await editLink.click();
  const editDialog = page.getByRole('dialog', { name: '编辑链接' });
  await editDialog
    .getByRole('textbox', { name: '显示文字' })
    .fill('Office docs');
  const linkSource = editDialog.getByRole('textbox', { name: '链接地址' });
  await linkSource.fill('https://a3s.dev/office');
  await linkSource.press('Enter');

  await expect(source).toHaveValue(
    '[Office docs](https://a3s.dev/office) after',
  );
  await expect(source).toBeFocused();
  await expect
    .poll(() =>
      source.evaluate((element) => ({
        end: element.selectionEnd,
        start: element.selectionStart,
      })),
    )
    .toEqual({ end: 12, start: 1 });

  await removeLink.click();
  await expect(source).toHaveValue('Office docs after');
  await expect(source).toBeFocused();
});

test('Markdown edits and removes links without losing the visual selection', async ({
  page,
}) => {
  await page.goto('/');
  await page.getByRole('button', { name: '# 产品说明 MD · 本次会话' }).click();

  const source = page.getByRole('textbox', { name: 'Markdown 源码' });
  await source.fill('[Office](https://a3s.dev)');
  await page.getByRole('tab', { name: '视图' }).click();
  await page
    .getByRole('region', { name: '编辑方式' })
    .getByRole('button', { name: '编辑' })
    .click();

  const visual = page.getByRole('textbox', { name: 'Markdown 编辑区' });
  const visualLink = visual.getByRole('link', { name: 'Office' });
  await visualLink.click();
  await page.getByRole('tab', { name: '插入' }).click();
  await page.getByRole('button', { name: '编辑链接' }).click();

  const dialog = page.getByRole('dialog', { name: '编辑链接' });
  const linkSource = dialog.getByRole('textbox', { name: '链接地址' });
  await linkSource.fill('https://a3s.dev/office');
  await dialog.getByRole('button', { name: '保存' }).click();
  await expect(visualLink).toHaveAttribute('href', 'https://a3s.dev/office');
  await expect(visual).toBeFocused();

  await page.getByRole('button', { name: '移除链接' }).click();
  await expect(visual.getByRole('link', { name: 'Office' })).toHaveCount(0);
  await expect(visual).toContainText('Office');
  await expect(visual).toBeFocused();
});

test('Markdown insert dialogs reject unsafe sources and preserve relative images', async ({
  page,
}) => {
  await page.goto('/');
  await page.getByRole('button', { name: '# 产品说明 MD · 本次会话' }).click();

  const source = page.getByRole('textbox', { name: 'Markdown 源码' });
  await expect(source).toBeVisible();
  await source.focus();
  await page.getByRole('tab', { name: '插入' }).click();
  await page.getByRole('button', { name: '添加链接' }).click();

  const linkDialog = page.getByRole('dialog', { name: '添加链接' });
  const linkSource = linkDialog.getByRole('textbox', { name: '链接地址' });
  const addLink = linkDialog.getByRole('button', { name: '添加' });
  await linkSource.fill('javascript:alert(1)');
  await expect(linkSource).toHaveAttribute('aria-invalid', 'true');
  await expect(linkDialog.getByRole('alert')).toHaveText(
    '请输入完整的 http、https、mailto 或 # 文档内地址。',
  );
  await expect(addLink).toBeDisabled();

  await linkDialog.getByRole('button', { name: '取消' }).click();
  await expect(source).toBeFocused();
  await page.getByRole('button', { name: '插入图片' }).click();

  const imageDialog = page.getByRole('dialog', { name: '插入图片' });
  await imageDialog
    .getByRole('textbox', { name: '替代文字（可选）' })
    .fill('Office 架构图');
  await imageDialog
    .getByRole('textbox', { name: '图片地址' })
    .fill('../assets/office diagram.png');
  await imageDialog.getByRole('button', { name: '插入' }).click();

  await expect(source).toHaveValue(
    /!\[Office 架构图\]\(\.\.\/assets\/office%20diagram\.png\)/,
  );
  await expect(source).toBeFocused();
});

test('Markdown view controls return keyboard focus to the selected surface', async ({
  page,
}) => {
  await page.goto('/');
  await page.getByRole('button', { name: '# 产品说明 MD · 本次会话' }).click();

  await page.getByRole('tab', { name: '视图' }).click();
  const viewControls = page.getByRole('region', { name: '编辑方式' });
  await viewControls.getByRole('button', { name: '源码' }).click();
  const source = page.getByRole('textbox', { name: 'Markdown 源码' });
  await expect(source).toBeFocused();
  const sourceBefore = await source.inputValue();
  await page.keyboard.press('Meta+b');
  await expect(source).not.toHaveValue(sourceBefore);

  await viewControls.getByRole('button', { name: '编辑' }).click();
  const visual = page.getByRole('textbox', { name: 'Markdown 编辑区' });
  await expect(visual).toBeFocused();

  await viewControls.getByRole('button', { name: '分屏' }).click();
  await expect(
    page.getByRole('textbox', { name: 'Markdown 源码' }),
  ).toBeFocused();
});

test('Markdown source selection keeps ribbon state synchronized', async ({
  page,
}) => {
  await page.goto('/');
  await page.getByRole('button', { name: '# 产品说明 MD · 本次会话' }).click();
  await page.getByRole('tab', { name: '视图' }).click();
  await page
    .getByRole('region', { name: '编辑方式' })
    .getByRole('button', { name: '源码' })
    .click();
  await page.getByRole('tab', { name: '开始' }).click();

  const source = page.getByRole('textbox', { name: 'Markdown 源码' });
  const bold = page.getByRole('button', { name: '加粗' });
  await source.fill('**bold** plain');
  await expect(bold).toHaveAttribute('aria-pressed', 'false');

  await setMarkdownSourceSelection(source, 4);
  await expect(bold).toHaveAttribute('aria-pressed', 'true');
  await bold.click();
  await expect(source).toHaveValue('bold plain');
  await expect(source).toBeFocused();
  await expect(bold).toHaveAttribute('aria-pressed', 'false');

  await source.fill('**bold** plain');

  await setMarkdownSourceSelection(source, 2, 6);
  await expect
    .poll(() =>
      source.evaluate((element) => ({
        end: element.selectionEnd,
        start: element.selectionStart,
      })),
    )
    .toEqual({ end: 6, start: 2 });
  await expect(bold).toHaveAttribute('aria-pressed', 'true');

  await setMarkdownSourceSelection(source, 9, 14);
  await expect(bold).toHaveAttribute('aria-pressed', 'false');

  await setMarkdownSourceSelection(source, 2, 6);
  await expect(bold).toHaveAttribute('aria-pressed', 'true');
  await bold.click();
  await expect(source).toHaveValue('bold plain');
  await expect(source).toBeFocused();
  await expect(bold).toHaveAttribute('aria-pressed', 'false');
});

test('Markdown opens source selection actions from the keyboard', async ({
  page,
}) => {
  await page.goto('/');
  await page.getByRole('button', { name: '# 产品说明 MD · 本次会话' }).click();

  const source = page.getByRole('textbox', { name: 'Markdown 源码' });
  await source.evaluate((element: HTMLTextAreaElement) => {
    element.focus();
    element.setSelectionRange(2, 12, 'forward');
    element.dispatchEvent(new Event('select', { bubbles: true }));
  });
  await source.press('Shift+F10');

  const menu = page.getByRole('menu', { name: '选中文本操作' });
  await expect(menu).toBeVisible();
  await expect(menu.getByRole('menuitem').first()).toBeFocused();
  const [sourceBounds, menuBounds] = await Promise.all([
    source.boundingBox(),
    menu.boundingBox(),
  ]);
  if (!(sourceBounds && menuBounds)) {
    throw new Error('Markdown keyboard menu geometry is unavailable.');
  }
  expect(menuBounds.x).toBeGreaterThanOrEqual(sourceBounds.x - 1);
  expect(menuBounds.x).toBeLessThan(sourceBounds.x + sourceBounds.width / 2);
  expect(menuBounds.y).toBeGreaterThanOrEqual(sourceBounds.y - 1);
  expect(menuBounds.y).toBeLessThanOrEqual(sourceBounds.y + 90);
  await page.keyboard.press('Escape');
  await expect(menu).toBeHidden();
  await expect(source).toBeFocused();
});

async function setMarkdownSourceSelection(
  source: Locator,
  start: number,
  end = start,
): Promise<void> {
  await source.evaluate(
    (element: HTMLTextAreaElement, selection) => {
      element.focus();
      element.setSelectionRange(selection.start, selection.end, 'forward');
    },
    { end, start },
  );
  await source.press('Shift');
}
