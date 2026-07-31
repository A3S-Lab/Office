import { expect, test } from '@playwright/test';

test('Office sidebar becomes modal only when it overlays the page', async ({
  page,
}) => {
  await page.goto('/');
  const compact = (page.viewportSize()?.width ?? 0) <= 839;

  const collapse = page.getByRole('button', { name: '收起办公侧边栏' });
  if (await collapse.isVisible()) await collapse.click();
  const trigger = page.getByRole('button', { name: '展开办公侧边栏' });
  await trigger.click();

  const sidebar = page.locator('.playground-sidebar');
  const close = sidebar.getByRole('button', { name: '收起办公侧边栏' });
  await expect(sidebar).toBeVisible();
  if (compact) {
    await expect(sidebar).toHaveAttribute('role', 'dialog');
    await expect(sidebar).toHaveAttribute('aria-modal', 'true');
    await expect(page.locator('.playground-main-pane')).toHaveAttribute(
      'inert',
      '',
    );
    await expect(close).toBeFocused();
    await page.keyboard.press('Escape');
    await expect(trigger).toBeFocused();
  } else {
    await expect(sidebar).not.toHaveAttribute('role', 'dialog');
    await expect(sidebar).not.toHaveAttribute('aria-modal', 'true');
    await expect(page.locator('.playground-main-pane')).not.toHaveAttribute(
      'inert',
      '',
    );
    await close.click();
  }
  await expect(sidebar).toBeHidden();
});

test('AI assistant becomes modal only when it overlays the editor', async ({
  page,
}) => {
  await page.goto('/');
  const compact = (page.viewportSize()?.width ?? 0) <= 1040;
  const collapse = page.getByRole('button', { name: '收起办公侧边栏' });
  if (await collapse.isVisible()) await collapse.click();
  await page
    .getByRole('button', { name: '新项目方案 DOCX · 本次会话' })
    .click();

  const trigger = page.getByRole('button', { name: '打开 AI 助手' });
  await trigger.click();
  const assistant = page.locator('.playground-assistant');
  const close = assistant.getByRole('button', { name: '关闭 AI 助手' });
  await expect(assistant).toBeVisible();
  if (compact) {
    await expect(assistant).toHaveAttribute('role', 'dialog');
    await expect(assistant).toHaveAttribute('aria-modal', 'true');
    await expect(page.locator('.playground-editor-host')).toHaveAttribute(
      'inert',
      '',
    );
    await expect(close).toBeFocused();
    await page.keyboard.press('Escape');
    await expect(trigger).toBeFocused();
  } else {
    await expect(assistant).not.toHaveAttribute('role', 'dialog');
    await expect(assistant).not.toHaveAttribute('aria-modal', 'true');
    await expect(page.locator('.playground-editor-host')).not.toHaveAttribute(
      'inert',
      '',
    );
    await close.click();
  }
  await expect(assistant).toBeHidden();
});

test('Creation feedback stays clear of the editor status bar', async ({
  page,
}) => {
  await page.goto('/');
  await page
    .getByRole('button', {
      name: '空白 Markdown 用轻量标记编写结构化内容',
    })
    .click();

  const toast = page.locator('.playground-toast.success');
  const status = page.locator('.work-markdown-status');
  await expect(toast).toBeVisible();
  await expect(status).toBeVisible();

  const boxes = await Promise.all([toast.boundingBox(), status.boundingBox()]);
  const toastBox = boxes[0];
  const statusBox = boxes[1];
  if (!toastBox || !statusBox) {
    throw new Error('Toast or editor status geometry is unavailable.');
  }
  expect(toastBox.y + toastBox.height).toBeLessThanOrEqual(statusBox.y - 8);
});
