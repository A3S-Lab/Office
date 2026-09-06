import { expect, type Page, test } from '@playwright/test';

test('Presentation phone comments restore the exact ribbon invoker', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });
  page.on('pageerror', (error) => browserErrors.push(error.message));

  await page.setViewportSize({ width: 390, height: 700 });
  await openPresentationFixture(page);
  await page.getByRole('tab', { name: '审阅', exact: true }).click();
  await page.getByRole('button', { name: '新建批注' }).click();
  const dialog = page.getByRole('dialog', { name: '批注内容' });
  await dialog
    .getByRole('textbox', { name: '批注内容' })
    .fill('需要补充这项结论的数据来源');
  await dialog.getByRole('button', { name: '添加批注' }).click();

  const pane = page.getByRole('dialog', {
    name: '演示批注审阅',
    exact: true,
  });
  await expect(pane).toHaveAttribute('aria-modal', 'true');
  await expect(page.locator('.work-presentation-ribbon')).toHaveAttribute(
    'inert',
    '',
  );
  await expect(page.locator('.work-presentation-workspace')).toHaveAttribute(
    'inert',
    '',
  );
  await expect(
    pane.getByRole('button', { name: '关闭演示批注审阅' }),
  ).toBeFocused();

  const draft = pane.getByRole('textbox', { name: '编辑演示批注 1' });
  await draft.fill('尚未提交的批注草稿');
  await draft.press('Escape');
  await expect(pane).toBeVisible();
  await page.screenshot({
    path: testInfo.outputPath('presentation-phone-comments.png'),
    animations: 'disabled',
  });

  await pane.getByRole('button', { name: '关闭演示批注审阅' }).click();
  await expect(pane).toBeHidden();
  const newComment = page.getByRole('button', { name: '新建批注' });
  await expect(newComment).toBeFocused();
  const viewComments = page.getByRole('button', { name: '查看批注（1）' });
  await viewComments.click();
  await expect(pane).toBeVisible();
  await pane.getByRole('button', { name: '关闭演示批注审阅' }).click();
  await expect(pane).toBeHidden();
  await expect(viewComments).toBeFocused();
  await expect(page.locator('.work-presentation-ribbon')).not.toHaveAttribute(
    'inert',
    '',
  );
  await expect(
    page.locator('.work-presentation-workspace'),
  ).not.toHaveAttribute('inert', '');
  expect(browserErrors).toEqual([]);
});

async function openPresentationFixture(page: Page): Promise<void> {
  await page.goto('/playground/');
  await page
    .getByRole('button', {
      name: '业务策略汇报 PPTX · 本次会话',
    })
    .click();
  await page.locator('.work-slide-canvas.interactive').waitFor();
  await expect(page.locator('.work-presentation-editor')).toHaveAttribute(
    'data-presentation-geometry-state',
    'idle',
  );
}
