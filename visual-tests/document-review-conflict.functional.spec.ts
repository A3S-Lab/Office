import { expect, test } from '@playwright/test';

test.use({ viewport: { width: 390, height: 844 } });

test('reports reviewed ranges changed by a controlled host update', async ({
  page,
}) => {
  const consoleErrors: string[] = [];
  const pageErrors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') consoleErrors.push(message.text());
  });
  page.on('pageerror', (error) => pageErrors.push(error.message));

  await page.goto('/?e2e=word-review-conflict');
  await page
    .locator('.playground-template-card')
    .filter({ hasText: '项目方案' })
    .click();
  await expect(page.getByRole('textbox', { name: '文档正文' })).toContainText(
    '描述当前情况、核心问题和可衡量的成功标准。',
  );

  await page.getByRole('button', { name: '加载审阅冲突夹具' }).click();
  await expect(
    page.locator(
      "[data-comment-id='e2e-controlled-review-comment'][data-document-comment='true']",
    ),
  ).toBeVisible();
  await expect(
    page.locator(
      "[data-change-id='e2e-controlled-review-change'][data-document-change='true']",
    ),
  ).toBeVisible();

  await page.getByRole('button', { name: '模拟外部审阅更新' }).click();

  const warning = page.locator('.work-document-review-conflict[role="alert"]');
  await expect(warning).toContainText('其中 2 个批注或修订范围发生变化。');
  await expect(page.getByText('检测到 2 个审阅冲突')).toBeVisible();
  await expect(page.getByRole('textbox', { name: '文档正文' })).toContainText(
    '外部改写内容',
  );

  await page.getByRole('button', { name: '关闭审阅冲突提示' }).click();
  await expect(warning).toHaveCount(0);
  expect(consoleErrors).toEqual([]);
  expect(pageErrors).toEqual([]);
});
