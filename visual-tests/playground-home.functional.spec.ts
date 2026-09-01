import { expect, test } from '@playwright/test';

test('Playground home omits the latest-capabilities promotion and keeps feature templates', async ({
  page,
}) => {
  await page.goto('/playground/');

  await expect(
    page.getByRole('heading', { name: '我的文档', level: 1 }),
  ).toBeVisible();
  await expect(page.getByRole('region', { name: '最新能力' })).toHaveCount(0);
  await expect(page.locator('.playground-latest-capabilities')).toHaveCount(0);
  await expect(page.getByText('main 已部署', { exact: true })).toHaveCount(0);

  await expect(
    page.getByRole('region', { name: '多人实时协作' }),
  ).toBeVisible();
  const templates = page.getByRole('region', { name: '新建', exact: true });
  await expect(templates).toBeVisible();
  const animationTemplate = templates.locator(
    "button[data-template-id='animated-deck']",
  );
  await expect(animationTemplate).toBeVisible();
  await animationTemplate.click();

  await expect(page.getByRole('textbox', { name: '文件名' })).toHaveValue(
    '进入与退出动画示例',
  );
});
