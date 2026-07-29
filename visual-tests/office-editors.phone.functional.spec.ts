import { expect, test } from '@playwright/test';
import { openPdfFixture, waitForPdfFixture } from './pdf-test-support';

test.use({ viewport: { width: 390, height: 700 } });

test('Presentation prioritizes its canvas and uses a dismissible slide drawer', async ({
  page,
}) => {
  await page.goto('/');
  await page
    .getByRole('button', { name: '业务策略汇报 PPTX · 本次会话' })
    .click();
  const canvas = page.locator('.work-slide-canvas.interactive');
  await canvas.waitFor();

  const toggle = page.getByRole('button', { name: '打开幻灯片导航' });
  const rail = page.getByRole('complementary', { name: '幻灯片' });
  await expect(toggle).toBeVisible();
  await expect(rail).toBeHidden();
  await expect
    .poll(async () => (await canvas.boundingBox())?.width ?? 0)
    .toBeGreaterThan(300);

  await toggle.click();
  await expect(rail).toBeVisible();
  await expect
    .poll(() =>
      rail.evaluate((element) => element.getBoundingClientRect().left),
    )
    .toBeGreaterThanOrEqual(-1);
  const drawerGeometry = await rail.evaluate((element) => {
    const rect = element.getBoundingClientRect();
    return {
      left: rect.left,
      right: rect.right,
      width: rect.width,
    };
  });
  expect(drawerGeometry.left).toBeGreaterThanOrEqual(-1);
  expect(drawerGeometry.right).toBeLessThanOrEqual(281);
  expect(drawerGeometry.width).toBeGreaterThanOrEqual(260);
  const close = page.getByRole('button', {
    name: '关闭幻灯片导航',
    exact: true,
  });
  await expect(close).toBeFocused();
  await page.getByRole('button', { name: '幻灯片 2 / 3：核心判断' }).click();
  await expect(rail).toBeHidden();
  await expect(toggle).toBeFocused();
});

test('PDF keeps compact tools clear of the file actions on phones', async ({
  page,
}) => {
  await page.goto('/');
  await openPdfFixture(page);
  await waitForPdfFixture(page);

  const toolbar = page.getByRole('toolbar', { name: 'PDF 工具栏' });
  const more = toolbar.getByRole('button', { name: '更多 PDF 工具' });
  const ai = page.getByRole('button', { name: '打开 AI 助手' });
  const download = page.getByRole('button', { name: '下载 PDF' });
  await expect(more).toBeVisible();
  await expect(ai).toBeVisible();
  await expect(download).toBeVisible();
  await expect(toolbar.getByRole('button', { name: '画笔' })).toBeHidden();
  await expect(toolbar.getByRole('button', { name: '批注样式' })).toBeHidden();

  const geometry = await page.evaluate(() => {
    const more = document.querySelector<HTMLElement>(
      '[aria-label="更多 PDF 工具"]',
    );
    const ai = document.querySelector<HTMLElement>(
      '[aria-label="打开 AI 助手"]',
    );
    const download = document.querySelector<HTMLElement>(
      '[aria-label="下载 PDF"]',
    );
    if (!(more && ai && download)) {
      throw new Error('Phone PDF command geometry is incomplete.');
    }
    const moreRect = more.getBoundingClientRect();
    const actionLeft = Math.min(
      ai.getBoundingClientRect().left,
      download.getBoundingClientRect().left,
    );
    return {
      actionLeft,
      moreLeft: moreRect.left,
      moreRight: moreRect.right,
      viewportWidth: document.documentElement.clientWidth,
    };
  });
  expect(geometry.moreLeft).toBeGreaterThanOrEqual(0);
  expect(geometry.moreRight).toBeLessThanOrEqual(geometry.actionLeft);
  expect(geometry.moreRight).toBeLessThanOrEqual(geometry.viewportWidth);

  await toolbar.getByRole('button', { name: '高亮' }).click();
  await more.click();
  let menu = page.getByRole('menu', { name: '更多 PDF 工具' });
  await expect(menu.getByRole('menuitemradio', { name: '选择' })).toBeVisible();
  await expect(menu.getByRole('menuitemradio', { name: '画笔' })).toBeVisible();
  await expect(
    menu.getByRole('menuitemradio', { name: '透明度 50%' }),
  ).toBeVisible();
  await menu.getByRole('menuitemradio', { name: '画笔' }).click();
  await more.click();
  menu = page.getByRole('menu', { name: '更多 PDF 工具' });
  await expect(
    menu.getByRole('menuitemradio', { name: '线宽 4' }),
  ).toBeVisible();
});
