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

  const status = page.locator('.work-presentation-status');
  const slideStatus = status.getByLabel('幻灯片状态');
  await expect(slideStatus).toHaveText('幻灯片 1 / 3');
  await expect(slideStatus).toBeVisible();
  await expect(status.getByLabel('演示备注状态')).toBeHidden();
  await expect(status.getByLabel('演示批注状态')).toBeHidden();
  await expect(status.getByLabel('演示保存状态')).toBeHidden();
  await expect
    .poll(() =>
      slideStatus.evaluate(
        (output) => output.scrollWidth <= output.clientWidth + 1,
      ),
    )
    .toBe(true);
  const stageGeometry = await page.evaluate(() => {
    const stage = document.querySelector<HTMLElement>('.work-slide-stage');
    const canvas = stage?.querySelector<HTMLElement>(
      ':scope > .work-slide-canvas',
    );
    if (!(stage && canvas)) {
      throw new Error('Phone Presentation stage geometry is unavailable.');
    }
    const stageBounds = stage.getBoundingClientRect();
    const canvasBounds = canvas.getBoundingClientRect();
    return {
      canvasTop: canvasBounds.top,
      stageTop: stageBounds.top,
    };
  });
  expect(stageGeometry.canvasTop - stageGeometry.stageTop).toBeLessThanOrEqual(
    72,
  );

  const toggle = page.getByRole('button', { name: '打开幻灯片导航' });
  const rail = page.locator('.work-slide-strip');
  await expect(toggle).toBeVisible();
  await expect(rail).toBeHidden();
  await expect
    .poll(async () => (await canvas.boundingBox())?.width ?? 0)
    .toBeGreaterThan(300);

  await toggle.click();
  await expect(rail).toBeVisible();
  await expect(rail).toHaveAttribute('role', 'dialog');
  await expect(rail).toHaveAttribute('aria-modal', 'true');
  await expect(toggle).toHaveAttribute('inert', '');
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
  await page.keyboard.press('Tab');
  const firstSlide = page.getByRole('button', {
    name: '幻灯片 1 / 3：封面',
  });
  await expect(firstSlide).toBeFocused();
  await page.keyboard.press('Shift+Tab');
  await expect(close).toBeFocused();
  await page.getByRole('button', { name: '幻灯片 2 / 3：核心判断' }).click();
  await expect(rail).toBeHidden();
  await expect(toggle).toBeFocused();
});

test('Phone Office sidebar owns focus until dismissed', async ({ page }) => {
  await page.goto('/');
  const trigger = page.getByRole('button', { name: '展开办公侧边栏' });
  await trigger.click();

  const sidebar = page.getByRole('dialog', { name: 'A3S Office 导航' });
  const close = page.getByRole('button', { name: '收起办公侧边栏' });
  await expect(sidebar).toHaveAttribute('aria-modal', 'true');
  await expect(page.locator('.playground-main-pane')).toHaveAttribute(
    'inert',
    '',
  );
  await expect(close).toBeFocused();

  await page.keyboard.press('Tab');
  await expect(sidebar.getByRole('button', { name: '编辑器' })).toBeFocused();
  await page.keyboard.press('Shift+Tab');
  await expect(close).toBeFocused();
  await page.keyboard.press('Escape');

  await expect(sidebar).toBeHidden();
  await expect(trigger).toBeFocused();
  await expect(page.locator('.playground-main-pane')).not.toHaveAttribute(
    'inert',
    '',
  );
});

test('Phone AI assistant keeps focus out of the obscured editor', async ({
  page,
}) => {
  await page.goto('/');
  await page
    .getByRole('button', { name: '新项目方案 DOCX · 本次会话' })
    .click();

  const trigger = page.getByRole('button', { name: '打开 AI 助手' });
  await trigger.click();
  const assistant = page.getByRole('dialog', { name: 'AI 助手' });
  const close = assistant.getByRole('button', { name: '关闭 AI 助手' });
  await expect(assistant).toHaveAttribute('aria-modal', 'true');
  await expect(page.locator('.playground-editor-host')).toHaveAttribute(
    'inert',
    '',
  );
  await expect(close).toBeFocused();

  await page.keyboard.press('Tab');
  await expect(close).toBeFocused();
  await page.keyboard.press('Shift+Tab');
  await expect(close).toBeFocused();
  await page.keyboard.press('Escape');

  await expect(assistant).toBeHidden();
  await expect(trigger).toBeFocused();
  await expect(page.locator('.playground-editor-host')).not.toHaveAttribute(
    'inert',
    '',
  );
});

test('Shared Office color picker exposes phone-sized touch targets', async ({
  page,
}) => {
  await page.setViewportSize({ width: 320, height: 568 });
  await page.goto('/');
  await page
    .getByRole('button', { name: '新项目方案 DOCX · 本次会话' })
    .click();

  const trigger = page.getByRole('button', { name: '文字颜色' });
  await trigger.click();
  const menu = page.getByRole('dialog', { name: '文字颜色' });
  const first = menu.getByRole('option', { name: '颜色 #111827' });
  const secondRow = menu.getByRole('option', { name: '颜色 #f9fafb' });
  await expect(first).toBeFocused();

  const geometry = await page.evaluate(() => {
    const menu = document.querySelector<HTMLElement>('.work-office-color-menu');
    const grid = document.querySelector<HTMLElement>(
      '.work-office-color-grid.theme',
    );
    const option = grid?.querySelector<HTMLElement>('[role="option"]');
    if (!(menu && grid && option)) {
      throw new Error('Phone Office color picker is unavailable.');
    }
    const menuBounds = menu.getBoundingClientRect();
    const optionBounds = option.getBoundingClientRect();
    return {
      columns: getComputedStyle(grid).gridTemplateColumns.split(' ').length,
      menuLeft: menuBounds.left,
      menuRight: menuBounds.right,
      optionHeight: optionBounds.height,
      optionWidth: optionBounds.width,
      viewportWidth: document.documentElement.clientWidth,
    };
  });
  expect(geometry.columns).toBe(8);
  expect(geometry.optionWidth).toBeGreaterThanOrEqual(30);
  expect(geometry.optionHeight).toBeGreaterThanOrEqual(30);
  expect(geometry.menuLeft).toBeGreaterThanOrEqual(0);
  expect(geometry.menuRight).toBeLessThanOrEqual(geometry.viewportWidth);

  await first.press('ArrowDown');
  await expect(secondRow).toBeFocused();
  await secondRow.click();
  await expect(menu).toBeHidden();
  await expect(trigger.locator('.work-office-color-swatch')).toHaveCSS(
    'background-color',
    'rgb(249, 250, 251)',
  );
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
