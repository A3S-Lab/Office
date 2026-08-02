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

  const toggle = page.locator('.work-presentation-slide-navigation-toggle');
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

  await page.setViewportSize({ width: 1280, height: 800 });
  await expect(rail).toBeVisible();
  await expect(rail).not.toHaveAttribute('role', 'dialog');
  await expect(rail).not.toHaveAttribute('aria-modal', 'true');
  await expect(toggle).not.toHaveAttribute('inert', '');
  await expect(page.locator('.work-slide-stage')).not.toHaveAttribute(
    'inert',
    '',
  );
  await expect(firstSlide).toBeFocused();

  await page.setViewportSize({ width: 390, height: 700 });
  await expect(rail).toHaveAttribute('role', 'dialog');
  await expect(rail).toHaveAttribute('aria-modal', 'true');
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
  await expect(
    sidebar.getByRole('button', { name: 'Playground' }),
  ).toBeFocused();
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

test('Phone Word comments own focus until the review drawer closes', async ({
  page,
}) => {
  await page.setViewportSize({ width: 390, height: 700 });
  await page.goto('/');
  await page
    .getByRole('button', { name: '新项目方案 DOCX · 本次会话' })
    .click();

  await page.getByRole('tab', { name: '审阅' }).click();
  const trigger = page.getByRole('button', {
    name: '查看批注',
    exact: true,
  });
  await trigger.click();

  const drawer = page.getByRole('dialog', { name: '批注审阅' });
  const close = drawer.getByRole('button', { name: '关闭批注审阅' });
  await expect(drawer).toHaveAttribute('aria-modal', 'true');
  await expect(page.locator('.work-document-page-stage')).toHaveAttribute(
    'inert',
    '',
  );
  await expect(page.locator('.work-document-ribbon')).toHaveAttribute(
    'inert',
    '',
  );
  await expect(close).toBeFocused();

  await page.keyboard.press('Tab');
  await expect(close).toBeFocused();
  await page.keyboard.press('Shift+Tab');
  await expect(close).toBeFocused();
  await page.keyboard.press('Escape');

  await expect(drawer).toBeHidden();
  await expect(trigger).toBeFocused();
  await expect(page.locator('.work-document-page-stage')).not.toHaveAttribute(
    'inert',
    '',
  );
  await expect(page.locator('.work-document-ribbon')).not.toHaveAttribute(
    'inert',
    '',
  );

  await page.setViewportSize({ width: 1280, height: 800 });
  await trigger.click();
  const desktopPane = page.locator('.work-document-comments-panel');
  await expect(desktopPane).toBeVisible();
  await expect(desktopPane).not.toHaveAttribute('role', 'dialog');
  await expect(desktopPane).not.toHaveAttribute('aria-modal', 'true');
  await expect(page.locator('.work-document-page-stage')).not.toHaveAttribute(
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

test('Phone Spreadsheet find keeps its input and actions touch-sized', async ({
  page,
}) => {
  await page.goto('/');
  await page
    .getByRole('button', { name: '季度执行计划 XLSX · 本次会话' })
    .click();

  const grid = page.locator('.fortune-sheet-overlay');
  await grid.waitFor();
  await grid.focus();
  await page.keyboard.press('Control+f');

  const query = page.getByRole('textbox', { name: '查找当前工作表' });
  await expect(query).toBeFocused();
  const geometry = await page
    .locator('.work-spreadsheet-find-bar')
    .evaluate((bar) => {
      const input = bar.querySelector<HTMLElement>(
        '[aria-label="查找当前工作表"]',
      );
      const buttons = Array.from(
        bar.querySelectorAll<HTMLElement>(
          '.work-spreadsheet-find-actions button',
        ),
      );
      if (!input || buttons.length !== 3) {
        throw new Error('Phone Spreadsheet Find controls are incomplete.');
      }
      const barBounds = bar.getBoundingClientRect();
      const inputBounds = input.getBoundingClientRect();
      return {
        barLeft: barBounds.left,
        barRight: barBounds.right,
        buttonBounds: buttons.map((button) => {
          const bounds = button.getBoundingClientRect();
          return { height: bounds.height, width: bounds.width };
        }),
        inputHeight: inputBounds.height,
        inputWidth: inputBounds.width,
        viewportWidth: document.documentElement.clientWidth,
      };
    });

  expect(geometry.barLeft).toBeGreaterThanOrEqual(0);
  expect(geometry.barRight).toBeLessThanOrEqual(geometry.viewportWidth);
  expect(geometry.inputHeight).toBeGreaterThanOrEqual(40);
  expect(geometry.inputWidth).toBeGreaterThanOrEqual(96);
  for (const bounds of geometry.buttonBounds) {
    expect(bounds.height).toBeGreaterThanOrEqual(40);
    expect(bounds.width).toBeGreaterThanOrEqual(40);
  }

  await query.fill('客户洞察报告');
  await expect(page.getByText('1 个匹配', { exact: true })).toBeVisible();
  await query.press('Enter');
  await expect(page.locator('.fortune-name-box')).toHaveText('A4');
  await query.press('Escape');
  await expect(query).toHaveCount(0);
  await expect(grid).toBeFocused();
});

test('Word page color keeps its custom controls visible at phone height', async ({
  page,
}) => {
  await page.setViewportSize({ width: 390, height: 700 });
  await page.goto('/');
  await page
    .getByRole('button', { name: '新项目方案 DOCX · 本次会话' })
    .click();

  await page.getByRole('tab', { name: '页面布局' }).click();
  await page.getByRole('button', { name: '页面颜色' }).click();

  const menu = page.getByRole('dialog', { name: '页面颜色' });
  await expect(
    menu.getByRole('textbox', { name: '自定义颜色值' }),
  ).toBeVisible();
  await expect(
    menu.getByRole('button', { name: '应用自定义颜色' }),
  ).toBeVisible();

  const geometry = await menu.evaluate((panel) => {
    const custom = panel.querySelector<HTMLElement>(
      '.work-office-color-custom',
    );
    if (!custom) {
      throw new Error('Custom color controls are unavailable.');
    }
    const panelBounds = panel.getBoundingClientRect();
    const customBounds = custom.getBoundingClientRect();
    return {
      customBottom: customBounds.bottom,
      customTop: customBounds.top,
      panelBottom: panelBounds.bottom,
      panelClientHeight: panel.clientHeight,
      panelScrollHeight: panel.scrollHeight,
      panelScrollTop: panel.scrollTop,
      panelTop: panelBounds.top,
      viewportHeight: document.documentElement.clientHeight,
    };
  });
  expect(geometry.panelScrollTop).toBe(0);
  expect(geometry.customTop).toBeGreaterThanOrEqual(geometry.panelTop);
  expect(geometry.customBottom).toBeLessThanOrEqual(geometry.panelBottom + 1);
  expect(geometry.customBottom).toBeLessThanOrEqual(geometry.viewportHeight);
  expect(geometry.panelScrollHeight).toBeLessThanOrEqual(
    geometry.panelClientHeight + 1,
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
  const search = toolbar.getByRole('searchbox', { name: '在 PDF 中搜索' });
  await expect(more).toBeVisible();
  await expect(ai).toBeVisible();
  await expect(download).toBeVisible();
  await expect(search).toBeVisible();
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
    const search = document.querySelector<HTMLElement>('.work-pdf-search');
    const toolbar = document.querySelector<HTMLElement>('.work-pdf-toolbar');
    if (!(more && ai && download && search && toolbar)) {
      throw new Error('Phone PDF command geometry is incomplete.');
    }
    const moreRect = more.getBoundingClientRect();
    const searchRect = search.getBoundingClientRect();
    const toolbarRect = toolbar.getBoundingClientRect();
    const actionLeft = Math.min(
      ai.getBoundingClientRect().left,
      download.getBoundingClientRect().left,
    );
    const actionBottom = Math.max(
      ai.getBoundingClientRect().bottom,
      download.getBoundingClientRect().bottom,
    );
    return {
      actionBottom,
      actionLeft,
      moreLeft: moreRect.left,
      moreRight: moreRect.right,
      searchBottom: searchRect.bottom,
      searchTop: searchRect.top,
      toolbarBottom: toolbarRect.bottom,
      toolbarHeight: toolbarRect.height,
      viewportWidth: document.documentElement.clientWidth,
    };
  });
  expect(geometry.moreLeft).toBeGreaterThanOrEqual(0);
  expect(geometry.moreRight).toBeLessThanOrEqual(geometry.actionLeft);
  expect(geometry.moreRight).toBeLessThanOrEqual(geometry.viewportWidth);
  expect(geometry.toolbarHeight).toBeGreaterThanOrEqual(70);
  expect(geometry.searchTop).toBeGreaterThanOrEqual(geometry.actionBottom);
  expect(geometry.searchBottom).toBeLessThanOrEqual(geometry.toolbarBottom);

  await search.fill('A3S');
  await toolbar.getByRole('button', { name: '清除搜索' }).click();
  await expect(search).toHaveValue('');
  await expect(page.getByText(/\.pdf 已下载$/)).toHaveCount(0);

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

test('PDF uses a dismissible page drawer on phones', async ({ page }) => {
  await page.goto('/');
  await openPdfFixture(page, { pageCount: 3 });
  await waitForPdfFixture(page);

  const toggle = page.getByRole('button', { name: '打开 PDF 页面导航' });
  const rail = page.locator('.work-pdf-thumbnail-rail');
  await expect(toggle).toBeVisible();
  await expect(rail).toBeHidden();
  const triggerPlacement = await toggle.evaluate((element) => {
    const toolbar = element.closest('.work-pdf-toolbar');
    const embed = document.querySelector('.work-pdf-embed');
    const toggleBounds = element.getBoundingClientRect();
    const toolbarBounds = toolbar?.getBoundingClientRect();
    const embedBounds = embed?.getBoundingClientRect();
    return {
      insideToolbar: Boolean(toolbar),
      toggleBottom: toggleBounds.bottom,
      toolbarBottom: toolbarBounds?.bottom ?? Number.NaN,
      embedTop: embedBounds?.top ?? Number.NaN,
    };
  });
  expect(triggerPlacement.insideToolbar).toBe(true);
  expect(triggerPlacement.toggleBottom).toBeLessThanOrEqual(
    triggerPlacement.toolbarBottom + 0.5,
  );
  expect(triggerPlacement.toggleBottom).toBeLessThanOrEqual(
    triggerPlacement.embedTop + 0.5,
  );

  await toggle.click();
  const dialog = page.getByRole('dialog', { name: 'PDF 页面' });
  const close = page.getByRole('button', {
    name: '关闭 PDF 页面导航',
    exact: true,
  });
  await expect(dialog).toBeVisible();
  await expect(dialog).toHaveAttribute('aria-modal', 'true');
  await expect(close).toBeFocused();
  await expect(page.locator('.work-pdf-embed')).toHaveAttribute('inert', '');
  await expect
    .poll(() =>
      dialog.evaluate((element) => element.getBoundingClientRect().left),
    )
    .toBeGreaterThanOrEqual(-1);
  const drawerGeometry = await dialog.evaluate((element) => {
    const bounds = element.getBoundingClientRect();
    return { left: bounds.left, right: bounds.right, width: bounds.width };
  });
  expect(drawerGeometry.left).toBeGreaterThanOrEqual(-1);
  expect(drawerGeometry.right).toBeLessThanOrEqual(281);
  expect(drawerGeometry.width).toBeGreaterThanOrEqual(260);

  await page.getByRole('button', { name: '第 2 页' }).click();
  await expect(dialog).toBeHidden();
  await expect(toggle).toBeFocused();
  await expect(page.getByRole('textbox', { name: '页码' })).toHaveValue('2');
  await expect(toggle).toContainText('第 2 页');
});
