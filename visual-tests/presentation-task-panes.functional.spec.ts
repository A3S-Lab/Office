import { expect, test, type Page } from '@playwright/test';

test('Presentation replaces task panes and closes them from the toolbar with Escape', async ({
  page,
}) => {
  await openPresentationFixture(page);

  await page.getByRole('tab', { name: '设计', exact: true }).click();
  await page.getByRole('button', { name: '母版和版式' }).click();
  const designPane = page.getByRole('region', { name: '母版与布局' });
  await expect(designPane).toBeVisible();

  await page.getByRole('tab', { name: '审阅', exact: true }).click();
  await page.getByRole('button', { name: '查看批注' }).click();
  await expect(designPane).toBeHidden();
  const commentsPane = page.getByRole('region', { name: '演示批注审阅' });
  await expect(commentsPane).toBeVisible();

  await page.keyboard.press('Escape');
  await expect(commentsPane).toBeHidden();

  await page.getByRole('tab', { name: '设计', exact: true }).click();
  await page.getByRole('button', { name: '母版和版式' }).click();
  await expect(designPane).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(designPane).toBeHidden();
});

test('Presentation keeps empty comments in the dialog until they are valid', async ({
  page,
}) => {
  await openPresentationFixture(page);

  await page.getByRole('tab', { name: '审阅', exact: true }).click();
  await page.getByRole('button', { name: '新建批注' }).click();
  const dialog = page.getByRole('dialog', { name: '批注内容' });
  const input = dialog.getByRole('textbox', { name: '批注内容' });
  const submit = dialog.getByRole('button', { name: '添加批注' });

  await expect(submit).toBeDisabled();
  await input.fill('   ');
  await expect(submit).toBeDisabled();
  await expect(dialog.getByRole('alert')).toHaveText('请输入批注内容。');

  await input.fill('需要补充数据来源');
  await expect(submit).toBeEnabled();
  await page.keyboard.press('Escape');
  await expect(dialog).toBeHidden();
});

test('Presentation closes the contextual chart pane with Escape', async ({
  page,
}) => {
  await openPresentationFixture(page);

  await page.getByRole('tab', { name: '插入', exact: true }).click();
  await page.getByRole('button', { name: '图表', exact: true }).click();
  const chartPane = page.locator('.work-presentation-chart-panel');
  await expect(chartPane).toBeVisible();

  const values = page.getByRole('textbox', {
    name: '演示图表系列 1 数据',
  });
  await values.fill('32, wrong, 61');
  await expect(values).toHaveAttribute('aria-invalid', 'true');
  await page.keyboard.press('Escape');
  await expect(values).toHaveValue('32, 48, 61');
  await expect(chartPane).toBeVisible();

  const minimum = chartPane.getByRole('textbox', {
    name: '演示图表纵轴最小值',
  });
  const maximum = chartPane.getByRole('textbox', {
    name: '演示图表纵轴最大值',
  });
  await minimum.fill('10');
  await minimum.press('Enter');
  await maximum.fill('5');
  await expect(maximum).toHaveAttribute('aria-invalid', 'true');
  await maximum.press('Enter');
  await expect(maximum).toHaveValue('');
  await expect(maximum).not.toHaveAttribute('aria-invalid');
  await expect(chartPane).toBeVisible();

  await page.keyboard.press('Escape');
  await expect(chartPane).toBeHidden();
});

test('Presentation contains the chart inspector and preserves selection when it closes', async ({
  page,
}, testInfo) => {
  await openPresentationFixture(page);

  await page.getByRole('tab', { name: '插入', exact: true }).click();
  await page.getByRole('button', { name: '图表', exact: true }).click();

  const workspace = page.locator('.work-presentation-workspace');
  const chartPane = page.locator('.work-presentation-chart-panel');
  const selectedChart = page.locator('.work-slide-element.chart.selected');
  await expect(workspace).toBeVisible();
  await expect(chartPane).toBeVisible();
  await expect(selectedChart).toHaveCount(1);
  if (testInfo.project.name === 'desktop-1280') {
    await expect(chartPane).not.toHaveAttribute('aria-modal', 'true');
    await expect(page.locator('.work-presentation-ribbon')).not.toHaveAttribute(
      'inert',
      '',
    );
  } else {
    await expect(chartPane).toHaveAttribute('aria-modal', 'true');
    await expect(page.locator('.work-presentation-ribbon')).toHaveAttribute(
      'inert',
      '',
    );
  }

  const geometry = await page.evaluate(() => {
    const panel = document.querySelector<HTMLElement>(
      '.work-presentation-chart-panel',
    );
    const body = document.querySelector<HTMLElement>(
      '.work-presentation-chart-controls',
    );
    const canvas = document.querySelector<HTMLElement>(
      '.work-slide-canvas.interactive',
    );
    if (!panel || !body || !canvas) {
      throw new Error('Presentation chart inspector geometry is incomplete.');
    }
    const panelBox = panel.getBoundingClientRect();
    const canvasBox = canvas.getBoundingClientRect();
    return {
      bodyClientWidth: body.clientWidth,
      bodyScrollWidth: body.scrollWidth,
      canvasLeft: canvasBox.left,
      panelLeft: panelBox.left,
      panelWidth: panelBox.width,
    };
  });

  expect(geometry.panelWidth).toBeGreaterThanOrEqual(340);
  expect(geometry.panelLeft).toBeGreaterThan(geometry.canvasLeft);
  expect(geometry.bodyScrollWidth).toBeLessThanOrEqual(
    geometry.bodyClientWidth + 1,
  );

  await chartPane.getByRole('button', { name: '关闭演示图表数据' }).click();
  await expect(chartPane).toBeHidden();
  await expect(selectedChart).toHaveCount(1);
  await expect(page.getByRole('region', { name: '演示编辑器' })).toBeVisible();
  await expect(
    page.getByRole('button', { name: '返回办公首页' }),
  ).toBeVisible();
});

test('Presentation task-pane text fields cancel a draft before the pane closes', async ({
  page,
}) => {
  await openPresentationFixture(page);

  await page.getByRole('tab', { name: '插入', exact: true }).click();
  await page.getByRole('button', { name: '图表', exact: true }).click();
  const chartPane = page.locator('.work-presentation-chart-panel');
  const title = chartPane.getByRole('textbox', { name: '演示图表标题' });
  const originalTitle = await title.inputValue();

  await title.fill('未保存的标题');
  await title.press('Escape');
  await expect(title).toHaveValue(originalTitle);
  await expect(chartPane).toBeVisible();

  await title.fill('已保存的标题');
  await title.press('Enter');
  await expect(title).toHaveValue('已保存的标题');
  await expect(chartPane).toBeVisible();

  await title.press('Escape');
  await expect(chartPane).toBeHidden();
});

test('Presentation contains the phone chart pane and restores chart focus', async ({
  page,
}, testInfo) => {
  test.skip(
    testInfo.project.name !== 'desktop-1280',
    'The phone contract only needs one browser project.',
  );
  await page.setViewportSize({ width: 390, height: 700 });
  await openPresentationFixture(page);

  await page.getByRole('tab', { name: '插入', exact: true }).click();
  await page.getByRole('button', { name: '图表', exact: true }).click();

  const pane = page.getByRole('dialog', {
    name: '演示图表数据',
    exact: true,
  });
  const close = pane.getByRole('button', { name: '关闭演示图表数据' });
  const removeChart = pane.getByRole('button', {
    name: '删除演示图表',
    exact: true,
  });
  const addSeries = pane.getByRole('button', { name: '添加图表系列' });
  const selectedChart = page.locator('.work-slide-element.chart.selected');

  await expect(pane).toHaveAttribute('aria-modal', 'true');
  await expect(page.locator('.work-presentation-ribbon')).toHaveAttribute(
    'inert',
    '',
  );
  await expect(page.locator('.work-presentation-layout')).toHaveAttribute(
    'inert',
    '',
  );
  await expect(page.locator('.work-presentation-status')).toHaveAttribute(
    'inert',
    '',
  );
  await expect(close).toBeFocused();

  await removeChart.focus();
  await removeChart.press('Shift+Tab');
  await expect(addSeries).toBeFocused();
  await addSeries.press('Tab');
  await expect(removeChart).toBeFocused();

  const values = pane.getByRole('textbox', {
    name: '演示图表系列 1 数据',
  });
  await values.fill('32, wrong, 61');
  await expect(values).toHaveAttribute('aria-invalid', 'true');
  await values.press('Escape');
  await expect(values).toHaveValue('32, 48, 61');
  await expect(pane).toBeVisible();

  await values.press('Escape');
  await expect(pane).toBeHidden();
  await expect(selectedChart).toBeFocused();
  await expect(page.locator('.work-presentation-ribbon')).not.toHaveAttribute(
    'inert',
    '',
  );
  await expect(page.locator('.work-presentation-layout')).not.toHaveAttribute(
    'inert',
    '',
  );
  await expect(page.locator('.work-presentation-status')).not.toHaveAttribute(
    'inert',
    '',
  );
});

test('Presentation applies the live transition timing draft to every slide', async ({
  page,
}) => {
  await openPresentationFixture(page);

  const thumbnails = page.locator('.work-slide-strip [data-slide-thumbnail]');
  const activeThumbnail = page.locator(
    '.work-slide-strip [data-slide-thumbnail].active',
  );
  await page.getByRole('tab', { name: '切换', exact: true }).click();
  const effect = page.getByRole('combobox', {
    name: '幻灯片切换效果',
  });
  const apply = page.getByRole('button', {
    name: '应用切换效果到全部幻灯片',
  });
  await expect(effect).toHaveText('无');
  await expect(apply).toBeDisabled();

  await effect.click();
  await page.getByRole('option', { name: '淡化' }).click();
  await expect(apply).toBeEnabled();
  const automatic = page.getByRole('checkbox', { name: '自动换片' });
  await automatic.click();
  const seconds = page.getByRole('textbox', { name: '自动换片秒数' });
  await seconds.fill('7.75');
  await expect(seconds).toHaveValue('7.75');

  await apply.click();
  await expect(apply).toBeDisabled();
  await expect(activeThumbnail).toBeFocused();

  await thumbnails.nth(1).click();
  await expect(seconds).toHaveValue('7.75');
  await expect(apply).toBeDisabled();
  const speed = page.getByRole('combobox', { name: '切换速度' });
  await speed.click();
  await page.getByRole('option', { name: '慢速' }).click();
  await expect(apply).toBeEnabled();
  await expect(activeThumbnail).toBeFocused();

  await thumbnails.first().click();
  await expect(seconds).toHaveValue('7.75');
  await expect(speed).toHaveText('中速');
  await expect(apply).toBeEnabled();
});

test('Presentation disables slide transitions while editing a layout', async ({
  page,
}) => {
  await openPresentationFixture(page);

  await page.getByRole('tab', { name: '设计', exact: true }).click();
  await page.getByRole('button', { name: '母版和版式' }).click();
  await page.getByRole('button', { name: '编辑当前布局' }).click();
  await page.getByRole('tab', { name: '切换', exact: true }).click();

  await expect(
    page.getByRole('combobox', { name: '幻灯片切换效果' }),
  ).toBeDisabled();
  await expect(page.getByRole('combobox', { name: '切换速度' })).toBeDisabled();
  await expect(
    page.getByRole('button', {
      name: '应用切换效果到全部幻灯片',
    }),
  ).toBeDisabled();
});

async function openPresentationFixture(page: Page) {
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
