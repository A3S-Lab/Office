import { expect, type Page, test } from '@playwright/test';

test('Presentation phone chart pane traps focus and restores chart focus', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });
  page.on('pageerror', (error) => browserErrors.push(error.message));

  await page.setViewportSize({ width: 390, height: 700 });
  await openPresentationFixture(page);
  await page.getByRole('tab', { name: '插入', exact: true }).click();
  await page.getByRole('button', { name: '图表', exact: true }).click();

  const pane = page.getByRole('dialog', {
    name: '演示图表数据',
    exact: true,
  });
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
  await expect(
    pane.getByRole('button', { name: '关闭演示图表数据' }),
  ).toBeFocused();

  const values = pane.getByRole('textbox', {
    name: '演示图表系列 1 数据',
  });
  await values.fill('32, wrong, 61');
  await expect(values).toHaveAttribute('aria-invalid', 'true');
  await values.press('Escape');
  await expect(values).toHaveValue('32, 48, 61');
  await expect(pane).toBeVisible();

  await page.screenshot({
    path: testInfo.outputPath('presentation-phone-chart-pane.png'),
    animations: 'disabled',
  });
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
