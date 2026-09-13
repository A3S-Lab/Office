import { expect, type Page, test } from '@playwright/test';

test('Presentation ribbon keeps OfficeSelect controls on the 74px row', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));

  await openPresentationFixture(page);

  const ribbon = page.locator('.work-presentation-ribbon');
  const toolbar = ribbon.locator('.presentation-toolbar');
  await expect(toolbar).toBeVisible();

  const toolbarBox = await toolbar.boundingBox();
  expect(toolbarBox?.height ?? 0).toBeGreaterThanOrEqual(70);
  expect(toolbarBox?.height ?? 0).toBeLessThanOrEqual(86);

  const title = page
    .locator('.work-slide-canvas.interactive > .work-slide-element')
    .nth(1);
  await title.click();

  const fontSelect = ribbon.locator('.presentation-font-family-select');
  const alignSelect = ribbon.locator('.presentation-align-select');
  await expect(fontSelect).toBeVisible();
  await expect(alignSelect).toBeVisible();

  const fontBox = await fontSelect.boundingBox();
  const alignBox = await alignSelect.boundingBox();
  expect(fontBox?.width ?? 0).toBeGreaterThanOrEqual(110);
  expect(alignBox?.width ?? 0).toBeGreaterThanOrEqual(100);

  const fontTrigger = fontSelect.getByRole('combobox', { name: '演示字体' });
  await fontTrigger.click();
  const fontMenu = page.getByRole('listbox', { name: '演示字体' });
  await expect(fontMenu).toBeVisible();
  const menuBox = await fontMenu.boundingBox();
  expect(menuBox?.width ?? 0).toBeGreaterThanOrEqual(200);
  expect(menuBox?.y ?? 0).toBeGreaterThan((fontBox?.y ?? 0) + 10);

  await page.screenshot({
    path: testInfo.outputPath(
      `presentation-ribbon-select-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
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
  await expect(page.locator('.work-presentation-editor')).toBeVisible();
}
