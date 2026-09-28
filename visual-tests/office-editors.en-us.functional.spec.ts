import { expect, type Page, test } from '@playwright/test';
import { openPdfFixture, waitForPdfFixture } from './pdf-test-support';
import {
  openDocumentFixture,
  waitForDocumentFixture,
} from './visual-test-support';

/**
 * Phase 3.1 en-US visual pass: open each editor with `?locale=en-US` and prove
 * catalog-backed chrome is English (no hard-coded zh-CN ribbon labels).
 */
async function gotoPlaygroundEnUS(page: Page): Promise<void> {
  await page.goto('/playground/?locale=en-US');
}

test.describe('en-US editor chrome', () => {
  test('Document ribbon tabs resolve through the en-US catalog', async ({
    page,
  }) => {
    await gotoPlaygroundEnUS(page);
    await openDocumentFixture(page);
    await waitForDocumentFixture(page);

    const ribbon = page.locator('.work-document-ribbon');
    await expect(ribbon.getByRole('tab', { name: 'Home' })).toHaveAttribute(
      'aria-selected',
      'true',
    );
    await expect(ribbon.getByRole('tab', { name: 'Insert' })).toBeVisible();
    await expect(ribbon.getByRole('tab', { name: 'Review' })).toBeVisible();
    await expect(ribbon.getByRole('tab', { name: '开始' })).toHaveCount(0);
  });

  test('Spreadsheet ribbon tabs resolve through the en-US catalog', async ({
    page,
  }) => {
    await gotoPlaygroundEnUS(page);
    await page
      .getByRole('button', { name: '季度执行计划 XLSX · 本次会话' })
      .click();
    await page.locator('.work-spreadsheet-canvas > .fortune-container').waitFor();

    const ribbon = page.locator('.work-spreadsheet-ribbon');
    await expect(ribbon.getByRole('tab', { name: 'Home' })).toHaveAttribute(
      'aria-selected',
      'true',
    );
    await expect(ribbon.getByRole('tab', { name: 'Insert' })).toBeVisible();
    await expect(ribbon.getByRole('tab', { name: 'Data' })).toBeVisible();
    await expect(ribbon.getByRole('tab', { name: 'Formulas' })).toBeVisible();
    await expect(ribbon.getByRole('tab', { name: '开始' })).toHaveCount(0);
  });

  test('Presentation ribbon tabs resolve through the en-US catalog', async ({
    page,
  }) => {
    await gotoPlaygroundEnUS(page);
    await page
      .getByRole('button', { name: '业务策略汇报 PPTX · 本次会话' })
      .click();
    await page.locator('.work-slide-canvas.interactive').waitFor();

    const ribbon = page.locator('.work-presentation-ribbon');
    await expect(ribbon.getByRole('tab', { name: 'Home' })).toBeVisible();
    await expect(ribbon.getByRole('tab', { name: 'Insert' })).toBeVisible();
    await expect(ribbon.getByRole('tab', { name: 'Design' })).toBeVisible();
    await expect(ribbon.getByRole('tab', { name: '开始' })).toHaveCount(0);
  });

  test('Markdown ribbon tabs resolve through the en-US catalog', async ({
    page,
  }) => {
    await gotoPlaygroundEnUS(page);
    await page
      .getByRole('button', { name: '# 产品说明 MD · 本次会话' })
      .click();
    await page
      .locator('.work-markdown-editor .ProseMirror')
      .waitFor({ state: 'attached' });

    const ribbon = page.locator('.work-markdown-ribbon');
    await expect(ribbon.getByRole('tab', { name: 'Home' })).toBeVisible();
    await expect(ribbon.getByRole('tab', { name: 'Insert' })).toBeVisible();
    await expect(ribbon.getByRole('tab', { name: '开始' })).toHaveCount(0);
  });

  test('PDF toolbar resolves through the en-US catalog', async ({ page }) => {
    await gotoPlaygroundEnUS(page);
    await openPdfFixture(page);
    await waitForPdfFixture(page);

    const toolbar = page.locator('.work-pdf-toolbar');
    await expect(toolbar).toHaveAttribute('aria-label', 'PDF toolbar');
    await expect(
      toolbar.getByRole('button', { name: 'Highlight' }),
    ).toBeVisible();
    await expect(toolbar.getByRole('button', { name: '高亮' })).toHaveCount(0);
  });
});
