import { expect, test } from '@playwright/test';
import {
  openDocumentFixture,
  waitForDocumentFixture,
} from './visual-test-support';

test('Writer opens the hyperlink dialog with WPS Ctrl+K', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });

  await page.goto('/playground/');
  await openDocumentFixture(page);
  await waitForDocumentFixture(page);

  await page.getByRole('tab', { name: '插入' }).click();
  const addLink = page.getByRole('button', { name: '添加链接' });
  await expect(addLink).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+K Meta+K',
  );

  // Prefer the title heading: first <p> Home/Shift+End is flaky under WASM
  // pagination on desktop CI.
  const heading = page
    .locator('.work-document-editable .ProseMirror h1')
    .first();
  await expect(heading).toContainText('新项目方案');
  await heading.click({ clickCount: 3 });
  await page.keyboard.press('Control+k');

  const linkDialog = page.getByRole('dialog', { name: '添加链接' });
  await expect(linkDialog).toBeVisible();
  const hrefField = linkDialog.getByRole('textbox', { name: '链接地址' });
  await expect(hrefField).toBeFocused();
  await hrefField.fill('https://a3s.dev/office');
  await linkDialog.getByRole('button', { name: '添加链接' }).click();
  await expect(
    page.locator(
      '.work-document-editable .ProseMirror a[href="https://a3s.dev/office"]',
    ),
  ).toHaveCount(1, { timeout: 15_000 });

  await page.screenshot({
    path: testInfo.outputPath(
      `writer-hyperlink-shortcut-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  expect(browserErrors).toEqual([]);
});
