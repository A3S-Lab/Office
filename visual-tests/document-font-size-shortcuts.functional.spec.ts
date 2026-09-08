import { expect, type Locator, test } from '@playwright/test';
import {
  openDocumentFixture,
  waitForDocumentFixture,
} from './visual-test-support';

test('Writer steps font size with WPS grow and shrink shortcuts', async ({
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

  const textRibbon = page.getByRole('region', { name: '文字功能区' });
  await expect(
    textRibbon.getByRole('button', { name: '增大字号' }),
  ).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+. Meta+Shift+. Control+] Meta+]',
  );
  await expect(
    textRibbon.getByRole('button', { name: '减小字号' }),
  ).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+, Meta+Shift+, Control+[ Meta+[',
  );

  const title = page.locator('.work-document-editable .ProseMirror h1').first();
  await title.click();
  await selectBlockText(title);
  await page.keyboard.press('Control+]');
  await expect(
    title.locator("[style*='font-size: 12pt']").first(),
  ).toBeVisible();
  await expect(textRibbon.getByRole('combobox', { name: '字号' })).toHaveText(
    '12',
  );

  await page.keyboard.press('Control+[');
  await expect(
    title.locator("[style*='font-size: 10.5pt']").first(),
  ).toBeVisible();
  await expect(
    page.locator('.work-document-editable .ProseMirror'),
  ).toBeFocused();

  await page.screenshot({
    path: testInfo.outputPath(
      `writer-font-size-shortcuts-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  expect(browserErrors).toEqual([]);
});

async function selectBlockText(block: Locator): Promise<void> {
  await block.evaluate((element) => {
    const range = document.createRange();
    range.selectNodeContents(element);
    const selection = window.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(range);
    document.dispatchEvent(new Event('selectionchange', { bubbles: true }));
  });
}
