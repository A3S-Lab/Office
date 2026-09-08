import { expect, type Locator, test } from '@playwright/test';
import {
  openDocumentFixture,
  waitForDocumentFixture,
} from './visual-test-support';

test('Writer toggles subscript and superscript with WPS shortcuts', async ({
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
  await expect(textRibbon.getByRole('button', { name: '下标' })).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+= Meta+=',
  );
  await expect(textRibbon.getByRole('button', { name: '上标' })).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+= Meta+Shift+=',
  );

  const title = page.locator('.work-document-editable .ProseMirror h1').first();
  await title.click();
  await selectBlockText(title);
  await page.keyboard.press('Control+=');
  await expect(title.locator('sub').first()).toBeVisible();
  await expect(textRibbon.getByRole('button', { name: '下标' })).toHaveAttribute(
    'aria-pressed',
    'true',
  );

  await page.keyboard.press('Control+Shift+=');
  await expect(title.locator('sup').first()).toBeVisible();
  await expect(title.locator('sub')).toHaveCount(0);
  await expect(textRibbon.getByRole('button', { name: '上标' })).toHaveAttribute(
    'aria-pressed',
    'true',
  );
  await expect(
    page.locator('.work-document-editable .ProseMirror'),
  ).toBeFocused();

  await page.screenshot({
    path: testInfo.outputPath(
      `writer-script-shortcuts-${testInfo.project.name}.png`,
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
