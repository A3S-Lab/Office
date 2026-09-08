import { expect, type Locator, test } from '@playwright/test';
import {
  openDocumentFixture,
  waitForDocumentFixture,
} from './visual-test-support';

test('Writer toggles bullet list with WPS Ctrl+Shift+L', async ({
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
  const bulletButton = textRibbon.getByRole('button', {
    name: '项目符号',
    exact: true,
  });
  await expect(bulletButton).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+L Meta+Shift+L',
  );

  const editor = page.locator('.work-document-editable .ProseMirror');
  const title = editor.getByText('新项目方案', { exact: true });
  await title.click();
  await selectBlockText(title);

  await editor.press('Control+Shift+l');
  await expect(
    editor.locator('ul > li > p', { hasText: '新项目方案' }),
  ).toBeVisible();
  await expect(bulletButton).toHaveClass(/active/);

  await editor.press('Control+Shift+l');
  await expect(
    editor.locator('ul > li > p', { hasText: '新项目方案' }),
  ).toHaveCount(0);
  await expect(editor.getByText('新项目方案', { exact: true })).toBeVisible();
  await expect(editor).toBeFocused();

  await page.screenshot({
    path: testInfo.outputPath(
      `writer-bullet-list-shortcut-${testInfo.project.name}.png`,
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
