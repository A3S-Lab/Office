import { expect, type Locator, type Page, test } from '@playwright/test';
import {
  openDocumentFixture,
  waitForDocumentFixture,
} from './visual-test-support';

const styleShortcutAds =
  'Control+Shift+N Meta+Shift+N Control+Alt+1 Meta+Alt+1 Control+Alt+2 Meta+Alt+2 Control+Alt+3 Meta+Alt+3';

test('Writer applies Normal / 正文 style with WPS Ctrl+Shift+N', async ({
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

  const gallery = page.getByRole('radiogroup', { name: '段落样式库' });
  const normalStyle = gallery.locator('label[data-document-style="paragraph"]');
  const styleSelect = page.getByRole('combobox', { name: '段落样式' });
  if (await normalStyle.isVisible()) {
    await expect(normalStyle).toHaveAttribute(
      'title',
      '正文（Cmd/Ctrl+Shift+N）',
    );
    await expect(
      gallery.getByRole('radio', { name: '应用样式：正文' }),
    ).toHaveAttribute('aria-keyshortcuts', 'Control+Shift+N Meta+Shift+N');
  } else {
    await expect(styleSelect).toHaveAttribute(
      'aria-keyshortcuts',
      styleShortcutAds,
    );
  }

  const editor = page.locator('.work-document-editable .ProseMirror');
  const title = editor.locator('h1').first();
  await title.click();
  await selectBlockText(title);
  await editor.press('Control+Alt+Digit2');
  const headingTwo = editor.locator('h2').last();
  await expect(headingTwo).toBeVisible();

  await selectBlockText(headingTwo);
  // Chrome steals real Ctrl+Shift+N (new window); deliver the chord to the
  // Writer capture listener. ACL/CDP covers the live key path.
  await dispatchWriterShortcut(page, { key: 'n', code: 'KeyN', shiftKey: true });
  if (await normalStyle.isVisible()) {
    await expect(normalStyle).toHaveClass(/active/);
  } else {
    await expect(styleSelect).toHaveText('正文');
  }
  await expect(editor).toBeFocused();

  await page.screenshot({
    path: testInfo.outputPath(
      `writer-normal-style-shortcut-${testInfo.project.name}.png`,
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

async function dispatchWriterShortcut(
  page: Page,
  event: { key: string; code: string; shiftKey?: boolean },
): Promise<void> {
  await page.locator('.work-document-editable .ProseMirror').evaluate(
    (element, detail) => {
      element.dispatchEvent(
        new KeyboardEvent('keydown', {
          key: detail.key,
          code: detail.code,
          ctrlKey: true,
          shiftKey: Boolean(detail.shiftKey),
          bubbles: true,
          cancelable: true,
        }),
      );
    },
    event,
  );
}
