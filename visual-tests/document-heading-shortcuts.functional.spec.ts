import { expect, type Locator, test } from '@playwright/test';
import {
  openDocumentFixture,
  waitForDocumentFixture,
} from './visual-test-support';

const styleShortcutAds =
  'Control+Shift+N Meta+Shift+N Control+Alt+1 Meta+Alt+1 Control+Alt+2 Meta+Alt+2 Control+Alt+3 Meta+Alt+3';

test('Writer steps heading styles with WPS Ctrl+Alt+1 / Ctrl+Alt+2 / Ctrl+Alt+3', async ({
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
  const styleSelect = page.getByRole('combobox', { name: '段落样式' });
  const headingOne = gallery.locator('label[data-document-style="h1"]');
  if (await headingOne.isVisible()) {
    await expect(headingOne).toHaveAttribute(
      'title',
      '标题 1（Cmd/Ctrl+Alt+1）',
    );
    await expect(
      gallery.getByRole('radio', { name: '应用样式：标题 1' }),
    ).toHaveAttribute('aria-keyshortcuts', 'Control+Alt+1 Meta+Alt+1');
    await expect(
      gallery.getByRole('radio', { name: '应用样式：标题 2' }),
    ).toHaveAttribute('aria-keyshortcuts', 'Control+Alt+2 Meta+Alt+2');
    await expect(
      gallery.getByRole('radio', { name: '应用样式：标题 3' }),
    ).toHaveAttribute('aria-keyshortcuts', 'Control+Alt+3 Meta+Alt+3');
  } else {
    await expect(styleSelect).toHaveAttribute(
      'aria-keyshortcuts',
      styleShortcutAds,
    );
  }

  const editor = page.locator('.work-document-editable .ProseMirror');
  const body = editor.getByText(
    '用一句话说明这项工作的目标，以及完成后会带来什么变化。',
    { exact: true },
  );
  await body.click();
  await selectBlockText(body);

  await editor.press('Control+Alt+Digit1');
  await expect(editor.locator('h1', { hasText: '用一句话说明' })).toBeVisible();
  if (await headingOne.isVisible()) {
    await expect(headingOne).toHaveClass(/active/);
  } else {
    await expect(styleSelect).toHaveText('标题 1');
  }

  await editor.press('Control+Alt+Digit2');
  await expect(editor.locator('h2', { hasText: '用一句话说明' })).toBeVisible();
  if (await headingOne.isVisible()) {
    await expect(
      gallery.locator('label[data-document-style="h2"]'),
    ).toHaveClass(/active/);
  } else {
    await expect(styleSelect).toHaveText('标题 2');
  }

  await editor.press('Control+Alt+Digit3');
  await expect(editor.locator('h3', { hasText: '用一句话说明' })).toBeVisible();
  if (await headingOne.isVisible()) {
    await expect(
      gallery.locator('label[data-document-style="h3"]'),
    ).toHaveClass(/active/);
  } else {
    await expect(styleSelect).toHaveText('标题 3');
  }
  await expect(editor).toBeFocused();

  await page.screenshot({
    path: testInfo.outputPath(
      `writer-heading-shortcuts-${testInfo.project.name}.png`,
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
