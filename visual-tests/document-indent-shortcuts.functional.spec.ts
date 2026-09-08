import { expect, test } from '@playwright/test';
import {
  openDocumentFixture,
  waitForDocumentFixture,
} from './visual-test-support';

test('Writer steps paragraph indent with WPS Ctrl+M and Shift+Alt chords', async ({
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
    textRibbon.getByRole('button', { name: '增加缩进' }),
  ).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+M Meta+M Shift+Alt+. Alt+Shift+ArrowRight',
  );
  await expect(
    textRibbon.getByRole('button', { name: '减少缩进' }),
  ).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+Shift+M Meta+Shift+M Shift+Alt+, Alt+Shift+ArrowLeft',
  );

  const editor = page.locator('.work-document-editable .ProseMirror');
  const paragraph = editor.locator('p').first();
  await paragraph.click();
  await editor.press('Home');
  await editor.press('Control+m');
  await expect(paragraph).toHaveAttribute('data-office-indent-level', '1');

  await editor.press('Control+Shift+m');
  await expect(paragraph).not.toHaveAttribute('data-office-indent-level');

  await editor.press('Shift+Alt+Period');
  await expect(paragraph).toHaveAttribute('data-office-indent-level', '1');

  await editor.press('Shift+Alt+Comma');
  await expect(paragraph).not.toHaveAttribute('data-office-indent-level');

  await editor.press('Alt+Shift+ArrowRight');
  await expect(paragraph).toHaveAttribute('data-office-indent-level', '1');

  await editor.press('Alt+Shift+ArrowLeft');
  await expect(paragraph).not.toHaveAttribute('data-office-indent-level');
  await expect(editor).toBeFocused();

  await page.screenshot({
    path: testInfo.outputPath(
      `writer-indent-shortcuts-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  expect(browserErrors).toEqual([]);
});
