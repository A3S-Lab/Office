import { expect, test } from '@playwright/test';
import {
  openDocumentFixture,
  waitForDocumentFixture,
} from './visual-test-support';

test('Writer opens find and replace with WPS Ctrl+F / Ctrl+H', async ({
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

  const findButton = page.getByRole('button', { name: '查找', exact: true });
  const replaceButton = page.getByRole('button', { name: '替换', exact: true });
  await expect(findButton).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+F Meta+F',
  );
  await expect(replaceButton).toHaveAttribute(
    'aria-keyshortcuts',
    'Control+H Meta+H',
  );

  const body = page.getByRole('textbox', { name: '文档正文' });
  await body.focus();
  await page.keyboard.press('Control+f');

  const findPanel = page.locator('.work-document-find-panel');
  await expect(findPanel).toBeVisible();
  await expect(findPanel.getByRole('tab', { name: '查找' })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  const query = findPanel.getByRole('textbox', { name: '查找内容' });
  await expect(query).toBeFocused();

  await page.keyboard.press('Control+h');
  await expect(findPanel.getByRole('tab', { name: '替换' })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  await expect(query).toBeFocused();
  await expect(
    findPanel.getByRole('textbox', { name: '替换为' }),
  ).toBeVisible();

  await page.screenshot({
    path: testInfo.outputPath(
      `writer-find-replace-shortcut-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  expect(browserErrors).toEqual([]);
});
