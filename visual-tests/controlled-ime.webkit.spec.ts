import { expect, test } from '@playwright/test';
import {
  openDocumentFixture,
  waitForDocumentFixture,
} from './visual-test-support';

test('controlled Document editing publishes only committed Chinese IME text in WebKit', async ({
  page,
}) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));

  await page.goto('/playground/');
  await openDocumentFixture(page);
  await waitForDocumentFixture(page);

  const editor = page.getByRole('textbox', { name: '文档正文' });
  const modifier = process.platform === 'darwin' ? 'Meta' : 'Control';
  const publicationCount = () =>
    editor.evaluate((element) =>
      Number(
        (element as HTMLElement).dataset.documentControlledPublishCount ?? 0,
      ),
    );

  await editor.focus();
  await page.keyboard.press(`${modifier}+a`);
  const publicationsBeforeComposition = await publicationCount();

  await editor.dispatchEvent('compositionstart', { data: 'qingwen' });
  await editor.fill('qingwen');
  await expect(editor).toHaveText('qingwen');
  await expect.poll(publicationCount).toBe(publicationsBeforeComposition);

  await editor.fill('请问');
  await expect(editor).toHaveText('请问');
  await expect.poll(publicationCount).toBe(publicationsBeforeComposition);

  await editor.dispatchEvent('compositionend', { data: '请问' });
  await expect.poll(publicationCount).toBe(publicationsBeforeComposition + 1);
  await expect(editor).toHaveText('请问');
  await expect(editor).not.toContainText('qingwen');

  await page.getByRole('button', { name: '返回办公首页' }).click();
  await openDocumentFixture(page);
  await waitForDocumentFixture(page);
  const reopened = page.getByRole('textbox', { name: '文档正文' });
  await expect(reopened).toHaveText('请问');
  await expect(reopened).not.toContainText('qingwen');
  expect(browserErrors).toEqual([]);
});
