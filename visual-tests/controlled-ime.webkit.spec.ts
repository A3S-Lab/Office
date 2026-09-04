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
  const publicationCount = () =>
    editor.evaluate((element) =>
      Number(
        (element as HTMLElement).dataset.documentControlledPublishCount ?? 0,
      ),
    );

  await editor.focus();
  const heading = editor.locator('h1').first();
  await heading.selectText();
  const publicationsBeforeComposition = await publicationCount();

  await editor.evaluate((element) => {
    element.dispatchEvent(
      new CompositionEvent('compositionstart', {
        bubbles: true,
        data: 'ni hao',
      }),
    );
  });
  await page.keyboard.insertText('ni hao');
  await expect(heading).toHaveText('ni hao');
  await expect.poll(publicationCount).toBe(publicationsBeforeComposition);

  await editor.evaluate((element) => {
    element.dispatchEvent(
      new CompositionEvent('compositionend', {
        bubbles: true,
        data: '你好',
      }),
    );
  });
  await page.waitForTimeout(5);
  await page.keyboard.insertText('你好');
  await expect.poll(publicationCount).toBe(publicationsBeforeComposition + 1);
  await expect(heading).toHaveText('你好');
  await expect(editor).not.toContainText('ni hao');

  await page.getByRole('button', { name: '返回办公首页' }).click();
  await openDocumentFixture(page);
  await waitForDocumentFixture(page);
  const reopened = page.getByRole('textbox', { name: '文档正文' });
  await expect(reopened.locator('h1').first()).toHaveText('你好');
  await expect(reopened).not.toContainText('ni hao');
  expect(browserErrors).toEqual([]);
});
