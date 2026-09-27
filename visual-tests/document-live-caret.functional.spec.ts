import { expect, test } from '@playwright/test';
import {
  openDocumentFixture,
  waitForDocumentFixture,
} from './visual-test-support';

test('two browsers paint the native splice caret after pagination at a non-100% zoom', async ({
  page,
}, testInfo) => {
  test.skip(testInfo.project.name !== 'desktop-1280', 'desktop pagination');
  const sourceErrors: string[] = [];
  page.on('pageerror', (error) => sourceErrors.push(error.message));
  const observer = await page.context().newPage();
  const observerErrors: string[] = [];
  observer.on('pageerror', (error) => observerErrors.push(error.message));

  await page.goto(
    '/playground/?e2e=collaboration-document-live-caret&peer=source',
  );
  await observer.goto(
    '/playground/?e2e=collaboration-document-live-caret&peer=observer',
  );
  await openDocumentFixture(page);
  await openDocumentFixture(observer);
  await waitForDocumentFixture(page);
  await waitForDocumentFixture(observer);

  const sourceZoom = page.locator('output[aria-label="文档缩放比例"]');
  const observerZoom = observer.locator('output[aria-label="文档缩放比例"]');
  await expect(sourceZoom).toHaveText('90%');
  await expect(observerZoom).toHaveText('90%');

  const advance = page.getByRole('button', { name: '插入下一字' });
  await advance.click();
  await advance.click();
  await expect(page.getByTestId('native-live-caret-status')).toHaveAttribute(
    'data-applied',
    '2',
  );
  await expect(
    observer.getByTestId('native-live-caret-status'),
  ).toHaveAttribute('data-applied', '2');
  await expect(page.getByTestId('native-live-caret-status')).toHaveAttribute(
    'data-caret-index',
    '96',
  );

  await page.getByRole('button', { name: '放大文档' }).click();
  await expect(sourceZoom).toHaveText('95%');
  await expect(observerZoom).toHaveText('90%');

  for (const browser of [page, observer]) {
    const wrapped = await browser.evaluate(() => {
      const paragraph = document.querySelector(
        '[data-office-paragraph-id="00000001"]',
      );
      const caret = document.querySelector('[data-frame-caret="true"]');
      if (!(paragraph instanceof HTMLElement) || !(caret instanceof HTMLElement)) {
        const prose = document.querySelector('.ProseMirror');
        return {
          lines: 0,
          missingParagraph: !(paragraph instanceof HTMLElement),
          missingCaret: !(caret instanceof HTMLElement),
          prose: prose?.innerHTML.slice(0, 400) ?? '',
        };
      }
      const range = document.createRange();
      range.selectNodeContents(paragraph);
      const lines = [...range.getClientRects()];
      const caretBox = caret.getBoundingClientRect();
      const first = lines[0];
      if (!first || lines.length < 2) return { lines: lines.length };
      return {
        lines: lines.length,
        caretOnLaterLine: caretBox.top >= first.bottom - 1,
        text: paragraph.textContent?.includes('中文') ?? false,
      };
    });
    if (wrapped && 'missingCaret' in wrapped) {
      throw new Error(JSON.stringify(wrapped));
    }
    expect(wrapped?.lines).toBeGreaterThan(1);
    expect(wrapped?.caretOnLaterLine).toBe(true);
    expect(wrapped?.text).toBe(true);
  }
  expect(sourceErrors).toEqual([]);
  expect(observerErrors).toEqual([]);
  await observer.close();
});
