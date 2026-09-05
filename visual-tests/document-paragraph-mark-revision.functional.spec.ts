import { expect, type Locator, test } from '@playwright/test';
import {
  createWordParagraphMarkRevisionFixture,
  PARAGRAPH_MARK_DELETION_TEXT,
  PARAGRAPH_MARK_INSERTION_TEXT,
  PARAGRAPH_MARK_STABLE_TEXT,
} from '../tests/fixtures/word-paragraph-mark-revision';
import { waitForDocumentFixture } from './visual-test-support';

test('Writer reviews WPS whole-paragraph revisions atomically on desktop and phone', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });
  if (testInfo.project.name === 'compact-768') {
    await page.setViewportSize({ width: 390, height: 844 });
  }

  await page.goto('/playground/');
  await page
    .locator('input[aria-label="打开 Office 或 PDF 文件"]')
    .setInputFiles({
      name: 'wps-paragraph-mark-revisions.docx',
      mimeType:
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      buffer: await createWordParagraphMarkRevisionFixture(),
    });
  await waitForDocumentFixture(page);

  const editor = page.getByRole('textbox', { name: '文档正文' });
  await expect(
    editor.locator(
      '[data-document-block-change="true"][data-block-change-kind="insertion"]',
    ),
  ).toContainText(PARAGRAPH_MARK_INSERTION_TEXT);
  await expect(
    editor.locator(
      '[data-document-block-change="true"][data-block-change-kind="deletion"]',
    ),
  ).toContainText(PARAGRAPH_MARK_DELETION_TEXT);
  await page.getByRole('tab', { name: '审阅' }).click();
  const reviewButton = page.getByRole('button', { name: '查看修订（2）' });
  await reviewButton.click();

  const pane = page.locator('.work-document-changes-panel');
  await expect(pane).toBeVisible();
  const list = pane.getByRole('list', { name: '待处理修订' });
  await expect(list).toHaveAttribute('data-document-change-count', '2');
  const insertionCard = pane.locator('.work-document-change-item.insertion');
  const deletionCard = pane.locator('.work-document-change-item.deletion');
  await expect(insertionCard).toHaveCount(1);
  await expect(deletionCard).toHaveCount(1);
  await expect(insertionCard.getByText('插入', { exact: true })).toBeVisible();
  await expect(deletionCard.getByText('删除', { exact: true })).toBeVisible();
  await expect(insertionCard).toContainText(PARAGRAPH_MARK_INSERTION_TEXT);
  await expect(deletionCard).toContainText(PARAGRAPH_MARK_DELETION_TEXT);
  await expect(insertionCard).toContainText('WPS Reference');
  await expect(deletionCard).toContainText('WPS Reference');

  if (testInfo.project.name === 'compact-768') {
    await expect(pane).toHaveAttribute('role', 'dialog');
    await expect(pane).toHaveAttribute('aria-modal', 'true');
    await expectContainedInViewport(pane);
    const actionHeights = await pane
      .locator('.work-document-change-item > div button')
      .evaluateAll((buttons) =>
        buttons.map((button) => button.getBoundingClientRect().height),
      );
    expect(Math.min(...actionHeights)).toBeGreaterThanOrEqual(44);
  } else {
    await insertionCard.getByRole('button', { name: '定位修订 1' }).focus();
    await page.keyboard.press('ArrowDown');
    await expect(
      deletionCard.getByRole('button', { name: '定位修订 2' }),
    ).toBeFocused();
  }

  await page.screenshot({
    path: testInfo.outputPath(
      `writer-paragraph-mark-review-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });

  await insertionCard.getByRole('button', { name: /接受修订/ }).click();
  await expect(list).toHaveAttribute('data-document-change-count', '1');
  await expect(editor).toContainText(PARAGRAPH_MARK_INSERTION_TEXT);
  await expect(
    editor.locator('ins').filter({ hasText: PARAGRAPH_MARK_INSERTION_TEXT }),
  ).toHaveCount(0);

  await pane
    .locator('.work-document-change-item.deletion')
    .getByRole('button', { name: /接受修订/ })
    .click();
  await expect(editor).not.toContainText(PARAGRAPH_MARK_DELETION_TEXT);
  await expect(editor).toContainText(PARAGRAPH_MARK_STABLE_TEXT);
  await expect(
    editor.locator('[data-document-block-change="true"]'),
  ).toHaveCount(0);
  await expect(pane.getByText('正在记录新的改动。')).toBeVisible();
  expect(browserErrors).toEqual([]);
});

async function expectContainedInViewport(element: Locator): Promise<void> {
  const geometry = await element.evaluate((node) => {
    const bounds = node.getBoundingClientRect();
    return {
      bottom: bounds.bottom,
      left: bounds.left,
      right: bounds.right,
      top: bounds.top,
      viewportHeight: document.documentElement.clientHeight,
      viewportWidth: document.documentElement.clientWidth,
    };
  });
  expect(geometry.left).toBeGreaterThanOrEqual(0);
  expect(geometry.right).toBeLessThanOrEqual(geometry.viewportWidth);
  expect(geometry.top).toBeGreaterThanOrEqual(0);
  expect(geometry.bottom).toBeLessThanOrEqual(geometry.viewportHeight);
}
