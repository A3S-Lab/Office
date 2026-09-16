import { expect, test } from '@playwright/test';

const LONG_MARKDOWN = Array.from({ length: 100 }, (_, index) => {
  const line = String(index + 1).padStart(2, '0');
  return `Line ${line} — keep Markdown source scrolling on the textarea so proportional preview sync stays attached to onScroll.`;
}).join('\n');

test('Markdown source wheel scrolls the textarea, not the pane shell', async ({
  page,
}) => {
  await page.goto('/playground/?e2e=markdown-source-scroll');
  await page.getByRole('button', { name: /产品说明/ }).click();

  const source = page.getByRole('textbox', { name: 'Markdown 源码' });
  const preview = page.getByRole('document', { name: 'Markdown 预览' });
  await expect(source).toBeVisible();
  await expect(preview).toBeVisible();

  await source.fill(LONG_MARKDOWN);

  const geometry = await source.evaluate((element: HTMLTextAreaElement) => {
    const pane = element.closest('.work-markdown-pane.source');
    if (!(pane instanceof HTMLElement)) {
      throw new Error('Missing Markdown source pane.');
    }
    return {
      sourceScrollable: element.scrollHeight > element.clientHeight + 8,
      paneScrollable: pane.scrollHeight > pane.clientHeight + 8,
      paneOverflow: getComputedStyle(pane).overflow,
      textareaOverflow: getComputedStyle(element).overflow,
      textareaMinHeight: getComputedStyle(element).minHeight,
    };
  });

  expect(geometry.sourceScrollable).toBe(true);
  expect(geometry.paneScrollable).toBe(false);
  expect(geometry.paneOverflow).toBe('hidden');
  expect(geometry.textareaOverflow).toMatch(/auto|scroll/);
  expect(geometry.textareaMinHeight).not.toBe('100%');

  await source.focus();
  await source.hover();
  await page.mouse.wheel(0, 900);

  await expect(source).toHaveAttribute('data-source-scrolled', 'true');
  await expect(
    page.locator('.work-markdown-pane.source[data-pane-scrolled="false"]'),
  ).toBeVisible();
  await expect(
    page.locator(
      '.work-markdown-pane.visual[data-preview-scroll-synced="true"]',
    ),
  ).toBeVisible();

  const afterScroll = await source.evaluate((element: HTMLTextAreaElement) => {
    const pane = element.closest('.work-markdown-pane.source');
    const visual = document.querySelector('.work-markdown-pane.visual');
    if (!(pane instanceof HTMLElement) || !(visual instanceof HTMLElement)) {
      throw new Error('Missing Markdown scroll panes.');
    }
    return {
      sourceScrollTop: element.scrollTop,
      paneScrollTop: pane.scrollTop,
      previewScrollTop: visual.scrollTop,
    };
  });

  expect(afterScroll.sourceScrollTop).toBeGreaterThan(0);
  expect(afterScroll.paneScrollTop).toBe(0);
  expect(afterScroll.previewScrollTop).toBeGreaterThan(0);
});
