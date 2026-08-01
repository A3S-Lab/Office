import { expect, test } from '@playwright/test';
import {
  openDocumentFixture,
  waitForDocumentFixture,
} from './visual-test-support';

test('Word page navigation previews measured pages and jumps without losing editor state', async ({
  page,
}, testInfo) => {
  const pageErrors: Error[] = [];
  page.on('pageerror', (error) => pageErrors.push(error));

  await page.goto('/');
  await openDocumentFixture(page);
  await waitForDocumentFixture(page);

  const editor = page.getByRole('textbox', { name: '文档正文' });
  const longDocument = Array.from(
    { length: 1_200 },
    (_, index) => `page-token-${String(index + 1).padStart(4, '0')}`,
  ).join(' ');
  await editor.fill(longDocument);
  await expect(editor).toHaveAttribute('data-pagination-state', 'ready');
  await expect
    .poll(async () =>
      Number(await editor.getAttribute('data-pagination-pages')),
    )
    .toBeGreaterThan(2);

  await page.getByRole('tab', { name: '视图' }).click();
  await page.getByRole('button', { name: '导航窗格' }).click();
  const pane = page.locator('.work-document-navigation-panel');
  await pane.getByRole('tab', { name: '页面' }).click();

  const navigation = pane.getByRole('navigation', { name: '文档页面' });
  const thumbnails = navigation.locator('[data-document-page-thumbnail]');
  await expect
    .poll(async () => {
      const pageCount = Number(
        await editor.getAttribute('data-pagination-pages'),
      );
      return (
        (await editor.getAttribute('data-pagination-state')) === 'ready' &&
        pageCount > 2 &&
        (await thumbnails.count()) === pageCount
      );
    })
    .toBe(true);
  const pageCount = Number(await editor.getAttribute('data-pagination-pages'));
  await expect(thumbnails).toHaveCount(pageCount);

  const firstPage = navigation.getByRole('button', {
    name: '第 1 页',
    exact: true,
  });
  const secondPage = navigation.getByRole('button', {
    name: '第 2 页',
    exact: true,
  });
  const thirdPage = navigation.getByRole('button', {
    name: '第 3 页',
    exact: true,
  });
  await expect(
    navigation.getByRole('button', {
      name: `第 ${pageCount} 页`,
      exact: true,
    }),
  ).toHaveAttribute('aria-current', 'page');
  const firstPreview = await firstPage
    .locator('.work-document-page-thumbnail > span')
    .textContent();
  const secondPreview = await secondPage
    .locator('.work-document-page-thumbnail > span')
    .textContent();
  expect(firstPreview?.trim()).toBeTruthy();
  expect(secondPreview?.trim()).toBeTruthy();
  expect(secondPreview).not.toBe(firstPreview);

  await secondPage.click();
  await expect(page.getByLabel('页码状态')).toContainText('第 2 页');

  if (testInfo.project.name === 'compact-768') {
    await expect(pane).toBeHidden();
    await expect(editor).toBeFocused();
  } else {
    await expect(secondPage).toHaveAttribute('aria-current', 'page');
    await secondPage.press('ArrowDown');
    await expect(thirdPage).toBeFocused();
    await thirdPage.press('Home');
    await expect(firstPage).toBeFocused();
  }

  expect(pageErrors).toEqual([]);
});
