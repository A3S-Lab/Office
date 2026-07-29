import { expect, test } from '@playwright/test';
import { openPdfFixture, waitForPdfFixture } from './pdf-test-support';

test('PDF toolbar shortcuts stay inside the editor command surface', async ({
  page,
}) => {
  await page.goto('/');
  await openPdfFixture(page);
  await waitForPdfFixture(page);

  const search = page.getByRole('searchbox', { name: '在 PDF 中搜索' });
  await page.getByRole('button', { name: '选择' }).focus();
  await page.keyboard.press('Meta+f');
  await expect(search).toBeFocused();

  const zoom = page.locator('output[aria-label="PDF 缩放比例"]');
  const initialZoom = await zoom.textContent();
  await page.keyboard.press('Meta+=');
  await expect(zoom).not.toHaveText(initialZoom ?? '');

  const pen = page.getByRole('button', { name: '画笔' });
  await pen.click();
  await expect(pen).toHaveAttribute('aria-pressed', 'true');
  await page.keyboard.press('Escape');
  await expect(pen).toHaveAttribute('aria-pressed', 'false');
});

test('PDF shortcuts do not take over controls outside the editor', async ({
  page,
}) => {
  await page.goto('/');
  await openPdfFixture(page);
  await waitForPdfFixture(page);

  const search = page.getByRole('searchbox', { name: '在 PDF 中搜索' });
  const hostAction = page.getByRole('button', {
    name: '打开 AI 助手',
    exact: true,
  });
  await hostAction.focus();
  await hostAction.press('Meta+f');
  await expect(hostAction).toBeFocused();
  await expect(search).not.toBeFocused();

  const editorAction = page.getByRole('button', { name: '选择' });
  await editorAction.focus();
  await editorAction.press('Meta+f');
  await expect(search).toBeFocused();
});

test('PDF overflow menu uses keyboard navigation and restores its trigger', async ({
  page,
}) => {
  await page.goto('/');
  await openPdfFixture(page);
  await waitForPdfFixture(page);

  const trigger = page.getByRole('button', { name: '更多 PDF 工具' });
  await expect(trigger).toBeVisible();
  await trigger.click();
  const menu = page.getByRole('menu', { name: '更多 PDF 工具' });
  await expect(menu).toBeVisible();
  const firstAction = menu.getByRole('menuitemradio', { name: '下划线批注' });
  await expect(firstAction).toBeFocused();
  await page.keyboard.press('ArrowDown');
  await expect(
    menu.getByRole('menuitemradio', { name: '删除线批注' }),
  ).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(menu).toBeHidden();
  await expect(trigger).toBeFocused();
});

test('PDF page Escape cancels navigation and annotation styles exit in document order', async ({
  page,
}) => {
  await page.goto('/');
  await openPdfFixture(page, { pageCount: 2 });
  await waitForPdfFixture(page);

  const pageField = page.getByRole('textbox', { name: '页码' });
  await expect(pageField).toHaveValue('1');
  await pageField.fill('2');
  await pageField.press('Escape');
  await expect(pageField).toHaveValue('1');
  await expect(pageField).not.toBeFocused();

  const pen = page.getByRole('button', { name: '画笔' });
  await pen.click();
  const styleTrigger = page.getByRole('button', { name: '批注样式' });
  await styleTrigger.click();
  const styleDialog = page.getByRole('dialog', { name: '批注样式' });
  await expect(styleDialog).toBeVisible();
  await expect(
    styleDialog.getByRole('radio', { name: '透明度 100%' }),
  ).toBeFocused();

  await page.keyboard.press('Tab');
  await expect(
    styleDialog.getByRole('radio', { name: '线宽 6' }),
  ).toBeFocused();
  await page.keyboard.press('Tab');
  await expect(styleDialog).toBeHidden();
  await expect(
    page.getByRole('button', { name: '更多 PDF 工具' }),
  ).toBeFocused();
});

test('PDF search advances a settled query and clear keeps the search field active', async ({
  page,
}, testInfo) => {
  await page.goto('/');
  await openPdfFixture(page, { pageCount: 2 });
  await waitForPdfFixture(page);

  const search = page.getByRole('searchbox', { name: '在 PDF 中搜索' });
  const status = page.locator('.work-pdf-search-state');
  await search.fill('A3S');
  await search.press('Enter');
  await expect(status).toHaveText('1 / 2', { timeout: 30_000 });

  await search.press('Enter');
  await expect(status).toHaveText('2 / 2');
  await search.press('Shift+Enter');
  await expect(status).toHaveText('1 / 2');

  await page.getByRole('button', { name: '清除搜索' }).click();
  await expect(search).toHaveValue('');
  await expect(search).toBeFocused();
  await expect(status).toBeEmpty();
  if (testInfo.project.name === 'compact-768') {
    await page.getByRole('button', { name: '更多 PDF 工具' }).click();
    const menu = page.getByRole('menu', { name: '更多 PDF 工具' });
    await expect(
      menu.getByRole('menuitem', { name: '上一个搜索结果' }),
    ).toBeDisabled();
    await expect(
      menu.getByRole('menuitem', { name: '下一个搜索结果' }),
    ).toBeDisabled();
  } else {
    await expect(
      page.getByRole('button', { name: '上一个搜索结果' }),
    ).toBeDisabled();
    await expect(
      page.getByRole('button', { name: '下一个搜索结果' }),
    ).toBeDisabled();
  }
});

test('PDF command states use the same product accent as the file identity', async ({
  page,
}) => {
  await page.goto('/');
  await openPdfFixture(page);
  await waitForPdfFixture(page);

  const productAccent = await page
    .locator('.work-file-kind-icon.pdf')
    .evaluate((element) => getComputedStyle(element).color);
  const selectedToolAccent = await page
    .getByRole('button', { name: '选择' })
    .evaluate((element) => getComputedStyle(element).color);
  expect(selectedToolAccent).toBe(productAccent);

  const search = page.getByRole('searchbox', { name: '在 PDF 中搜索' });
  await search.click();
  const searchFocusAccent = await page
    .locator('.work-pdf-search')
    .evaluate((element) => getComputedStyle(element).borderTopColor);
  expect(searchFocusAccent).toBe(productAccent);
});
