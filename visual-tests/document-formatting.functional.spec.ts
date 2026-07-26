import { expect, test, type Locator, type Page } from '@playwright/test';
import {
  openDocumentFixture,
  waitForDocumentFixture,
} from './visual-test-support';

const fontPreviews = [
  ['微软雅黑', '"Microsoft YaHei", "PingFang SC", sans-serif'],
  ['宋体', 'SimSun, "Songti SC", serif'],
  ['黑体', 'SimHei, "Heiti SC", sans-serif'],
  ['楷体', 'KaiTi, "Kaiti SC", serif'],
  ['Arial', 'Arial, sans-serif'],
  ['Times New Roman', '"Times New Roman", serif'],
] as const;

test('Word formatting controls apply computed styles and preview their fonts', async ({
  page,
}) => {
  const pageErrors: Error[] = [];
  page.on('pageerror', (error) => pageErrors.push(error));
  await page.goto('/');
  await openDocumentFixture(page);
  await waitForDocumentFixture(page);

  const editor = page.locator('.work-document-editable .ProseMirror');
  const selectedText = await selectParagraphText(page, 4, 10);
  const ribbon = page.getByRole('region', { name: '文字功能区' });
  await ribbon.getByRole('button', { name: '斜体', exact: true }).click();

  const italic = editor.locator('em').filter({ hasText: selectedText });
  await expect(italic).toHaveCount(1);
  await expect(italic).toHaveCSS('font-style', 'italic');
  await expect(editor).toContainText(selectedText);

  await ribbon.getByRole('combobox', { name: '字体' }).click();
  const fontMenu = page.getByRole('listbox', { name: '字体' });
  await expect(fontMenu).toBeVisible();
  await expectWithinViewport(fontMenu);
  for (const [label, family] of fontPreviews) {
    const option = fontMenu.getByRole('option', { name: label });
    await expect(option).toBeVisible();
    expect(
      await option
        .locator('span')
        .evaluate((element) => getComputedStyle(element).fontFamily),
    ).toBe(family);
  }
  await page.keyboard.press('Escape');

  await ribbon.getByRole('button', { name: '文字颜色' }).click();
  const colorMenu = page.getByRole('dialog', { name: '文字颜色' });
  await expect(colorMenu).toBeVisible();
  await expect(colorMenu.getByText('主题颜色')).toBeVisible();
  await expect(colorMenu.getByText('标准色')).toBeVisible();
  await expect(colorMenu.getByRole('option')).toHaveCount(60);
  await expectWithinViewport(colorMenu);
  await colorMenu.getByRole('option', { name: '颜色 #0070c0' }).click();
  await expect(italic).toHaveCSS('color', 'rgb(0, 112, 192)');
  expect(pageErrors).toEqual([]);
});

test('Word paper starts without rulers and persists explicit page controls', async ({
  page,
}) => {
  const pageErrors: Error[] = [];
  page.on('pageerror', (error) => pageErrors.push(error));
  await page.goto('/');
  await openDocumentFixture(page);
  await waitForDocumentFixture(page);

  await expect(page.getByRole('slider', { name: '左页边距' })).toHaveCount(0);
  await expect(page.getByRole('slider', { name: '上页边距' })).toHaveCount(0);

  const header = page.locator('.work-document-page-header');
  const footer = page.locator('.work-document-page-footer');
  await expect(
    header.locator('.work-document-page-chrome-placeholder'),
  ).toHaveCount(0);
  await expect(
    footer.locator('.work-document-page-chrome-placeholder'),
  ).toHaveCount(0);
  await header.hover();
  await expect(header).toHaveCSS('border-bottom-color', 'rgba(0, 0, 0, 0)');
  await expect(footer).toHaveCSS('border-top-color', 'rgba(0, 0, 0, 0)');
  const chromeGeometry = await page.evaluate(() => {
    const paper = document.querySelector<HTMLElement>('.work-document-page');
    const body = document.querySelector<HTMLElement>('.work-document-editable');
    const header = document.querySelector<HTMLElement>(
      '.work-document-page-header',
    );
    const footer = document.querySelector<HTMLElement>(
      '.work-document-page-footer',
    );
    if (!(paper && body && header && footer)) {
      throw new Error('Document page chrome geometry is unavailable.');
    }
    const paperRect = paper.getBoundingClientRect();
    const bodyRect = body.getBoundingClientRect();
    const headerRect = header.getBoundingClientRect();
    const footerRect = footer.getBoundingClientRect();
    return {
      paperTop: paperRect.top,
      paperBottom: paperRect.bottom,
      bodyTop: bodyRect.top,
      bodyBottom: bodyRect.bottom,
      headerTop: headerRect.top,
      headerBottom: headerRect.bottom,
      footerTop: footerRect.top,
      footerBottom: footerRect.bottom,
    };
  });
  expect(
    Math.abs(chromeGeometry.headerTop - chromeGeometry.paperTop),
  ).toBeLessThanOrEqual(1);
  expect(
    Math.abs(chromeGeometry.headerBottom - chromeGeometry.bodyTop),
  ).toBeLessThanOrEqual(1);
  expect(chromeGeometry.footerTop).toBeGreaterThanOrEqual(
    chromeGeometry.bodyBottom,
  );
  expect(
    Math.abs(chromeGeometry.footerBottom - chromeGeometry.paperBottom),
  ).toBeLessThanOrEqual(1);
  await header.dblclick();
  await expect(page.getByRole('textbox', { name: '页内页眉' })).toBeFocused();
  await expect(page.getByRole('tab', { name: '页眉和页脚' })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  await expect(header).not.toHaveCSS('border-bottom-color', 'rgba(0, 0, 0, 0)');
  await page.keyboard.press('Escape');
  await expect(page.getByRole('textbox', { name: '文档正文' })).toBeFocused();
  await expect(page.getByRole('tab', { name: '页眉和页脚' })).toHaveCount(0);

  await page.getByRole('tab', { name: '插入' }).click();
  await page.getByRole('button', { name: '页脚' }).click();
  await expect(page.getByRole('textbox', { name: '页内页脚' })).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(page.getByRole('textbox', { name: '文档正文' })).toBeFocused();

  await page.getByRole('tab', { name: '页面布局' }).click();
  await page.getByRole('button', { name: '页面颜色' }).click();
  const pageColorMenu = page.getByRole('dialog', { name: '页面颜色' });
  await expect(pageColorMenu).toBeVisible();
  await expect(pageColorMenu.getByText('主题颜色')).toBeVisible();
  await expect(pageColorMenu.getByText('标准色')).toBeVisible();
  await expect(pageColorMenu.getByRole('option')).toHaveCount(60);
  await expectWithinViewport(pageColorMenu);
  await pageColorMenu.getByRole('option', { name: '颜色 #fff2cc' }).click();
  await expect(page.locator('.work-document-page')).toHaveCSS(
    'background-color',
    'rgb(255, 242, 204)',
  );

  await page.getByRole('tab', { name: '视图' }).click();
  const ruler = page.getByRole('button', { name: '标尺', exact: true });
  await expect(ruler).toHaveAttribute('aria-pressed', 'false');
  await ruler.click();
  await expect(page.getByRole('slider', { name: '左页边距' })).toBeVisible();
  await expect(page.getByRole('slider', { name: '上页边距' })).toBeVisible();
  await expect(ruler).toHaveAttribute('aria-pressed', 'true');

  const viewModes = page.getByRole('region', { name: '文档视图' });
  await viewModes.getByRole('button', { name: '网页视图' }).click();
  await expect(page.getByRole('slider', { name: '左页边距' })).toHaveCount(0);
  await expect(ruler).toBeDisabled();
  await viewModes.getByRole('button', { name: '页面视图' }).click();
  await expect(page.getByRole('slider', { name: '左页边距' })).toBeVisible();
  expect(pageErrors).toEqual([]);
});

async function selectParagraphText(
  page: Page,
  paragraphIndex: number,
  characterCount: number,
): Promise<string> {
  return page.locator('.work-document-editable .ProseMirror').evaluate(
    (root, options) => {
      const paragraph = root.querySelectorAll('p')[options.paragraphIndex];
      const text = paragraph?.firstChild;
      if (!(text instanceof Text))
        throw new Error('Paragraph text is missing.');
      const selectedText = text.textContent?.slice(0, options.characterCount);
      if (!selectedText) throw new Error('Paragraph text is empty.');
      const range = document.createRange();
      range.setStart(text, 0);
      range.setEnd(text, selectedText.length);
      const selection = window.getSelection();
      selection?.removeAllRanges();
      selection?.addRange(range);
      document.dispatchEvent(new Event('selectionchange', { bubbles: true }));
      return selectedText;
    },
    { paragraphIndex, characterCount },
  );
}

async function expectWithinViewport(surface: Locator): Promise<void> {
  const geometry = await surface.evaluate((element) => {
    const rect = element.getBoundingClientRect();
    return {
      viewportWidth: document.documentElement.clientWidth,
      viewportHeight: document.documentElement.clientHeight,
      left: rect.left,
      top: rect.top,
      right: rect.right,
      bottom: rect.bottom,
    };
  });
  expect(geometry.left).toBeGreaterThanOrEqual(8);
  expect(geometry.top).toBeGreaterThanOrEqual(8);
  expect(geometry.right).toBeLessThanOrEqual(geometry.viewportWidth - 8);
  expect(geometry.bottom).toBeLessThanOrEqual(geometry.viewportHeight - 8);
}
