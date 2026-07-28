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

test('Word keeps browser-synthesized bold text on deterministic layout', async ({
  page,
}) => {
  const pageErrors: Error[] = [];
  page.on('pageerror', (error) => pageErrors.push(error));
  await page.goto('/');
  await page
    .getByRole('button', {
      name: '空白文字 从一张干净的 A4 页面开始',
    })
    .click();

  const editor = page.getByRole('textbox', { name: '文档正文' });
  await expect(editor).toHaveAttribute('data-pagination-state', 'ready');
  await editor.click();
  await page
    .getByRole('region', { name: '文字功能区' })
    .getByRole('button', { name: '加粗', exact: true })
    .click();
  await expect(editor).toBeFocused();
  await page.keyboard.type('Bold layout metrics');

  await expect(editor.locator('strong')).toHaveText('Bold layout metrics');
  await expect(editor).toHaveAttribute('data-pagination-state', 'ready');
  await expect(editor).toHaveAttribute('data-pagination-text-engine', 'wasm');
  await expect(editor).toHaveAttribute('data-pagination-text-candidates', '1');
  await expect(editor).toHaveAttribute(
    'data-pagination-shaped-paragraphs',
    '1',
  );
  await expect(editor).toHaveAttribute('data-pagination-shaped-runs', '1');
  await expect(editor).toHaveAttribute('data-pagination-unsupported-text', '0');
  expect(pageErrors).toEqual([]);
});

test('Word preview reuses the live WASM pagination result', async ({
  page,
}) => {
  const pageErrors: Error[] = [];
  page.on('pageerror', (error) => pageErrors.push(error));
  await page.goto('/');
  await page
    .getByRole('button', {
      name: '空白文字 从一张干净的 A4 页面开始',
    })
    .click();

  const editor = page.getByRole('textbox', { name: '文档正文' });
  await expect(editor).toHaveAttribute('data-pagination-state', 'ready');
  await editor.fill('A3S Office deterministic pagination. '.repeat(1_200));
  await expect
    .poll(async () =>
      Number(await editor.getAttribute('data-pagination-pages')),
    )
    .toBeGreaterThan(1);
  const pageCount = await editor.getAttribute('data-pagination-pages');
  const engine = await editor.getAttribute('data-pagination-engine');
  const breakCount = await editor
    .locator('.work-document-auto-page-break')
    .count();
  expect(engine).toBe('wasm');
  expect(breakCount).toBeGreaterThan(0);
  await editor.evaluate((element) => {
    element.setAttribute('data-preview-identity', 'canonical');
  });

  await page.getByRole('button', { name: '预览' }).click();
  const preview = page.locator(
    '.work-document-preview-page .ProseMirror[data-preview-identity="canonical"]',
  );
  await expect(preview).toBeVisible();
  await expect(preview).toHaveAttribute('role', 'document');
  await expect(preview).toHaveAttribute('contenteditable', 'false');
  await expect(preview).toHaveAttribute('data-pagination-state', 'ready');
  await expect(preview).toHaveAttribute(
    'data-pagination-pages',
    pageCount ?? '',
  );
  await expect(preview).toHaveAttribute('data-pagination-engine', engine ?? '');
  await expect(preview.locator('.work-document-auto-page-break')).toHaveCount(
    breakCount,
  );
  expect(pageErrors).toEqual([]);
});

test('Word edit, preview, and PDF surfaces share one typography baseline', async ({
  page,
}) => {
  const pageErrors: Error[] = [];
  page.on('pageerror', (error) => pageErrors.push(error));
  await page.goto('/');
  await openDocumentFixture(page);
  await waitForDocumentFixture(page);

  const editing = await documentTypography(
    page.locator('.work-document-editable .ProseMirror'),
  );
  await page.getByRole('button', { name: '预览' }).click();
  const previewPage = page.locator('.work-document-preview-page');
  await expect(previewPage).toHaveCount(1);
  await expect(previewPage).toBeVisible();
  await expect(previewPage.locator(':scope > header')).toHaveCount(0);
  const preview = await documentTypography(previewPage);

  await page.evaluate(() => {
    const preview = document.querySelector<HTMLElement>(
      '.work-document-preview-page',
    );
    const body = preview?.querySelector<HTMLElement>('.ProseMirror');
    if (!body) throw new Error('Document preview body is unavailable.');
    const probe = document.createElement('section');
    probe.id = 'document-pdf-typography-probe';
    probe.className = 'work-pdf-export-page document a4 portrait';
    probe.style.position = 'fixed';
    probe.style.left = '-100000px';
    probe.innerHTML = body.innerHTML;
    document.body.append(probe);
  });
  const pdfProbe = page.locator('#document-pdf-typography-probe');
  const pdf = await documentTypography(pdfProbe);
  await pdfProbe.evaluate((element) => element.remove());

  expect(preview).toEqual(editing);
  expect(pdf).toEqual(editing);
  expect(pageErrors).toEqual([]);
});

test('Word preview keeps page chrome inside the physical page margins', async ({
  page,
}) => {
  const pageErrors: Error[] = [];
  page.on('pageerror', (error) => pageErrors.push(error));
  await page.goto('/');
  await openDocumentFixture(page);
  await waitForDocumentFixture(page);

  const editorHeader = page.locator('.work-document-page-header');
  await editorHeader.dblclick();
  const headerEditor = page.getByRole('textbox', { name: '页内页眉' });
  await headerEditor.fill('项目页眉');
  await page.keyboard.press('Escape');
  await expect(editorHeader).toContainText('项目页眉');

  const editorFooter = page.locator('.work-document-page-footer');
  await editorFooter.dblclick();
  const footerEditor = page.getByRole('textbox', { name: '页内页脚' });
  await footerEditor.fill('内部资料');
  await page.keyboard.press('Escape');
  await expect(editorFooter).toContainText('内部资料');

  await page.getByRole('button', { name: '预览' }).click();
  const previewPage = page.locator('.work-document-preview-page');
  await expect(previewPage).toHaveCount(1);
  const previewHeader = previewPage.locator(':scope > header');
  const previewFooter = previewPage.locator(':scope > footer');
  await expect(previewHeader).toContainText('项目页眉');
  await expect(previewFooter).toContainText('内部资料');

  const geometry = await previewPage.evaluate((paper) => {
    const header = paper.querySelector<HTMLElement>(':scope > header');
    const body = paper.querySelector<HTMLElement>(
      ':scope > .work-document-editable',
    );
    const footer = paper.querySelector<HTMLElement>(':scope > footer');
    if (!(header && body && footer)) {
      throw new Error('Document preview page chrome is incomplete.');
    }
    const paperRect = paper.getBoundingClientRect();
    const headerRect = header.getBoundingClientRect();
    const bodyRect = body.getBoundingClientRect();
    const footerRect = footer.getBoundingClientRect();
    const paperStyle = getComputedStyle(paper);
    return {
      headerPosition: getComputedStyle(header).position,
      footerPosition: getComputedStyle(footer).position,
      pageTop: paperRect.top,
      pageBottom: paperRect.bottom,
      headerTop: headerRect.top,
      headerBottom: headerRect.bottom,
      headerLeft: headerRect.left,
      headerRight: headerRect.right,
      bodyTop: bodyRect.top,
      bodyBottom: bodyRect.bottom,
      bodyLeft: bodyRect.left,
      bodyRight: bodyRect.right,
      footerTop: footerRect.top,
      footerBottom: footerRect.bottom,
      footerLeft: footerRect.left,
      footerRight: footerRect.right,
      scale: paper.offsetWidth > 0 ? paperRect.width / paper.offsetWidth : 1,
      paddingTop: Number.parseFloat(paperStyle.paddingTop),
      paddingBottom: Number.parseFloat(paperStyle.paddingBottom),
    };
  });
  expect(geometry.headerPosition).toBe('absolute');
  expect(geometry.footerPosition).toBe('absolute');
  expect(Math.abs(geometry.headerTop - geometry.pageTop)).toBeLessThanOrEqual(
    2,
  );
  expect(
    Math.abs(
      geometry.headerBottom -
        (geometry.pageTop + geometry.paddingTop * geometry.scale),
    ),
  ).toBeLessThanOrEqual(2);
  expect(
    Math.abs(geometry.bodyTop - geometry.headerBottom),
  ).toBeLessThanOrEqual(2);
  expect(Math.abs(geometry.headerLeft - geometry.bodyLeft)).toBeLessThanOrEqual(
    2,
  );
  expect(
    Math.abs(geometry.headerRight - geometry.bodyRight),
  ).toBeLessThanOrEqual(2);
  expect(
    Math.abs(
      geometry.footerTop -
        (geometry.pageBottom - geometry.paddingBottom * geometry.scale),
    ),
  ).toBeLessThanOrEqual(2);
  expect(geometry.footerTop).toBeGreaterThanOrEqual(geometry.bodyBottom);
  expect(
    Math.abs(geometry.footerBottom - geometry.pageBottom),
  ).toBeLessThanOrEqual(2);
  expect(Math.abs(geometry.footerLeft - geometry.bodyLeft)).toBeLessThanOrEqual(
    2,
  );
  expect(
    Math.abs(geometry.footerRight - geometry.bodyRight),
  ).toBeLessThanOrEqual(2);
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

async function documentTypography(surface: Locator) {
  return surface.evaluate((root) => {
    const typographyFor = (element: HTMLElement) => {
      const style = getComputedStyle(element);
      return {
        color: style.color,
        fontFamily: style.fontFamily,
        fontSize: style.fontSize,
        fontStyle: style.fontStyle,
        fontSynthesis: style.fontSynthesis,
        fontWeight: style.fontWeight,
        letterSpacing: style.letterSpacing,
        lineHeight: style.lineHeight,
        overflowWrap: style.overflowWrap,
      };
    };
    const blockStyleFor = (selector: string) => {
      const element = root.querySelector<HTMLElement>(selector);
      if (!element) {
        throw new Error(`Document typography element is missing: ${selector}`);
      }
      const style = getComputedStyle(element);
      return {
        ...typographyFor(element),
        marginBottom: style.marginBottom,
        marginLeft: style.marginLeft,
        marginRight: style.marginRight,
        marginTop: style.marginTop,
        paddingBottom: style.paddingBottom,
        paddingLeft: style.paddingLeft,
        paddingRight: style.paddingRight,
        paddingTop: style.paddingTop,
      };
    };
    return {
      body: typographyFor(root as HTMLElement),
      heading1: blockStyleFor('h1'),
      heading2: blockStyleFor('h2'),
      paragraph: blockStyleFor('p'),
      blockquote: blockStyleFor('blockquote'),
      unorderedList: blockStyleFor('ul'),
      orderedList: blockStyleFor('ol'),
    };
  });
}
