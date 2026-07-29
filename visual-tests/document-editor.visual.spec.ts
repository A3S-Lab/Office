import { expect, test, type Locator, type Page } from '@playwright/test';
import {
  openDocumentFixture,
  stabilizeVisualSurface,
  waitForDocumentFixture,
} from './visual-test-support';

interface BrowserPageErrorDiagnostic {
  message: string;
  stack: string | null;
  causeMessage: string | null;
  causeStack: string | null;
}

const fixtures = [
  {
    kind: 'document' as const,
    open: openDocumentFixture,
    ready: waitForDocumentFixture,
  },
];

test('document styles stay visible, semantic, and compact when needed', async ({
  page,
}) => {
  const fixture = fixtures.find((candidate) => candidate.kind === 'document');
  if (!fixture) throw new Error('Missing document visual fixture.');

  await page.setViewportSize({ width: 1280, height: 800 });
  await page.goto('/');
  await fixture.open(page);
  await fixture.ready(page);

  const gallery = page.getByRole('radiogroup', { name: '段落样式库' });
  const compactSelect = page.getByRole('combobox', { name: '段落样式' });
  await expect(gallery).toBeVisible();
  await expect(compactSelect).toBeHidden();
  const galleryGeometry = await gallery.evaluate((element) => {
    const rect = element.getBoundingClientRect();
    return { width: rect.width, height: rect.height };
  });
  expect(galleryGeometry.width).toBeGreaterThanOrEqual(250);
  expect(galleryGeometry.width).toBeLessThanOrEqual(270);
  expect(galleryGeometry.height).toBeLessThanOrEqual(50);

  const paragraph = page
    .locator('.work-document-editable .ProseMirror p')
    .nth(1);
  const paragraphText = (await paragraph.textContent())?.trim();
  if (!paragraphText) {
    throw new Error('Document style fixture paragraph is unavailable.');
  }
  await paragraph.click();
  const headingStyle = gallery.getByRole('radio', {
    name: '应用样式：标题 2',
  });
  await headingStyle.click();
  const styledHeading = page
    .locator('.work-document-editable .ProseMirror h2')
    .filter({ hasText: paragraphText });
  await expect(styledHeading).toHaveCount(1);
  await expect(headingStyle).toBeChecked();

  await headingStyle.click();
  await expect(styledHeading).toHaveCount(1);

  await page.setViewportSize({ width: 768, height: 900 });
  await expect(gallery).toBeHidden();
  await compactSelect.scrollIntoViewIfNeeded();
  await expect(compactSelect).toBeVisible();
  await expect(compactSelect).toContainText('标题 2');
});

test('document list libraries keep styles and numbering settings in context', async ({
  page,
}) => {
  const fixture = fixtures.find((candidate) => candidate.kind === 'document');
  if (!fixture) throw new Error('Missing document visual fixture.');

  await page.goto('/');
  await fixture.open(page);
  await fixture.ready(page);

  const bulletItem = page
    .locator('.work-document-editable .ProseMirror ul > li > p')
    .first();
  await expect(bulletItem).toBeVisible();
  await bulletItem.click();
  const bulletTrigger = page.getByRole('button', { name: '项目符号库' });
  await bulletTrigger.click();
  const bulletLibrary = page.getByRole('dialog', { name: '项目符号库' });
  await expect(bulletLibrary).toBeVisible();
  await expect(
    bulletLibrary.getByRole('menuitemradio', { name: '实心圆点' }),
  ).toBeFocused();
  await page.keyboard.press('ArrowRight');
  await expect(
    bulletLibrary.getByRole('menuitemradio', { name: '空心圆点' }),
  ).toBeFocused();
  await page.keyboard.press('Enter');
  await expect(bulletLibrary).toBeHidden();
  await expect(page.getByRole('textbox', { name: '文档正文' })).toBeFocused();
  await expect(bulletItem.locator('xpath=ancestor::ul[1]')).toHaveAttribute(
    'data-office-bullet-style',
    'circle',
  );

  const orderedItem = page
    .locator('.work-document-editable .ProseMirror ol > li > p')
    .first();
  await expect(orderedItem).toBeVisible();
  await orderedItem.click();
  const numberingTrigger = page.getByRole('button', { name: '编号库' });
  await numberingTrigger.click();
  const numberingLibrary = page.getByRole('dialog', { name: '编号库' });
  await expect(numberingLibrary).toBeVisible();
  await expect(numberingLibrary.getByRole('menuitemradio')).toHaveCount(5);
  await expect(
    numberingLibrary.getByRole('button', { name: '继续前一列表' }),
  ).toBeDisabled();

  const geometry = await numberingLibrary.evaluate((element) => {
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

  const start = numberingLibrary.getByRole('textbox', { name: '起始编号' });
  await start.fill('4');
  await numberingLibrary.getByRole('button', { name: '应用起始值' }).click();
  await expect(numberingLibrary).toBeHidden();
  await expect(page.getByRole('textbox', { name: '文档正文' })).toBeFocused();
  await expect(orderedItem.locator('xpath=ancestor::ol[1]')).toHaveAttribute(
    'start',
    '4',
  );

  await numberingTrigger.click();
  await expect(numberingLibrary).toBeVisible();
  await stabilizeVisualSurface(page);
  await expect(page).toHaveScreenshot('document-numbering-library.png');
});

test('document navigation keeps a live outline beside the editing surface', async ({
  page,
}, testInfo) => {
  const fixture = fixtures.find((candidate) => candidate.kind === 'document');
  if (!fixture) throw new Error('Missing document visual fixture.');

  await page.goto('/');
  await fixture.open(page);
  await fixture.ready(page);
  await page.getByRole('tab', { name: '视图' }).click();
  const navigationToggle = page.getByRole('button', { name: '导航窗格' });
  await navigationToggle.click();

  const pane = page.getByRole('complementary', { name: '文档导航' });
  const search = pane.getByRole('searchbox', { name: '搜索标题' });
  await expect(pane).toBeVisible();
  await expect(search).toBeFocused();
  await expect(pane.getByText('5 个标题')).toBeVisible();
  await expect(pane.getByRole('button', { name: '背景与目标' })).toBeVisible();

  await search.fill('范围');
  await expect(pane.getByText('1 个匹配')).toBeVisible();
  await expect(pane.getByRole('button', { name: '工作范围' })).toBeVisible();
  await search.fill('');
  await pane.getByRole('button', { name: '背景与目标' }).click();
  await expect(
    pane.getByRole('button', { name: '背景与目标' }),
  ).toHaveAttribute('aria-current', 'location');

  const geometry = await page.evaluate(() => {
    const workspace = document.querySelector('.work-document-workspace');
    const pane = document.querySelector('.work-document-navigation-panel');
    const scroll = document.querySelector('.work-document-scroll');
    if (!(workspace && pane && scroll)) {
      throw new Error('Document navigation geometry is unavailable.');
    }
    const workspaceRect = workspace.getBoundingClientRect();
    const paneRect = pane.getBoundingClientRect();
    const scrollRect = scroll.getBoundingClientRect();
    return {
      position: getComputedStyle(pane).position,
      workspaceLeft: workspaceRect.left,
      paneLeft: paneRect.left,
      paneRight: paneRect.right,
      paneWidth: paneRect.width,
      paneBottom: paneRect.bottom,
      scrollLeft: scrollRect.left,
      viewportHeight: document.documentElement.clientHeight,
    };
  });
  expect(geometry.paneLeft).toBeCloseTo(geometry.workspaceLeft, 0);
  expect(geometry.paneWidth).toBeGreaterThanOrEqual(224);
  expect(geometry.paneWidth).toBeLessThanOrEqual(292);
  expect(geometry.paneBottom).toBeLessThanOrEqual(geometry.viewportHeight);
  if (testInfo.project.name === 'compact-768') {
    expect(geometry.position).toBe('absolute');
    expect(geometry.scrollLeft).toBeCloseTo(geometry.workspaceLeft, 0);
  } else {
    expect(geometry.position).toBe('static');
    expect(geometry.scrollLeft).toBeCloseTo(geometry.paneRight, 0);
  }

  await stabilizeVisualSurface(page);
  await expect(pane).toHaveScreenshot('document-navigation-pane.png');
  await page.keyboard.press('Escape');
  await expect(pane).toBeHidden();
  await expect(page.getByRole('textbox', { name: '文档正文' })).toBeFocused();
});

test('document selection toolbar keeps formatting and review in context', async ({
  page,
}) => {
  const pageErrors: Error[] = [];
  page.on('pageerror', (error) => pageErrors.push(error));
  await page.addInitScript(() => {
    const state = window as Window & {
      __a3sOfficePageErrors?: BrowserPageErrorDiagnostic[];
    };
    state.__a3sOfficePageErrors = [];
    window.addEventListener(
      'error',
      (event) => {
        const error = event.error as
          | (Error & { cause?: unknown })
          | null
          | undefined;
        const cause = error?.cause;
        state.__a3sOfficePageErrors?.push({
          message: error?.message ?? event.message,
          stack: error?.stack ?? null,
          causeMessage:
            cause instanceof Error
              ? cause.message
              : cause == null
                ? null
                : String(cause),
          causeStack: cause instanceof Error ? (cause.stack ?? null) : null,
        });
      },
      true,
    );
  });
  const fixture = fixtures.find((candidate) => candidate.kind === 'document');
  if (!fixture) throw new Error('Missing document visual fixture.');

  await page.goto('/');
  await fixture.open(page);
  await fixture.ready(page);
  await expectNoPageErrors(
    page,
    pageErrors,
    'opening the document must not raise page errors',
  );
  await selectDocumentParagraphText(page, 1, 12);

  const toolbar = page.getByRole('toolbar', {
    name: '文本快捷工具栏',
  });
  await expect(toolbar).toBeVisible();
  const toolbarGeometry = await toolbar.evaluate((element) => {
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
  expect(toolbarGeometry.left).toBeGreaterThanOrEqual(8);
  expect(toolbarGeometry.top).toBeGreaterThanOrEqual(8);
  expect(toolbarGeometry.right).toBeLessThanOrEqual(
    toolbarGeometry.viewportWidth - 8,
  );
  expect(toolbarGeometry.bottom).toBeLessThanOrEqual(
    toolbarGeometry.viewportHeight - 8,
  );
  const fontFamily = toolbar.getByRole('combobox', { name: '快捷字体' });
  await expect(fontFamily).toBeVisible();
  const fontSize = toolbar.getByRole('combobox', { name: '快捷字号' });
  await expect(fontSize).toBeVisible();
  await expect(fontSize).toContainText('10.5');
  expect(
    await fontSize
      .locator('span')
      .evaluate((label) => label.scrollWidth <= label.clientWidth),
  ).toBe(true);
  await expectNoPageErrors(
    page,
    pageErrors,
    'showing the selection toolbar must not raise page errors',
  );

  await fontFamily.click();
  const fontMenu = page.getByRole('listbox', { name: '快捷字体' });
  await expect(fontMenu).toBeVisible();
  await expectFloatingSurfaceWithinViewport(fontMenu);
  await page.keyboard.press('Escape');
  await expect(fontMenu).toBeHidden();
  await expect(toolbar).toBeVisible();
  await expectNoPageErrors(
    page,
    pageErrors,
    'font menu interaction must not raise page errors',
  );

  await toolbar.getByRole('button', { name: '快捷文字颜色' }).click();
  const colorMenu = page.getByRole('dialog', { name: '快捷文字颜色' });
  await expect(colorMenu).toBeVisible();
  await expectFloatingSurfaceWithinViewport(colorMenu);
  await page.keyboard.press('Escape');
  await expect(colorMenu).toBeHidden();
  await expect(toolbar).toBeVisible();
  await expectNoPageErrors(
    page,
    pageErrors,
    'color menu interaction must not raise page errors',
  );

  await page.getByRole('tab', { name: '审阅' }).click();
  await expect(toolbar).toBeHidden();
  await expectNoPageErrors(
    page,
    pageErrors,
    'toolbar dismissal must not raise page errors',
  );
  await page.getByRole('tab', { name: '开始' }).click();
  await page.getByRole('textbox', { name: '文档正文' }).focus();
  await selectDocumentParagraphText(page, 1, 12);
  await expect(toolbar).toBeVisible();

  const selectedText = await page.evaluate(
    () => window.getSelection()?.toString() ?? '',
  );
  expect(selectedText.length).toBeGreaterThan(0);
  const bold = toolbar.getByRole('button', { name: '加粗' });
  await bold.click();
  await expect(bold).toHaveAttribute('aria-pressed', 'true');
  await expect
    .poll(() =>
      page.locator('.work-document-editable .ProseMirror').evaluate((root) => {
        const selection = window.getSelection();
        if (!selection || selection.rangeCount === 0 || selection.isCollapsed)
          return false;
        const range = selection.getRangeAt(0);
        return [...root.querySelectorAll('strong')].some(
          (element) =>
            range.intersectsNode(element) &&
            element.textContent?.includes(selection.toString()),
        );
      }),
    )
    .toBe(true);
  await expect
    .poll(() => page.evaluate(() => window.getSelection()?.toString() ?? ''))
    .toBe(selectedText);
  await expectNoPageErrors(
    page,
    pageErrors,
    'selection formatting must not raise page errors',
  );

  await toolbar.getByRole('button', { name: '添加批注' }).click();
  const composer = page.getByRole('dialog', { name: '添加批注' });
  await expect(composer).toBeVisible();
  await expect(page.locator('[data-document-comment-draft]')).toHaveCount(1);
  await page.keyboard.press('Escape');
  await expect(composer).toBeHidden();
  await expect(page.locator('[data-document-comment-draft]')).toHaveCount(0);
  await expectNoPageErrors(
    page,
    pageErrors,
    'comment composer interaction must not raise page errors',
  );
});

test('document comments align with their review rail', async ({
  page,
}, testInfo) => {
  const fixture = fixtures.find((candidate) => candidate.kind === 'document');
  if (!fixture) throw new Error('Missing document visual fixture.');

  await page.clock.setFixedTime(new Date('2026-07-29T08:00:00.000Z'));
  await page.goto('/');
  await fixture.open(page);
  await fixture.ready(page);
  const selectedRange = await selectDocumentParagraphText(page, 1, 12);
  await page.getByRole('tab', { name: '审阅' }).click();
  await page.getByRole('button', { name: '添加批注' }).click();
  const composer = page.getByRole('dialog', { name: '添加批注' });
  await expect(page.locator('.ds-dialog-backdrop')).toHaveCount(0);
  await expect(composer).toHaveClass(/work-document-comment-composer/);
  const composerGeometry = await composer.evaluate((element) => {
    const rect = element.getBoundingClientRect();
    const panel = element.closest('.work-document-comments-panel');
    const panelRect = panel?.getBoundingClientRect();
    return {
      top: rect.top,
      left: rect.left,
      right: rect.right,
      panelLeft: panelRect?.left ?? 0,
      panelRight: panelRect?.right ?? 0,
    };
  });
  expect(composerGeometry.left).toBeGreaterThanOrEqual(
    composerGeometry.panelLeft + 8,
  );
  expect(composerGeometry.right).toBeLessThanOrEqual(
    composerGeometry.panelRight - 8,
  );
  expect(Math.abs(composerGeometry.top - selectedRange.top)).toBeLessThan(72);

  await composer
    .getByRole('textbox', { name: '批注内容' })
    .fill('这里需要补充可衡量的验收标准。');
  await composer.getByRole('button', { name: '添加批注', exact: true }).click();

  const panel = page.getByRole('complementary', { name: '批注审阅' });
  const mark = page.locator('[data-document-comment]');
  const connector = page.locator('.work-document-comment-connectors path');
  await expect(panel).toBeVisible();
  await expect(mark).toHaveCount(1);
  await expect(mark).toHaveClass(/is-active-comment/);
  await expect(connector).toHaveCount(1);
  await expect(page.getByText('这里需要补充可衡量的验收标准。')).toBeVisible();

  const geometry = await page.evaluate(() => {
    const panel = document.querySelector<HTMLElement>(
      '.work-document-comments-panel',
    );
    const pageElement = document.querySelector<HTMLElement>(
      '.work-document-page',
    );
    if (!panel || !pageElement) {
      throw new Error('Document review geometry is unavailable.');
    }
    const panelRect = panel.getBoundingClientRect();
    const pageRect = pageElement.getBoundingClientRect();
    return {
      viewportWidth: document.documentElement.clientWidth,
      panelLeft: panelRect.left,
      panelRight: panelRect.right,
      panelWidth: panelRect.width,
      pageRight: pageRect.right,
    };
  });
  if (testInfo.project.name === 'compact-768') {
    expect(geometry.panelLeft).toBeGreaterThanOrEqual(
      geometry.viewportWidth - geometry.panelWidth - 16,
    );
  } else {
    expect(geometry.panelLeft).toBeGreaterThanOrEqual(geometry.pageRight - 1);
  }
  expect(geometry.panelRight).toBeLessThanOrEqual(geometry.viewportWidth + 1);
  expect(geometry.panelWidth).toBeGreaterThanOrEqual(270);

  await stabilizeVisualSurface(page);
  await expect(page).toHaveScreenshot('document-comments.png');

  const reply = panel.getByRole('textbox', { name: '回复批注 1' });
  await reply.fill('尚未发送的回复');
  await reply.press('Escape');
  const discardReply = page.getByRole('dialog', {
    name: '放弃未完成的批注？',
  });
  await expect(discardReply).toBeVisible();
  await discardReply.getByRole('button', { name: '取消' }).click();
  await expect(reply).toBeFocused();
  await expect(reply).toHaveValue('尚未发送的回复');

  await reply.press('Escape');
  await discardReply.getByRole('button', { name: '放弃内容' }).click();
  await expect(panel).toBeHidden();
  await expect(page.getByRole('textbox', { name: '文档正文' })).toBeFocused();
});

test('document comment drafts clean up and stacked comments do not overlap', async ({
  page,
}) => {
  const fixture = fixtures.find((candidate) => candidate.kind === 'document');
  if (!fixture) throw new Error('Missing document visual fixture.');

  await page.goto('/');
  await fixture.open(page);
  await fixture.ready(page);
  await selectDocumentParagraphText(page, 1, 12);
  await page.getByRole('tab', { name: '审阅' }).click();

  const toolbarAddComment = page
    .getByRole('region', { name: '文字功能区' })
    .getByRole('button', { name: '添加批注', exact: true });
  await expect(toolbarAddComment).toBeEnabled();
  await toolbarAddComment.click();
  await expect(page.locator('[data-document-comment-draft]')).toHaveCount(1);
  await expect(page.getByRole('dialog', { name: '添加批注' })).toBeVisible();

  await page.keyboard.press('Escape');
  await expect(page.getByRole('dialog', { name: '添加批注' })).toHaveCount(0);
  await expect(page.locator('[data-document-comment-draft]')).toHaveCount(0);
  await expect(page.getByRole('textbox', { name: '文档正文' })).toBeFocused();

  await selectDocumentParagraphText(page, 1, 12);
  await expect(toolbarAddComment).toBeEnabled();
  await toolbarAddComment.click();
  let composer = page.getByRole('dialog', { name: '添加批注' });
  await composer.getByRole('textbox', { name: '批注内容' }).fill('第一条批注');
  await composer
    .getByRole('textbox', { name: '批注内容' })
    .press('Control+Enter');
  await expect(page.locator('[data-document-comment]')).toHaveCount(1);

  await selectDocumentParagraphText(page, 2, 10);
  await expect(toolbarAddComment).toBeEnabled();
  await toolbarAddComment.click();
  composer = page.getByRole('dialog', { name: '添加批注' });
  await composer.getByRole('textbox', { name: '批注内容' }).fill('第二条批注');
  await composer.getByRole('button', { name: '添加批注', exact: true }).click();

  const cards = page.locator(
    '.work-document-comment-track > article[data-comment-id]',
  );
  await expect(cards).toHaveCount(2);
  await expect(
    page.locator('.work-document-comment-connectors path'),
  ).toHaveCount(2);
  const cardGeometry = await cards.evaluateAll((elements) =>
    elements
      .map((element) => {
        const rect = element.getBoundingClientRect();
        return { top: rect.top, bottom: rect.bottom };
      })
      .sort((left, right) => left.top - right.top),
  );
  expect(cardGeometry[1].top).toBeGreaterThanOrEqual(
    cardGeometry[0].bottom + 8,
  );

  const alignedOffsetBeforeScroll = await documentCommentAlignmentOffset(page);
  await page.locator('.work-document-scroll').evaluate((element) => {
    element.scrollTop += 120;
    element.dispatchEvent(new Event('scroll'));
  });
  await page.evaluate(
    () =>
      new Promise<void>((resolve) =>
        requestAnimationFrame(() => requestAnimationFrame(() => resolve())),
      ),
  );
  const alignedOffsetAfterScroll = await documentCommentAlignmentOffset(page);
  expect(alignedOffsetAfterScroll).toBeCloseTo(alignedOffsetBeforeScroll, 0);
});

test('document task panes and dialogs preserve the editing context', async ({
  page,
}) => {
  const fixture = fixtures.find((candidate) => candidate.kind === 'document');
  if (!fixture) throw new Error('Missing document visual fixture.');

  await page.goto('/');
  await fixture.open(page);
  await fixture.ready(page);

  await page.getByRole('tab', { name: '开始' }).click();
  await page.getByRole('button', { name: '查找', exact: true }).click();
  const findPane = page.getByRole('complementary', { name: '查找' });
  await expect(findPane).toBeVisible();
  await expect(
    findPane.getByRole('textbox', { name: '查找内容' }),
  ).toBeFocused();
  const paneGeometry = await findPane.evaluate((element) => {
    const rect = element.getBoundingClientRect();
    return {
      viewportWidth: document.documentElement.clientWidth,
      left: rect.left,
      right: rect.right,
      width: rect.width,
    };
  });
  expect(paneGeometry.width).toBeGreaterThanOrEqual(300);
  expect(paneGeometry.width).toBeLessThanOrEqual(380);
  expect(paneGeometry.right).toBeLessThanOrEqual(
    paneGeometry.viewportWidth + 1,
  );

  await page.keyboard.press('Escape');
  await expect(findPane).toBeHidden();
  await expect(page.getByRole('textbox', { name: '文档正文' })).toBeFocused();

  await page.getByRole('tab', { name: '页面布局' }).click();
  const paragraphSpacingTrigger = page.getByRole('button', {
    name: '段落间距',
  });
  await paragraphSpacingTrigger.click();
  const paragraphSpacing = page.getByRole('dialog', {
    name: '段落间距选项',
  });
  await expect(paragraphSpacing).toBeVisible();
  await expect(
    paragraphSpacing.getByRole('textbox', { name: '段前间距（磅）' }),
  ).toBeFocused();
  const popoverGeometry = await paragraphSpacing.evaluate((element) => {
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
  expect(popoverGeometry.left).toBeGreaterThanOrEqual(8);
  expect(popoverGeometry.top).toBeGreaterThanOrEqual(8);
  expect(popoverGeometry.right).toBeLessThanOrEqual(
    popoverGeometry.viewportWidth - 8,
  );
  expect(popoverGeometry.bottom).toBeLessThanOrEqual(
    popoverGeometry.viewportHeight - 8,
  );
  await page.keyboard.press('Escape');
  await expect(paragraphSpacing).toBeHidden();
  await expect(paragraphSpacingTrigger).toBeFocused();

  await page.getByRole('tab', { name: '引用' }).click();
  await page.getByRole('button', { name: '插入图片题注' }).click();
  const captionDialog = page.getByRole('dialog', { name: '插入图片题注' });
  const captionField = captionDialog.getByRole('textbox', {
    name: '题注文字',
  });
  await expect(captionDialog).toBeVisible();
  await expect(captionField).toBeFocused();
  await expect(page.locator('body > [inert]')).not.toHaveCount(0);
  const dialogGeometry = await captionDialog.evaluate((element) => {
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
  expect(dialogGeometry.left).toBeGreaterThanOrEqual(8);
  expect(dialogGeometry.top).toBeGreaterThanOrEqual(8);
  expect(dialogGeometry.right).toBeLessThanOrEqual(
    dialogGeometry.viewportWidth - 8,
  );
  expect(dialogGeometry.bottom).toBeLessThanOrEqual(
    dialogGeometry.viewportHeight - 8,
  );
  await captionField.fill('系统架构');
  await captionField.press('Enter');
  await expect(captionDialog).toBeHidden();
  await expect(page.getByRole('textbox', { name: '文档正文' })).toBeFocused();

  await page.getByRole('tab', { name: '插入' }).click();
  await page.getByRole('button', { name: '添加链接' }).click();
  const linkDialog = page.getByRole('dialog', { name: '添加链接' });
  const linkField = linkDialog.getByRole('textbox', { name: '链接地址' });
  const addLink = linkDialog.getByRole('button', { name: '添加链接' });
  await expect(linkField).toHaveValue('https://');
  await expect(addLink).toBeDisabled();
  await linkField.fill('https://a3s.dev/docs');
  await expect(addLink).toBeEnabled();
  await addLink.click();
  await expect(linkDialog).toBeHidden();
  await expect(page.getByRole('textbox', { name: '文档正文' })).toBeFocused();

  await page.getByRole('tab', { name: '引用' }).click();
  await page.getByRole('button', { name: '文献库' }).click();
  const citationsPane = page.getByRole('complementary', { name: '文献库' });
  await expect(citationsPane).toBeVisible();
  await expect(
    citationsPane.getByRole('form', { name: '新建文献' }),
  ).toBeVisible();
  await expect(
    citationsPane.getByRole('textbox', { name: '出版者' }),
  ).toBeHidden();
  await citationsPane.getByText('更多出版信息').click();
  await expect(
    citationsPane.getByRole('textbox', { name: '出版者' }),
  ).toBeVisible();
  await citationsPane
    .getByRole('textbox', { name: '文献简称' })
    .fill('a3s2026');
  await citationsPane
    .getByRole('textbox', { name: '文献标题' })
    .fill('A3S Office');
  await citationsPane.getByRole('button', { name: '保存文献' }).click();
  await expect(
    citationsPane.getByRole('form', { name: '编辑文献' }),
  ).toBeVisible();
  await expect(
    citationsPane.getByRole('button', { name: '插入引文' }),
  ).toBeEnabled();

  const citationTitle = citationsPane.getByRole('textbox', {
    name: '文献标题',
  });
  await citationTitle.fill('A3S Office edited');
  const closeCitations = citationsPane.getByRole('button', {
    name: '关闭文献库',
  });
  await closeCitations.click();
  const discardCitation = page.getByRole('dialog', {
    name: '放弃未保存的文献更改？',
  });
  await discardCitation.getByRole('button', { name: '取消' }).click();
  await expect(closeCitations).toBeFocused();
  await expect(citationTitle).toHaveValue('A3S Office edited');

  await closeCitations.click();
  await discardCitation.getByRole('button', { name: '放弃更改' }).click();
  await expect(citationsPane).toBeHidden();
  await expect(page.getByRole('textbox', { name: '文档正文' })).toBeFocused();
});

test('document revision decisions require an explicit confirmation', async ({
  page,
}) => {
  const fixture = fixtures.find((candidate) => candidate.kind === 'document');
  if (!fixture) throw new Error('Missing document visual fixture.');

  await page.goto('/');
  await fixture.open(page);
  await fixture.ready(page);

  await page.getByRole('tab', { name: '审阅' }).click();
  await page.getByRole('button', { name: '修订模式' }).click();
  const editor = page.getByRole('textbox', { name: '文档正文' });
  await editor.click();
  await page.keyboard.press('End');
  await page.keyboard.type(' tracked');
  await page.getByRole('button', { name: /查看修订/ }).click();

  const changesPane = page.getByRole('complementary', { name: '修订审阅' });
  await expect(changesPane).toBeVisible();
  const rejectAll = changesPane.getByRole('button', { name: '全部拒绝' });
  await expect(rejectAll).toBeVisible();
  await rejectAll.click();

  const confirmation = page.getByRole('dialog', {
    name: '拒绝全部修订？',
  });
  await expect(
    confirmation.getByRole('button', { name: '取消' }),
  ).toBeFocused();
  await confirmation.getByRole('button', { name: '取消' }).click();
  await expect(rejectAll).toBeFocused();

  await rejectAll.click();
  await confirmation.getByRole('button', { name: '全部拒绝' }).click();
  await expect(
    changesPane.getByText('开启修订后，改动会显示在这里。'),
  ).toBeVisible();
  await expect(
    changesPane.getByRole('button', { name: '全部拒绝' }),
  ).toBeHidden();
  await expect(editor).toBeFocused();
});

test('document review views remain usable at phone width', async ({ page }) => {
  const fixture = fixtures.find((candidate) => candidate.kind === 'document');
  if (!fixture) throw new Error('Missing document visual fixture.');

  await page.setViewportSize({ width: 390, height: 700 });
  await page.goto('/');
  await fixture.open(page);
  await fixture.ready(page);

  const compactChrome = await page
    .locator('.work-editor-shell.document')
    .evaluate((shell) => {
      const header = shell.querySelector<HTMLElement>(
        '.playground-editor-header',
      );
      const ribbonRow = shell.querySelector<HTMLElement>(
        '.work-office-ribbon-tabs-row',
      );
      const ribbonTabs = shell.querySelector<HTMLElement>(
        '.work-office-ribbon-tabs',
      );
      const tabs = Array.from(
        ribbonTabs?.querySelectorAll<HTMLElement>('[role="tab"]') ?? [],
      );
      if (!(header && ribbonRow && ribbonTabs) || tabs.length < 2) {
        throw new Error('Compact document command bar is incomplete.');
      }
      const headerRect = header.getBoundingClientRect();
      const ribbonRect = ribbonRow.getBoundingClientRect();
      return {
        headerBottom: headerRect.bottom,
        ribbonTop: ribbonRect.top,
        overflowX: getComputedStyle(ribbonTabs).overflowX,
        tabsClientWidth: ribbonTabs.clientWidth,
        tabsScrollWidth: ribbonTabs.scrollWidth,
        tabRects: tabs.map((tab) => {
          const rect = tab.getBoundingClientRect();
          return {
            left: rect.left,
            right: rect.right,
            flexShrink: getComputedStyle(tab).flexShrink,
          };
        }),
      };
    });
  expect(compactChrome.ribbonTop).toBeGreaterThanOrEqual(
    compactChrome.headerBottom - 1,
  );
  expect(compactChrome.overflowX).toBe('auto');
  expect(compactChrome.tabsScrollWidth).toBeGreaterThanOrEqual(
    compactChrome.tabsClientWidth,
  );
  for (const [index, tab] of compactChrome.tabRects.entries()) {
    expect(tab.flexShrink).toBe('0');
    if (index === 0) continue;
    expect(tab.left).toBeGreaterThanOrEqual(
      compactChrome.tabRects[index - 1].right - 1,
    );
  }

  const ribbonTabs = page.getByRole('tablist', { name: '文字功能区' });
  const homeTab = ribbonTabs.getByRole('tab', { name: '开始' });
  const insertTab = ribbonTabs.getByRole('tab', { name: '插入' });
  await homeTab.focus();
  await page.keyboard.press('ArrowRight');
  await expect(insertTab).toBeFocused();
  await expect(insertTab).toHaveAttribute('aria-selected', 'true');

  await page.getByRole('tab', { name: '页面布局' }).click();
  await page.getByRole('button', { name: '页面设置', exact: true }).click();
  const layoutPane = page.getByRole('complementary', { name: '页面设置' });
  const layoutGeometry = await layoutPane.evaluate((element) => {
    const rect = element.getBoundingClientRect();
    return {
      viewportWidth: document.documentElement.clientWidth,
      left: rect.left,
      right: rect.right,
      width: rect.width,
    };
  });
  expect(layoutGeometry.left).toBeGreaterThanOrEqual(0);
  expect(layoutGeometry.right).toBeLessThanOrEqual(
    layoutGeometry.viewportWidth + 1,
  );
  expect(layoutGeometry.width).toBeCloseTo(layoutGeometry.viewportWidth, 0);
  const layoutHeaderGeometry = await layoutPane
    .locator('.work-document-task-pane-header')
    .evaluate((element) => {
      const pane = element.parentElement;
      if (!pane) throw new Error('Page Setup task pane is unavailable.');
      const headerRect = element.getBoundingClientRect();
      const paneRect = pane.getBoundingClientRect();
      return {
        headerLeft: headerRect.left,
        headerRight: headerRect.right,
        headerWidth: headerRect.width,
        paneLeft: paneRect.left,
        paneRight: paneRect.right,
        paneWidth: paneRect.width,
      };
    });
  expect(
    Math.abs(layoutHeaderGeometry.headerLeft - layoutHeaderGeometry.paneLeft),
  ).toBeLessThanOrEqual(1);
  expect(
    Math.abs(layoutHeaderGeometry.headerRight - layoutHeaderGeometry.paneRight),
  ).toBeLessThanOrEqual(1);
  expect(
    Math.abs(layoutHeaderGeometry.headerWidth - layoutHeaderGeometry.paneWidth),
  ).toBeLessThanOrEqual(1);
  await expect(layoutPane.getByRole('tab', { name: '页面' })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  await expect(
    layoutPane.getByRole('textbox', { name: '默认页页眉' }),
  ).toHaveCount(0);
  await layoutPane.getByRole('tab', { name: '分栏与分节' }).click();
  await expect(
    layoutPane.getByRole('combobox', { name: '分节方式' }),
  ).toBeVisible();
  await layoutPane.getByRole('tab', { name: '页眉页脚' }).click();
  await expect(
    layoutPane.getByRole('textbox', { name: '默认页页眉' }),
  ).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(layoutPane).toBeHidden();

  await selectDocumentParagraphText(page, 1, 12);
  const selectionToolbar = page.getByRole('toolbar', {
    name: '文本快捷工具栏',
  });
  await expect(selectionToolbar).toBeVisible();
  const selectionToolbarGeometry = await selectionToolbar.evaluate(
    (element) => {
      const rect = element.getBoundingClientRect();
      return {
        viewportWidth: document.documentElement.clientWidth,
        left: rect.left,
        right: rect.right,
      };
    },
  );
  expect(selectionToolbarGeometry.left).toBeGreaterThanOrEqual(8);
  expect(selectionToolbarGeometry.right).toBeLessThanOrEqual(
    selectionToolbarGeometry.viewportWidth - 8,
  );
  await expect(
    selectionToolbar.getByRole('combobox', { name: '快捷字体' }),
  ).toBeHidden();
  const compactFontSize = selectionToolbar.getByRole('combobox', {
    name: '快捷字号',
  });
  await expect(compactFontSize).toContainText('10.5');
  expect(
    await compactFontSize
      .locator('span')
      .evaluate((label) => label.scrollWidth <= label.clientWidth),
  ).toBe(true);
  await expect(
    selectionToolbar.getByRole('button', { name: '清除格式' }),
  ).toBeHidden();

  await page.getByRole('tab', { name: '审阅' }).click();
  await page.getByRole('button', { name: '添加批注' }).click();
  const commentPanel = page.getByRole('complementary', { name: '批注审阅' });
  const composer = page.getByRole('dialog', { name: '添加批注' });
  const reviewGeometry = await commentPanel.evaluate((element) => {
    const rect = element.getBoundingClientRect();
    const connector = document.querySelector<HTMLElement>(
      '.work-document-comment-connectors',
    );
    const composerElement = element.querySelector<HTMLElement>(
      '.work-document-comment-composer',
    );
    return {
      viewportWidth: document.documentElement.clientWidth,
      panelLeft: rect.left,
      panelRight: rect.right,
      panelWidth: rect.width,
      connectorDisplay: connector ? getComputedStyle(connector).display : '',
      composerPosition: composerElement
        ? getComputedStyle(composerElement).position
        : '',
    };
  });
  expect(reviewGeometry.panelLeft).toBeGreaterThanOrEqual(0);
  expect(reviewGeometry.panelRight).toBeLessThanOrEqual(
    reviewGeometry.viewportWidth + 1,
  );
  expect(reviewGeometry.panelWidth).toBeGreaterThanOrEqual(350);
  expect(reviewGeometry.connectorDisplay).toBe('none');
  expect(reviewGeometry.composerPosition).toBe('relative');
  await expect(composer).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(composer).toBeHidden();
});

async function selectDocumentParagraphText(
  page: Page,
  paragraphIndex: number,
  characterCount: number,
): Promise<{ top: number; bottom: number; left: number; right: number }> {
  return page.locator('.work-document-editable .ProseMirror').evaluate(
    (root, options) => {
      const paragraph = root.querySelectorAll('p')[options.paragraphIndex];
      const text = paragraph?.firstChild;
      if (!(text instanceof Text)) {
        throw new Error('Document comment fixture text is unavailable.');
      }
      const range = document.createRange();
      range.setStart(text, 0);
      range.setEnd(text, Math.min(options.characterCount, text.length));
      const selection = window.getSelection();
      selection?.removeAllRanges();
      selection?.addRange(range);
      document.dispatchEvent(new Event('selectionchange', { bubbles: true }));
      const rect = range.getBoundingClientRect();
      return {
        top: rect.top,
        bottom: rect.bottom,
        left: rect.left,
        right: rect.right,
      };
    },
    { paragraphIndex, characterCount },
  );
}

async function expectNoPageErrors(
  page: Page,
  errors: Error[],
  context: string,
): Promise<void> {
  if (errors.length === 0) return;
  const browserDiagnostics = await page.evaluate(
    () =>
      (
        window as Window & {
          __a3sOfficePageErrors?: BrowserPageErrorDiagnostic[];
        }
      ).__a3sOfficePageErrors ?? [],
  );
  throw new Error(
    `${context}\n${JSON.stringify(
      {
        playwright: errors.map((error) => ({
          message: error.message,
          stack: error.stack ?? null,
        })),
        browser: browserDiagnostics,
      },
      null,
      2,
    )}`,
  );
}

async function expectFloatingSurfaceWithinViewport(
  surface: Locator,
): Promise<void> {
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

async function documentCommentAlignmentOffset(page: Page): Promise<number> {
  return page.evaluate(() => {
    const card = document.querySelector<HTMLElement>(
      '.work-document-comment-track > article[data-comment-id]',
    );
    const id = card?.dataset.commentId;
    const mark = [
      ...document.querySelectorAll<HTMLElement>(
        '[data-document-comment][data-comment-id]',
      ),
    ].find((candidate) => candidate.dataset.commentId === id);
    if (!card || !mark) {
      throw new Error('Document comment alignment geometry is unavailable.');
    }
    return card.getBoundingClientRect().top - mark.getBoundingClientRect().top;
  });
}
