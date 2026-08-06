import { expect, test } from '@playwright/test';
import {
  openDocumentFixture,
  waitForDocumentFixture,
} from './visual-test-support';

test('Word starts a blank document with an empty history', async ({ page }) => {
  await page.goto('/');
  await page
    .getByRole('button', {
      name: '空白文字 从一张干净的 A4 页面开始',
    })
    .click();
  await page.locator('.work-document-editable .ProseMirror').waitFor();

  const undo = page.getByRole('button', { name: '撤销' });
  const redo = page.getByRole('button', { name: '重做' });
  await expect(undo).toBeDisabled();
  await expect(redo).toBeDisabled();

  const body = page.getByRole('textbox', { name: '文档正文' });
  await body.fill('History starts here');
  await expect(undo).toBeEnabled();
  await expect(redo).toBeDisabled();

  await page.keyboard.press('Meta+z');
  await expect(body).toBeEmpty();
  await expect(undo).toBeDisabled();
  await expect(redo).toBeEnabled();
});

test('Word paragraph pagination menu updates live and restores style defaults', async ({
  page,
}) => {
  await page.goto('/');
  await openDocumentFixture(page);
  await waitForDocumentFixture(page);

  await page.getByRole('tab', { name: '页面布局' }).click();
  const trigger = page.getByRole('button', { name: '段落分页' });
  await expect(trigger).toHaveAttribute('aria-pressed', 'false');
  await trigger.click();

  const dialog = page.getByRole('dialog', { name: '段落分页选项' });
  const keepLines = dialog.getByRole('checkbox', { name: '段落不跨页' });
  const keepWithNext = dialog.getByRole('checkbox', { name: '与下一段同页' });
  const widowControl = dialog.getByRole('checkbox', {
    name: '避免页首、页尾单行',
  });
  const reset = dialog.getByRole('button', { name: '恢复默认分页规则' });
  await expect(keepLines).toBeFocused();
  await expect(widowControl).toBeChecked();
  await expect(reset).toBeDisabled();

  await page.keyboard.press('Tab');
  await expect(keepWithNext).toBeFocused();
  await expect(dialog).toBeVisible();

  await keepLines.click();
  await expect(keepLines).toBeChecked();
  await expect(trigger).toHaveAttribute('aria-pressed', 'true');
  await expect(reset).toBeEnabled();

  await reset.click();
  await expect(keepLines).not.toBeChecked();
  await expect(widowControl).toBeChecked();
  await expect(trigger).toHaveAttribute('aria-pressed', 'false');
  await expect(reset).toBeDisabled();

  await page.keyboard.press('Escape');
  await expect(trigger).toBeFocused();
});

test('Word paragraph spacing cancels a dirty draft before closing', async ({
  page,
}) => {
  await page.goto('/');
  await openDocumentFixture(page);
  await waitForDocumentFixture(page);

  await page.getByRole('tab', { name: '页面布局' }).click();
  const trigger = page.getByRole('button', { name: '段落间距' });
  await trigger.click();
  const dialog = page.getByRole('dialog', { name: '段落间距选项' });
  const before = dialog.getByRole('textbox', { name: '段前间距（磅）' });
  await expect(before).toBeFocused();

  await before.fill('18');
  await page.keyboard.press('Escape');
  await expect(dialog).toBeVisible();
  await expect(before).toHaveValue('');
  await expect(trigger).toHaveAttribute('aria-pressed', 'false');

  await page.keyboard.press('Escape');
  await expect(dialog).toBeHidden();
  await expect(trigger).toBeFocused();

  await trigger.click();
  await expect(before).toHaveValue('');
});

test('Word transient ribbon commands return keyboard control to the document', async ({
  page,
}) => {
  await page.goto('/');
  await openDocumentFixture(page);
  await waitForDocumentFixture(page);

  const body = page.getByRole('textbox', { name: '文档正文' });
  await body.focus();

  await page.getByRole('tab', { name: '视图' }).click();
  const viewModes = page.getByRole('region', { name: '文档视图' });
  await viewModes.getByRole('button', { name: '网页视图' }).click();
  await expect(body).toBeFocused();
  await viewModes.getByRole('button', { name: '页面视图' }).click();
  await expect(body).toBeFocused();

  await page.getByRole('button', { name: '标尺', exact: true }).click();
  await expect(body).toBeFocused();
  await page.getByRole('button', { name: '缩放至 125%' }).click();
  await expect(body).toBeFocused();

  await page.getByRole('tab', { name: '审阅' }).click();
  await page.getByRole('button', { name: '拼写检查' }).click();
  await expect(body).toBeFocused();
  await page.getByRole('button', { name: '修订模式' }).click();
  await expect(body).toBeFocused();

  const status = page.locator('.work-document-footer');
  await status.getByRole('button', { name: '网页视图' }).click();
  await expect(body).toBeFocused();
  await status.getByRole('button', { name: '页面视图' }).click();
  await expect(body).toBeFocused();
});

test('Word repeats Find and Replace commands by returning focus to the query', async ({
  page,
}) => {
  await page.goto('/');
  await openDocumentFixture(page);
  await waitForDocumentFixture(page);

  const body = page.getByRole('textbox', { name: '文档正文' });
  const modifier = process.platform === 'darwin' ? 'Meta' : 'Control';
  const modalPane =
    (page.viewportSize()?.width ?? Number.POSITIVE_INFINITY) <= 900;
  await body.focus();
  await page.keyboard.press(`${modifier}+f`);
  const query = page.getByRole('textbox', { name: '查找内容' });
  await expect(query).toBeFocused();
  await query.fill('项目');
  await expect
    .poll(() =>
      page
        .locator('.work-document-find-match')
        .evaluateAll(
          (elements) =>
            new Set(
              elements.map((element) =>
                element.getAttribute('data-document-find-index'),
              ),
            ).size,
        ),
    )
    .toBeGreaterThan(1);
  await expect(page.locator('.work-document-find-match.active')).toHaveCount(0);

  await query.press('Enter');
  await expect(query).toBeFocused();
  await expect(
    page.locator('.work-document-find-match.active').first(),
  ).toHaveAttribute('data-document-find-index', '0');
  await query.press('Enter');
  await expect(query).toBeFocused();
  await expect(
    page.locator('.work-document-find-match.active').first(),
  ).toHaveAttribute('data-document-find-index', '1');

  if (modalPane) await query.press(`${modifier}+f`);
  else {
    await body.focus();
    await page.keyboard.press(`${modifier}+f`);
  }
  await expect(query).toBeFocused();
  await expect(query).toHaveJSProperty('selectionStart', 0);
  await expect(query).toHaveJSProperty('selectionEnd', 2);

  if (modalPane) await query.press(`${modifier}+h`);
  else {
    await body.focus();
    await page.keyboard.press(`${modifier}+h`);
  }
  await expect(page.getByRole('tab', { name: '替换' })).toHaveAttribute(
    'aria-selected',
    'true',
  );
  await expect(query).toBeFocused();

  const replacement = page.getByRole('textbox', { name: '替换为' });
  await replacement.fill('方案');
  await replacement.press('Enter');
  await expect(replacement).toBeFocused();
  await expect(page.getByText('已替换当前匹配')).toBeVisible();
});

test('Word list galleries apply a style without trapping focus in the ribbon', async ({
  page,
}) => {
  await page.goto('/');
  await openDocumentFixture(page);
  await waitForDocumentFixture(page);

  const body = page.getByRole('textbox', { name: '文档正文' });
  await body.focus();

  await page.getByRole('button', { name: '项目符号库' }).click();
  let library = page.getByRole('dialog', { name: '项目符号库' });
  const disc = library.getByRole('menuitemradio', { name: '实心圆点' });
  const circle = library.getByRole('menuitemradio', { name: '空心圆点' });
  await expect(disc).toBeFocused();
  await page.keyboard.press('ArrowRight');
  await expect(circle).toBeFocused();
  await expect(circle).toHaveAttribute('tabindex', '0');
  await expect(disc).toHaveAttribute('tabindex', '-1');
  await page.keyboard.press('Tab');
  await expect(library).toBeHidden();
  await expect(
    page.getByRole('button', { name: '编号', exact: true }),
  ).toBeFocused();

  await page.getByRole('button', { name: '项目符号库' }).click();
  library = page.getByRole('dialog', { name: '项目符号库' });
  await library.getByRole('menuitemradio', { name: '空心圆点' }).click();
  await expect(body).toBeFocused();

  await page.getByRole('button', { name: '项目符号库' }).click();
  library = page.getByRole('dialog', { name: '项目符号库' });
  await library.getByRole('button', { name: '清除项目符号' }).click();
  await expect(body).toBeFocused();

  await page.getByRole('button', { name: '编号库' }).click();
  const numbering = page.getByRole('dialog', { name: '编号库' });
  await numbering.getByRole('menuitemradio', { name: '小写字母' }).click();
  await expect(body).toBeFocused();
});

test('Word page and reference commands keep the document ready for typing', async ({
  page,
}) => {
  await page.goto('/');
  await openDocumentFixture(page);
  await waitForDocumentFixture(page);

  const body = page.getByRole('textbox', { name: '文档正文' });
  await body.focus();

  await page.getByRole('tab', { name: '页面布局' }).click();
  await page.getByRole('button', { name: '显示页码' }).click();
  await expect(body).toBeFocused();
  await page.getByRole('button', { name: '插入分节符' }).click();
  await expect(body).toBeFocused();

  await page.getByRole('button', { name: '页面颜色' }).click();
  const pageColor = page.getByRole('dialog', { name: '页面颜色' });
  await pageColor.getByRole('option', { name: '颜色 #d9ead3' }).click();
  await expect(body).toBeFocused();

  await page.getByRole('tab', { name: '插入' }).click();
  await page.getByRole('combobox', { name: '插入页码或日期' }).click();
  await expect(page.getByRole('option', { name: '页码或日期' })).toBeDisabled();
  await expect(
    page.getByRole('option', { name: '页码', exact: true }),
  ).toBeFocused();
  await page.getByRole('option', { name: '当前日期' }).click();
  await expect(body).toBeFocused();

  await page.getByRole('tab', { name: '引用' }).click();
  await page.getByRole('button', { name: '插入脚注' }).click();
  await expect(body).toBeFocused();
  await page.getByRole('button', { name: '更新页码和日期' }).click();
  await expect(body).toBeFocused();
});

test('Word keeps cross-references truthful when their caption is deleted', async ({
  page,
}) => {
  await page.goto('/');
  await openDocumentFixture(page);
  await waitForDocumentFixture(page);

  const body = page.getByRole('textbox', { name: '文档正文' });
  await body.focus();
  await body.press('Control+End');
  await page.getByRole('tab', { name: '引用' }).click();
  await page.getByRole('button', { name: '插入图片题注' }).click();
  const captionDialog = page.getByRole('dialog', { name: '插入图片题注' });
  await captionDialog
    .getByRole('textbox', { name: '题注文字' })
    .fill('Architecture');
  await captionDialog.getByRole('button', { name: '插入题注' }).click();

  const caption = page.locator('.work-document-caption');
  await expect(caption).toHaveAttribute('aria-label', '图 1 Architecture');
  await page.getByRole('button', { name: '插入交叉引用' }).click();
  const referenceDialog = page.getByRole('dialog', { name: '插入交叉引用' });
  await expect(
    referenceDialog.getByRole('radio', { name: '图 1 Architecture' }),
  ).toBeChecked();
  await referenceDialog.getByRole('button', { name: '插入引用' }).click();

  const reference = page.locator('.work-document-cross-reference');
  await expect(reference).toHaveText('图 1');
  await expect(reference).not.toHaveAttribute(
    'data-reference-orphaned',
    'true',
  );

  await caption.selectText();
  await page.keyboard.press('Backspace');
  await expect(caption).toHaveAttribute('aria-label', '图 1');
  await page.keyboard.press('Backspace');

  await expect(caption).toHaveCount(0);
  await expect(reference).toHaveAttribute('data-reference-orphaned', 'true');
  await expect(reference).toHaveText('引用缺失');
  await expect(reference).toHaveCSS('color', 'rgb(181, 59, 59)');
});

test('Word opens the selected-text menu from the keyboard at the selection', async ({
  page,
}) => {
  await page.goto('/');
  await page
    .getByRole('button', { name: '项目方案 目标、范围、里程碑与风险' })
    .click();
  const editor = page.getByRole('textbox', { name: '文档正文' });
  await editor.evaluate((element) => {
    const text = element.querySelector('h1')?.firstChild;
    if (!(text instanceof Text)) {
      throw new Error('Document context-menu selection is unavailable.');
    }
    const range = document.createRange();
    range.setStart(text, 0);
    range.setEnd(text, Math.min(4, text.length));
    const selection = window.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(range);
    (element as HTMLElement).focus();
    document.dispatchEvent(new Event('selectionchange', { bubbles: true }));
  });
  await editor.press('Shift+F10');

  const menu = page.getByRole('menu', { name: '选中文本操作' });
  await expect(menu).toBeVisible();
  await expect(menu.getByRole('menuitem').first()).toBeFocused();
  const geometry = await page.evaluate(() => {
    const selection = window.getSelection();
    const menu = document.querySelector<HTMLElement>('.workspace-context-menu');
    if (!selection?.rangeCount || !menu) {
      throw new Error('Document keyboard menu geometry is unavailable.');
    }
    const selectionRect = selection.getRangeAt(0).getBoundingClientRect();
    const menuRect = menu.getBoundingClientRect();
    return {
      expectedLeft: selectionRect.left + selectionRect.width / 2,
      expectedTop: selectionRect.top + selectionRect.height / 2,
      menuLeft: menuRect.left,
      menuTop: menuRect.top,
    };
  });
  // Moving focus into the menu can round the live selection rectangle by one
  // CSS pixel, depending on the browser and viewport width.
  expect(
    Math.abs(geometry.menuLeft - geometry.expectedLeft),
  ).toBeLessThanOrEqual(1);
  expect(Math.abs(geometry.menuTop - geometry.expectedTop)).toBeLessThanOrEqual(
    1,
  );

  await page.keyboard.press('Escape');
  await expect(menu).toBeHidden();
  await expect(editor).toBeFocused();
});

test('Word asks for a question before preparing selected-text context for AI', async ({
  page,
}) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await page.goto('/');
  await page
    .getByRole('button', { name: '项目方案 目标、范围、里程碑与风险' })
    .click();
  const editor = page.getByRole('textbox', { name: '文档正文' });
  await editor.evaluate((element) => {
    const text = element.querySelector('h1')?.firstChild;
    if (!(text instanceof Text)) {
      throw new Error('Document AI question selection is unavailable.');
    }
    const range = document.createRange();
    range.setStart(text, 0);
    range.setEnd(text, Math.min(4, text.length));
    const selection = window.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(range);
    (element as HTMLElement).focus();
    document.dispatchEvent(new Event('selectionchange', { bubbles: true }));
  });
  await editor.press('Shift+F10');
  await page.getByRole('menuitem', { name: '询问 AI 助手' }).click();

  const assistant = page.getByRole('dialog', { name: 'AI 助手' });
  const question = assistant.getByRole('textbox', {
    name: '向 AI 助手提问',
  });
  const send = assistant.getByRole('button', { name: '发送问题' });
  await expect(question).toBeFocused();
  await expect(question).toBeEmpty();
  await expect(send).toBeDisabled();
  await expect(
    assistant.getByText('已附带选中文本和文档上下文。'),
  ).toBeVisible();

  await question.fill('这段结论有哪些依据？');
  await expect(send).toBeEnabled();
  await send.click();

  await expect(
    assistant.getByRole('heading', {
      name: '请结合已附带的文档上下文回答： 这段结论有哪些依据？',
    }),
  ).toBeVisible();
  const context = assistant.getByText('查看附带上下文').locator('..');
  await expect(context).not.toHaveAttribute('open');
  await expect(context.locator('blockquote')).toBeHidden();
  await context.getByText('查看附带上下文').click();
  await expect(context.locator('blockquote')).toBeVisible();
});
