import { expect, test } from '@playwright/test';

test('Presentation formatting shortcuts update only the active text selection', async ({
  page,
}) => {
  await page.goto('/');
  await page
    .getByRole('button', { name: '业务策略汇报 PPTX · 本次会话' })
    .click();

  const canvas = page.locator('.work-slide-canvas.interactive');
  await canvas.waitFor();
  const title = canvas.locator(':scope > .work-slide-element').nth(1);
  await title.dblclick();
  const textbox = canvas.getByRole('textbox', { name: '幻灯片文本' });
  await expect(textbox).toBeFocused();

  const selectedText = await textbox.evaluate((element) => {
    const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT);
    const text = walker.nextNode();
    if (!(text instanceof Text) || !text.data.length) {
      throw new Error('Presentation title text is unavailable.');
    }
    const end = Math.min(4, text.data.length);
    const range = document.createRange();
    range.setStart(text, 0);
    range.setEnd(text, end);
    const selection = window.getSelection();
    selection?.removeAllRanges();
    selection?.addRange(range);
    document.dispatchEvent(new Event('selectionchange', { bubbles: true }));
    return text.data.slice(0, end);
  });

  await page.keyboard.press('Control+i');
  const italic = textbox.locator('em');
  await expect(italic).toHaveText(selectedText);
  const fullText = (await textbox.textContent()) ?? '';
  expect(selectedText.length).toBeLessThan(fullText.length);

  await page.keyboard.press('Control+u');
  await expect(textbox.locator('u')).toHaveText(selectedText);
  await expect(textbox).toBeFocused();

  await page.getByRole('button', { name: '加粗' }).click();
  await expect(textbox.locator('strong')).toHaveText(
    fullText.slice(selectedText.length),
  );
  await expect(textbox).toHaveText(fullText);
  await expect(textbox).toBeFocused();
});

test('Presentation clipboard shortcuts keep object selection and focus coherent', async ({
  page,
}) => {
  await page.goto('/');
  await page
    .getByRole('button', { name: '业务策略汇报 PPTX · 本次会话' })
    .click();

  const canvas = page.locator('.work-slide-canvas.interactive');
  await canvas.waitFor();
  const objects = canvas.locator(':scope > .work-slide-element');
  await expect(objects).toHaveCount(3);
  const title = objects.nth(1);
  await title.click();
  await expect(title).toBeFocused();
  const titleBounds = await title.boundingBox();
  if (!titleBounds) throw new Error('Presentation title geometry is missing.');

  await expect(
    page.getByRole('button', { name: '复制', exact: true }),
  ).toHaveAttribute('aria-keyshortcuts', 'Control+C Meta+C');
  await page.keyboard.press('Meta+c');
  await page.keyboard.press('Meta+v');
  await expect(objects).toHaveCount(4);
  const pasted = canvas.locator(
    ':scope > .work-slide-element[data-slide-element-selected="true"]',
  );
  await expect(pasted).toHaveCount(1);
  await expect(pasted).toBeFocused();
  const pastedBounds = await pasted.boundingBox();
  if (!pastedBounds) {
    throw new Error('Pasted presentation object geometry is missing.');
  }
  expect(pastedBounds.x - titleBounds.x).toBeGreaterThanOrEqual(20);
  expect(pastedBounds.y - titleBounds.y).toBeGreaterThanOrEqual(10);

  await page.keyboard.press('Control+x');
  await expect(objects).toHaveCount(3);
  await expect(
    page.locator('.work-slide-strip [data-slide-thumbnail].active'),
  ).toBeFocused();
  await page.keyboard.press('Control+v');
  await expect(objects).toHaveCount(4);
  await expect(pasted).toHaveCount(1);
  await expect(pasted).toBeFocused();
});

test('Presentation history shortcuts stay scoped to the editor', async ({
  page,
}) => {
  await page.goto('/');
  await page
    .getByRole('button', { name: '业务策略汇报 PPTX · 本次会话' })
    .click();

  const thumbnails = page.locator('.work-slide-strip [data-slide-thumbnail]');
  await expect(thumbnails).toHaveCount(3);
  await page.getByRole('button', { name: '新建幻灯片' }).click();
  await expect(thumbnails).toHaveCount(4);

  const hostAction = page.getByRole('button', {
    name: '打开 AI 助手',
    exact: true,
  });
  await hostAction.focus();
  await hostAction.press('Meta+z');
  await expect(thumbnails).toHaveCount(4);
  await expect(hostAction).toBeFocused();

  const activeSlide = page.locator(
    '.work-slide-strip [data-slide-thumbnail].active',
  );
  await expect(activeSlide).toHaveCount(1);
  await activeSlide.focus();
  await activeSlide.press('Meta+z');
  await expect(thumbnails).toHaveCount(3);
});

test('Presentation numeric menus commit complete font and timing values', async ({
  page,
}) => {
  await page.goto('/');
  await page
    .getByRole('button', { name: '业务策略汇报 PPTX · 本次会话' })
    .click();

  const canvas = page.locator('.work-slide-canvas.interactive');
  await canvas.waitFor();
  const title = canvas.locator(':scope > .work-slide-element').nth(1);
  await title.dblclick();

  const fontSize = page.getByRole('textbox', { name: '演示字号' });
  const initialFontSize = await fontSize.inputValue();
  await fontSize.fill('');
  await expect(fontSize).toHaveValue('');
  await fontSize.fill('120');
  await expect(fontSize).toHaveAttribute('aria-invalid', 'true');
  await expect(fontSize).toHaveValue('120');
  await page.keyboard.press('Enter');
  await expect(fontSize).toHaveValue('96');
  expect(initialFontSize).not.toBe('96');
  await expect(fontSize).toBeFocused();

  await page.getByRole('tab', { name: '切换' }).click();
  const effect = page.getByRole('combobox', { name: '幻灯片切换效果' });
  if ((await effect.textContent())?.trim() === '无') {
    await effect.click();
    await page.getByRole('option', { name: '淡化' }).click();
  }
  const automatic = page.getByRole('checkbox', { name: '自动换片' });
  if (!(await automatic.isChecked())) await automatic.click();

  const seconds = page.getByRole('textbox', { name: '自动换片秒数' });
  await expect(seconds).toBeEnabled();
  await seconds.fill('');
  await expect(seconds).toHaveValue('');
  await seconds.fill('7.75');
  await expect(seconds).toHaveValue('7.75');
  await page.keyboard.press('Enter');
  await expect(seconds).toHaveValue('7.75');
  await expect(seconds).toBeFocused();

  await page.getByRole('tab', { name: '开始' }).click();
  await page.getByRole('tab', { name: '切换' }).click();
  await expect(page.getByRole('textbox', { name: '自动换片秒数' })).toHaveValue(
    '7.75',
  );
});

test('Presentation font controls keep standard single-border geometry', async ({
  page,
}) => {
  await page.goto('/');
  await page
    .getByRole('button', { name: '业务策略汇报 PPTX · 本次会话' })
    .click();

  const canvas = page.locator('.work-slide-canvas.interactive');
  await canvas.waitFor();
  await canvas.locator(':scope > .work-slide-element').nth(1).click();

  const fontGroup = page.getByRole('region', { name: '字体' });
  const fontFamily = page.getByRole('combobox', { name: '演示字体' });
  const fontSize = page.getByRole('textbox', { name: '演示字号' });
  const fontSizeControl = fontSize.locator('..');
  const bold = page.getByRole('button', { name: '加粗' });

  await expect(fontGroup.getByText('字号', { exact: true })).toHaveCount(0);
  await expect(fontSizeControl).toHaveClass(/work-office-number-field/);
  await expect(fontSizeControl).toHaveCSS('border-top-width', '1px');
  await expect(fontSizeControl).toHaveCSS('height', '29px');

  const [familyBox, sizeBox, boldBox] = await Promise.all([
    fontFamily.boundingBox(),
    fontSizeControl.boundingBox(),
    bold.boundingBox(),
  ]);
  if (!familyBox || !sizeBox || !boldBox) {
    throw new Error('Presentation font control geometry is unavailable.');
  }
  expect(sizeBox.width).toBeGreaterThanOrEqual(72);
  expect(sizeBox.width).toBeLessThanOrEqual(88);
  expect(sizeBox.x).toBeGreaterThanOrEqual(familyBox.x + familyBox.width);
  expect(boldBox.x).toBeGreaterThanOrEqual(sizeBox.x + sizeBox.width);
});

test('Presentation returns object keyboard control after ribbon formatting', async ({
  page,
}) => {
  await page.goto('/');
  await page
    .getByRole('button', { name: '业务策略汇报 PPTX · 本次会话' })
    .click();

  const canvas = page.locator('.work-slide-canvas.interactive');
  await canvas.waitFor();
  const title = canvas.locator(':scope > .work-slide-element').nth(1);
  await title.click();
  await expect(title).toHaveAttribute('data-slide-element-selected', 'true');
  await expect(title).toBeFocused();

  await page.getByRole('button', { name: '加粗' }).click();
  await expect(title).toBeFocused();
  const leftBefore = await title.evaluate((element) => element.style.left);
  await page.keyboard.press('ArrowRight');
  await expect
    .poll(() => title.evaluate((element) => element.style.left))
    .not.toBe(leftBefore);

  await page.getByRole('combobox', { name: '演示字体' }).click();
  await page.getByRole('option', { name: 'Arial' }).click();
  await expect(title).toBeFocused();
  const topBefore = await title.evaluate((element) => element.style.top);
  await page.keyboard.press('ArrowDown');
  await expect
    .poll(() => title.evaluate((element) => element.style.top))
    .not.toBe(topBefore);
});

test('Presentation returns keyboard control to newly inserted objects', async ({
  page,
}) => {
  await page.goto('/');
  await page
    .getByRole('button', { name: '业务策略汇报 PPTX · 本次会话' })
    .click();

  const canvas = page.locator('.work-slide-canvas.interactive');
  await canvas.waitFor();
  await page.getByRole('tab', { name: '插入' }).click();

  await page.getByRole('button', { name: '形状' }).click();
  const selectedShape = canvas.locator(
    ':scope > .work-slide-element[data-slide-element-selected="true"]',
  );
  await expect(selectedShape).toHaveCount(1);
  await expect(selectedShape).toBeFocused();
  const shapeLeftBefore = await selectedShape.evaluate(
    (element) => element.style.left,
  );
  await page.keyboard.press('ArrowRight');
  await expect
    .poll(() => selectedShape.evaluate((element) => element.style.left))
    .not.toBe(shapeLeftBefore);

  await page.getByRole('button', { name: '图表' }).click();
  const selectedChart = canvas.locator(
    ':scope > .work-slide-element[data-slide-element-selected="true"]',
  );
  await expect(selectedChart).toHaveCount(1);
  await expect(selectedChart).toBeFocused();
  const chartTopBefore = await selectedChart.evaluate(
    (element) => element.style.top,
  );
  await page.keyboard.press('ArrowDown');
  await expect
    .poll(() => selectedChart.evaluate((element) => element.style.top))
    .not.toBe(chartTopBefore);
});

test('Presentation returns keyboard control after slide mutations and blocks them in layout editing', async ({
  page,
}) => {
  await page.goto('/');
  await page
    .getByRole('button', { name: '业务策略汇报 PPTX · 本次会话' })
    .click();
  await page.locator('.work-slide-canvas.interactive').waitFor();

  const strip = page.locator('.work-slide-strip');
  const thumbnails = strip.locator('[data-slide-thumbnail]');
  const activeThumbnail = strip.locator('[data-slide-thumbnail].active');
  const newSlide = page.getByRole('button', { name: '新建幻灯片' });
  await newSlide.click();
  await expect(thumbnails).toHaveCount(4);
  await expect(activeThumbnail).toBeFocused();

  await page.getByRole('button', { name: '复制幻灯片' }).click();
  await expect(thumbnails).toHaveCount(5);
  await expect(activeThumbnail).toBeFocused();

  await page.getByRole('button', { name: '删除幻灯片' }).click();
  await expect(thumbnails).toHaveCount(4);
  await expect(activeThumbnail).toBeFocused();

  await page.getByRole('tab', { name: '设计', exact: true }).click();
  await page.getByRole('button', { name: '母版和版式' }).click();
  await page.getByRole('button', { name: '编辑当前布局' }).click();
  const homeTab = page.getByRole('tab', { name: '开始', exact: true });
  await homeTab.click();
  await expect(newSlide).toBeDisabled();
  await homeTab.press('Control+m');
  await expect(thumbnails).toHaveCount(4);
});

test('Presentation returns keyboard control after cutting a selected object', async ({
  page,
}) => {
  await page.goto('/');
  await page
    .getByRole('button', { name: '业务策略汇报 PPTX · 本次会话' })
    .click();
  const canvas = page.locator('.work-slide-canvas.interactive');
  await canvas.waitFor();
  const objects = canvas.locator(':scope > .work-slide-element');
  await expect(objects).toHaveCount(3);
  const title = objects.nth(1);
  await title.click();

  await page.getByRole('button', { name: '剪切' }).click();
  await expect(objects).toHaveCount(2);
  await expect(
    page.locator('.work-slide-strip [data-slide-thumbnail].active'),
  ).toBeFocused();
});

test('Presentation view switches move focus into the active slide view', async ({
  page,
}) => {
  await page.goto('/');
  await page
    .getByRole('button', { name: '业务策略汇报 PPTX · 本次会话' })
    .click();

  const canvas = page.locator('.work-slide-canvas.interactive');
  await canvas.waitFor();
  const elements = canvas.locator(':scope > .work-slide-element');
  await expect(elements).toHaveCount(3);
  const title = elements.nth(1);
  await title.click();
  await expect(title).toBeFocused();

  await page.getByRole('tab', { name: '视图' }).click();
  await page.getByRole('button', { name: '幻灯片浏览', exact: true }).click();
  const sorterSelection = page.locator(
    '.work-presentation-sorter [data-slide-thumbnail].active',
  );
  await expect(sorterSelection).toHaveCount(1);
  await expect(sorterSelection).toBeFocused();

  await page.getByRole('button', { name: '普通视图', exact: true }).click();
  const stripSelection = page.locator(
    '.work-slide-strip [data-slide-thumbnail].active',
  );
  await expect(stripSelection).toHaveCount(1);
  await expect(stripSelection).toBeFocused();

  const status = page.locator('.work-presentation-status');
  await status.getByRole('button', { name: '幻灯片浏览视图' }).click();
  await expect(sorterSelection).toBeFocused();
});

test('Presentation opens slide menus from the keyboard and preserves tab order', async ({
  page,
}) => {
  await page.goto('/');
  await page
    .getByRole('button', { name: '业务策略汇报 PPTX · 本次会话' })
    .click();
  await page.locator('.work-slide-canvas.interactive').waitFor();

  const firstSlide = page.getByRole('button', {
    name: '幻灯片 1 / 3：封面',
  });
  const secondSlide = page.getByRole('button', {
    name: '幻灯片 2 / 3：核心判断',
  });
  await firstSlide.focus();
  await firstSlide.press('Shift+F10');

  const menu = page.getByRole('menu', { name: '幻灯片操作' });
  await expect(menu).toBeVisible();
  await expect(menu.getByRole('menuitem').first()).toBeFocused();
  const geometry = await page.evaluate(() => {
    const slide = document.querySelector<HTMLElement>(
      '[data-slide-thumbnail][data-slide-index="0"]',
    );
    const menu = document.querySelector<HTMLElement>('.workspace-context-menu');
    if (!(slide && menu))
      throw new Error('Slide menu geometry is unavailable.');
    const slideRect = slide.getBoundingClientRect();
    const menuRect = menu.getBoundingClientRect();
    return {
      expectedLeft: slideRect.left + slideRect.width / 2,
      expectedTop: slideRect.top + slideRect.height / 2,
      menuLeft: menuRect.left,
      menuTop: menuRect.top,
    };
  });
  expect(geometry.menuLeft).toBeCloseTo(geometry.expectedLeft, 0);
  expect(geometry.menuTop).toBeCloseTo(geometry.expectedTop, 0);

  await page.keyboard.press('Escape');
  await expect(menu).toBeHidden();
  await expect(firstSlide).toBeFocused();

  const titleObject = page.getByRole('group', {
    name: '业务策略汇报',
    exact: true,
  });
  await titleObject.focus();
  await titleObject.press('Shift+F10');
  const objectMenu = page.getByRole('menu', { name: '演示对象操作' });
  await expect(objectMenu).toBeVisible();
  await expect(objectMenu.getByRole('menuitem').first()).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(objectMenu).toBeHidden();
  await expect(titleObject).toBeFocused();

  await firstSlide.focus();
  await firstSlide.press('Shift+F10');
  await expect(menu).toBeVisible();
  await page.keyboard.press('Tab');
  await expect(menu).toBeHidden();
  await page.evaluate(
    () =>
      new Promise<void>((resolve) =>
        requestAnimationFrame(() => requestAnimationFrame(() => resolve())),
      ),
  );
  await expect(secondSlide).toBeFocused();
});
