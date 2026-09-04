import { expect, type Locator, test } from '@playwright/test';

const revisedDocument = [
  '<h1>发布审阅基线</h1>',
  '<p><strong>基线版本：</strong>1.1</p>',
  '<h2>比较步骤</h2>',
  '<p>打开“审阅”，选择“比较文档”，再导入 DOCX、HTML 或 TXT 修订版本并生成修订。</p>',
  '<h2>审阅范围</h2>',
  '<p style="text-align: center">当前基线包含架构说明、协作协议与发布说明。</p>',
  '<p>新增的发布摘要必须在交付前确认。</p>',
  '<h2>决策要求</h2>',
  '<p>发布前必须逐项接受或拒绝所有确定性修订。</p>',
].join('');

const movedRevisionDocument = [
  '<h1>发布审阅基线</h1>',
  '<p><strong>基线版本：</strong>1.0</p>',
  '<h2>比较步骤</h2>',
  '<p>打开“审阅”，选择“比较文档”，再导入 TXT、DOCX、HTML 或修订版本。</p>',
  '<h2>审阅范围</h2>',
  '<p>当前基线包含架构说明与发布说明。</p>',
  '<h2>决策要求</h2>',
  '<p>发布前必须逐项接受或拒绝所有生成的修订。</p>',
].join('');

test('Writer compares an imported version and opens deterministic review responsively', async ({
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
  await page.locator("button[data-template-id='document-comparison']").click();

  const editor = page.getByRole('textbox', { name: '文档正文' });
  await expect(editor).toHaveAttribute('data-pagination-state', 'ready');
  await expect(page.getByLabel('文件名')).toHaveValue('文档比较示例');
  await page.getByRole('tab', { name: '审阅' }).click();
  const compareButton = page.getByRole('button', { name: '比较文档' });
  await compareButton.click();

  const dialog = page.getByRole('dialog', { name: '比较与合并文档' });
  await expect(dialog).toBeVisible();
  await expect(
    dialog.locator('button.work-document-compare-file'),
  ).toBeFocused();
  await expectDialogInsideViewport(dialog);
  await dialog.getByLabel('选择修订版本文件').setInputFiles({
    name: 'release-review.html',
    mimeType: 'text/html',
    buffer: Buffer.from(revisedDocument),
  });
  await dialog.getByLabel('比较结果修订者名称').fill('A3S Reviewer');
  await page.screenshot({
    path: testInfo.outputPath(
      `writer-document-comparison-dialog-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  await dialog.getByRole('button', { name: '生成比较结果' }).click();

  await expect(dialog).toBeHidden();
  const changes = page.getByRole('list', { name: '待处理修订' });
  await expect(changes).toBeVisible();
  const originalCount = await changes.getAttribute(
    'data-document-change-count',
  );
  expect(Number(originalCount)).toBeGreaterThan(0);
  await expect(
    editor.locator('ins[data-document-change="true"]'),
  ).not.toHaveCount(0);
  await expect(
    editor.locator('del[data-document-change="true"]'),
  ).not.toHaveCount(0);
  await expect(
    editor.locator('[data-change-kind="paragraph-formatting"]'),
  ).not.toHaveCount(0);
  await expect(
    editor.locator('[data-document-block-change="true"]'),
  ).not.toHaveCount(0);
  if (testInfo.project.name === 'compact-768') {
    await expect(
      page.getByRole('button', { name: '关闭修订审阅' }),
    ).toBeFocused();
  } else {
    await expect(compareButton).toBeFocused();
  }

  const acceptFirst = page.getByRole('button', {
    name: '接受修订 1',
    exact: true,
  });
  await acceptFirst.click();
  await expect
    .poll(async () =>
      Number(await changes.getAttribute('data-document-change-count')),
    )
    .toBe(Number(originalCount) - 1);
  if (Number(originalCount) > 1) {
    await expect(
      page.getByRole('button', { name: '接受修订 1', exact: true }),
    ).toBeFocused();
  }
  expect(browserErrors).toEqual([]);
});

test('Writer surfaces inferred text moves as native review cards', async ({
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
  await page.locator("button[data-template-id='document-comparison']").click();
  const editor = page.getByRole('textbox', { name: '文档正文' });
  await expect(editor).toHaveAttribute('data-pagination-state', 'ready');
  await page.getByRole('tab', { name: '审阅' }).click();
  const compareButton = page.getByRole('button', { name: '比较文档' });
  await compareButton.click();

  const dialog = page.getByRole('dialog', { name: '比较与合并文档' });
  await dialog.getByLabel('选择修订版本文件').setInputFiles({
    name: 'move-review.html',
    mimeType: 'text/html',
    buffer: Buffer.from(movedRevisionDocument),
  });
  await dialog.getByLabel('比较结果修订者名称').fill('Move Reviewer');
  await dialog.getByRole('button', { name: '生成比较结果' }).click();

  await expect(dialog).toBeHidden();
  const changes = page.getByRole('list', { name: '待处理修订' });
  await expect(changes).toBeVisible();
  const moveCards = changes.locator('.work-document-change-item.move');
  await expect(moveCards).not.toHaveCount(0);
  await expect(
    moveCards.first().getByText('移动', { exact: true }),
  ).toBeVisible();
  await expect(moveCards.first().locator('strong')).toContainText(/DOCX|TXT/);
  await expect(moveCards.first()).toContainText('Move Reviewer');
  await expect(
    editor.locator(
      '[data-document-change="true"][data-change-kind="move"][data-change-move-role="from"]',
    ),
  ).not.toHaveCount(0);
  await expect(
    editor.locator(
      '[data-document-change="true"][data-change-kind="move"][data-change-move-role="to"]',
    ),
  ).not.toHaveCount(0);
  await page.screenshot({
    path: testInfo.outputPath(
      `writer-document-comparison-move-${testInfo.project.name}.png`,
    ),
    animations: 'disabled',
  });
  expect(browserErrors).toEqual([]);
});

async function expectDialogInsideViewport(dialog: Locator): Promise<void> {
  const geometry = await dialog.evaluate((element) => {
    const bounds = element.getBoundingClientRect();
    return {
      bottom: bounds.bottom,
      centerX: bounds.left + bounds.width / 2,
      left: bounds.left,
      right: bounds.right,
      top: bounds.top,
      viewportHeight: document.documentElement.clientHeight,
      viewportWidth: document.documentElement.clientWidth,
    };
  });
  expect(geometry.left).toBeGreaterThanOrEqual(8);
  expect(geometry.right).toBeLessThanOrEqual(geometry.viewportWidth - 8);
  expect(geometry.top).toBeGreaterThanOrEqual(8);
  expect(geometry.bottom).toBeLessThanOrEqual(geometry.viewportHeight - 8);
  expect(geometry.centerX).toBeCloseTo(geometry.viewportWidth / 2, 0);
}
