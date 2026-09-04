import { expect, type Page, test } from '@playwright/test';
import { waitForDocumentFixture } from './visual-test-support';

const movedText = '可复用的段落';

test('Writer accepts a paired move from one destination-focused review card', async ({
  page,
}, testInfo) => {
  test.skip(
    testInfo.project.name !== 'desktop-1280',
    'Destination selection is intentionally covered on the non-modal desktop pane.',
  );
  const pane = await openMoveRevisionDemo(page);
  const card = pane.locator('.work-document-change-item.move');

  await expect(card).toBeVisible();
  await expect(card.getByText('移动', { exact: true })).toBeVisible();
  await expect(card.getByText(movedText, { exact: true })).toBeVisible();
  await expect(card.getByText(/Nora Reviewer/)).toBeVisible();
  await card.getByRole('button', { name: '定位修订 1' }).click();
  await expectDestinationSelection(page);

  await page.screenshot({
    path: testInfo.outputPath('writer-move-revision-review.png'),
    fullPage: false,
  });

  await card.getByRole('button', { name: '接受修订 1' }).click();
  const editor = page.locator('.work-document-editable .ProseMirror');
  await expect(editor.locator('[data-change-kind="move"]')).toHaveCount(0);
  await expect(
    editor.locator('[data-document-change="true"][data-change-kind="move"]'),
  ).toHaveCount(0);
  await expect(
    editor.locator('p').filter({ hasText: '目标位置：' }),
  ).toContainText(movedText);
  await expect(
    editor.locator('p').filter({ hasText: '留在原位置' }),
  ).not.toContainText(movedText);
  await expect(pane.getByText('正在记录新的改动。')).toBeVisible();
});

test('Writer rejects a paired move atomically in the phone review dialog', async ({
  page,
}, testInfo) => {
  await page.setViewportSize({ width: 390, height: 844 });
  const pane = await openMoveRevisionDemo(page);

  await expect(pane).toHaveAttribute('role', 'dialog');
  await expect(pane).toHaveAttribute('aria-modal', 'true');
  const bounds = await pane.boundingBox();
  expect(bounds).not.toBeNull();
  expect((bounds?.x ?? 0) + (bounds?.width ?? 0)).toBeLessThanOrEqual(391);
  expect((bounds?.y ?? 0) + (bounds?.height ?? 0)).toBeLessThanOrEqual(845);

  const card = pane.locator('.work-document-change-item.move');
  await expect(card.getByRole('button', { name: '接受修订 1' })).toBeVisible();
  await expect(card.getByRole('button', { name: '拒绝修订 1' })).toBeVisible();
  await page.screenshot({
    path: testInfo.outputPath('writer-move-revision-phone.png'),
    fullPage: false,
  });

  await card.getByRole('button', { name: '拒绝修订 1' }).click();
  const editor = page.locator('.work-document-editable .ProseMirror');
  await expect(editor.locator('[data-change-kind="move"]')).toHaveCount(0);
  await expect(
    editor.locator('p').filter({ hasText: '留在原位置' }),
  ).toContainText(movedText);
  await expect(
    editor.locator('p').filter({ hasText: '目标位置：' }),
  ).not.toContainText(movedText);
  await expect(pane.getByText('正在记录新的改动。')).toBeVisible();
});

async function openMoveRevisionDemo(page: Page) {
  await page.goto('/playground/');
  await page.getByRole('button', { name: '体验移动修订' }).click();
  await waitForDocumentFixture(page);
  await expect(
    page.locator(
      '[data-document-change="true"][data-change-kind="move"][data-change-move-role="from"]',
    ),
  ).toBeVisible();
  await expect(
    page.locator(
      '[data-document-change="true"][data-change-kind="move"][data-change-move-role="to"]',
    ),
  ).toBeVisible();
  await page.getByRole('tab', { name: '审阅' }).click();
  const trigger = page.getByRole('button', { name: '查看修订（1）' });
  await expect(trigger).toBeVisible();
  await trigger.click();
  const pane = page.locator('.work-document-changes-panel');
  await expect(pane).toBeVisible();
  await expect(pane.getByRole('list', { name: '待处理修订' })).toHaveAttribute(
    'data-document-change-count',
    '1',
  );
  return pane;
}

async function expectDestinationSelection(page: Page): Promise<void> {
  await expect
    .poll(() =>
      page.evaluate(() => {
        const selection = window.getSelection();
        const anchor = selection?.anchorNode?.parentElement;
        const focus = selection?.focusNode?.parentElement;
        const roles = [anchor, focus]
          .map((node) =>
            node
              ?.closest('[data-change-move-role]')
              ?.getAttribute('data-change-move-role'),
          )
          .filter((value): value is string => Boolean(value));
        return {
          role: roles.includes('to') ? 'to' : undefined,
          text: selection?.toString() ?? '',
        };
      }),
    )
    .toEqual({ role: 'to', text: movedText });
}
