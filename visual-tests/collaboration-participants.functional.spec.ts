import { expect, test } from '@playwright/test';
import {
  openDocumentFixture,
  waitForDocumentFixture,
} from './visual-test-support';

test('projects host-owned participants across edit and preview chrome', async ({
  page,
}) => {
  const browserErrors: string[] = [];
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await page.clock.install();

  await page.goto('/playground/?e2e=collaboration-presence');
  await openDocumentFixture(page);
  await waitForDocumentFixture(page);

  const trigger = page.getByRole('button', {
    exact: true,
    name: '查看协作者，2 位协作者',
  });
  await expect(trigger).toBeVisible();
  await expect(trigger).toHaveAttribute('data-collaboration-count', '2');
  await expect(
    page.locator(
      '.work-office-remote-caret[data-participant-id="playground-agent"]',
    ),
  ).toBeVisible();
  await expect(
    page.locator(
      '.work-office-remote-selection[data-participant-id="playground-agent"]',
    ),
  ).not.toHaveCount(0);

  await page.clock.fastForward(31_000);
  await expect(trigger).toHaveAttribute('data-collaboration-count', '2');
  await expect(
    page.locator(
      '.work-office-remote-caret[data-participant-id="playground-agent"]',
    ),
  ).toBeVisible();

  await trigger.click();
  const roster = page.getByRole('dialog', { name: '协作者' });
  await expect(roster).toBeVisible();
  await expect(roster).toContainText('林澄');
  await expect(roster).toContainText('A3S Agent');
  await expect(roster).toContainText('正在使用 · 建议 · 已选择 6 个位置');
  const bounds = await roster.boundingBox();
  expect(bounds).not.toBeNull();
  expect((bounds?.x ?? 0) + (bounds?.width ?? 0)).toBeLessThanOrEqual(
    page.viewportSize()?.width ?? 0,
  );

  await page.keyboard.press('Escape');
  await expect(roster).toBeHidden();
  await expect(trigger).toBeFocused();

  await trigger.click();
  const locateAgent = page.getByRole('button', {
    name: '跳转到 A3S Agent 的位置，已选择 6 个位置',
  });
  await expect(locateAgent).toBeFocused();
  await locateAgent.click();
  const documentBody = page.getByRole('textbox', { name: '文档正文' });
  await expect(roster).toBeHidden();
  await expect(documentBody).toBeFocused();
  expect(
    await page.evaluate(() => window.getSelection()?.toString().length ?? 0),
  ).toBeGreaterThan(0);

  await page.getByRole('button', { exact: true, name: '预览' }).click();
  const previewBar = page.getByRole('region', { name: '文字预览工具' });
  const previewTrigger = previewBar.getByRole('button', {
    exact: true,
    name: '查看协作者，2 位协作者',
  });
  await expect(previewTrigger).toBeVisible();
  await previewTrigger.click();
  await expect(page.getByRole('dialog', { name: '协作者' })).toContainText(
    '2 个在线会话',
  );
  await page.keyboard.press('Escape');
  await expect(previewTrigger).toBeFocused();

  expect(browserErrors).toEqual([]);
});
