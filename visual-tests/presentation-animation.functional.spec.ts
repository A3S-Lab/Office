import { expect, test } from '@playwright/test';

test('Presentation entrance and exit animations author and play ordered cues', async ({
  page,
}) => {
  await page.goto('/playground/');
  await page.locator("button[data-template-id='animated-deck']").click();

  const canvas = page.locator('.work-slide-canvas.interactive');
  await expect(canvas).toBeVisible();
  await expect(page.getByRole('textbox', { name: '文件名' })).toHaveValue(
    '进入与退出动画示例',
  );

  const flyingObject = canvas
    .locator('[data-slide-element-id]')
    .filter({ hasText: '从左侧飞入，再向右侧飞出' });
  await flyingObject.click();
  await page.getByRole('tab', { name: '动画', exact: true }).click();

  const animationClass = page.getByRole('combobox', {
    name: '对象动画类型',
  });
  const effect = page.getByRole('combobox', {
    name: '对象动画效果',
  });
  const trigger = page.getByRole('combobox', {
    name: '对象动画触发方式',
  });
  const duration = page.getByRole('textbox', {
    name: '对象动画持续秒数',
  });
  const direction = page.getByRole('combobox', {
    name: '对象动画方向',
  });
  await expect(animationClass).toHaveText('进入');
  await expect(effect).toHaveText('飞入');
  await expect(trigger).toHaveText('上一动画之后');
  await expect(direction).toHaveText('从左侧');

  await animationClass.click();
  await page.getByRole('option', { name: '退出' }).click();
  await expect(effect).toHaveText('飞出');
  await expect(trigger).toHaveText('上一动画之后');
  await expect(direction).toHaveText('向右侧');
  await direction.click();
  await page.getByRole('option', { name: '向下方' }).click();
  await duration.fill('0.8');
  await duration.press('Enter');
  await expect(direction).toHaveText('向下方');
  await expect(duration).toHaveValue('0.8');

  const preview = page.getByRole('button', {
    name: '预览当前幻灯片动画',
  });
  await preview.click();
  const dialog = page.getByRole('dialog', { name: '幻灯片放映' });
  await expect(dialog).toBeVisible();
  await expect(
    dialog.locator('[data-slide-animation-state="hidden"]'),
  ).toHaveCount(4);

  await dialog.getByRole('button', { name: '单击换到下一张幻灯片' }).click();
  await expect(
    dialog.locator(
      '[data-slide-animation-state="playing"][data-slide-animation-final-visibility="visible"]',
    ),
  ).toHaveCount(4);

  await dialog.getByRole('button', { name: '单击换到下一张幻灯片' }).click();
  await expect(
    dialog.locator(
      '[data-slide-animation-effect="fly-out"][data-slide-animation-state="playing"][data-slide-animation-final-visibility="hidden"]',
    ),
  ).toHaveCount(1);
  await expect(
    dialog.locator(
      '[data-slide-animation-state="playing"][data-slide-animation-final-visibility="hidden"]',
    ),
  ).toHaveCount(4);
  await expect(
    dialog.getByRole('button', { name: '下一张', exact: true }),
  ).toBeDisabled();

  await page.keyboard.press('Escape');
  await expect(dialog).toBeHidden();
  await expect(preview).toBeFocused();
});
