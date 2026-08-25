import { expect, test } from '@playwright/test';

test('Presentation entrance animations author and play ordered cues', async ({
  page,
}) => {
  await page.goto('/');
  await page.getByRole('button', { name: '打开最新能力：入场动画' }).click();

  const canvas = page.locator('.work-slide-canvas.interactive');
  await expect(canvas).toBeVisible();
  await expect(page.getByRole('textbox', { name: '文件名' })).toHaveValue(
    '入场动画示例',
  );

  const title = canvas
    .locator('[data-slide-element-id]')
    .filter({ hasText: '让信息按叙事顺序出现' });
  await title.click();
  await page.getByRole('tab', { name: '动画', exact: true }).click();

  const effect = page.getByRole('combobox', {
    name: '对象入场动画效果',
  });
  const trigger = page.getByRole('combobox', {
    name: '对象入场动画触发方式',
  });
  const duration = page.getByRole('textbox', {
    name: '对象入场动画持续秒数',
  });
  await expect(effect).toHaveText('出现');
  await expect(trigger).toHaveText('单击时');
  await expect(duration).toHaveValue('0.3');

  await trigger.click();
  await page.getByRole('option', { name: '上一动画之后' }).click();
  await duration.fill('0.8');
  await duration.press('Enter');
  await expect(trigger).toHaveText('上一动画之后');
  await expect(duration).toHaveValue('0.8');

  const preview = page.getByRole('button', {
    name: '预览当前幻灯片动画',
  });
  await preview.click();
  const dialog = page.getByRole('dialog', { name: '幻灯片放映' });
  await expect(dialog).toBeVisible();
  await expect(
    dialog.locator('[data-slide-animation-state="playing"]'),
  ).toHaveCount(3);
  await expect(
    dialog.locator('[data-slide-animation-state="pending"]'),
  ).toHaveCount(1);

  await dialog.getByRole('button', { name: '单击换到下一张幻灯片' }).click();
  await expect(
    dialog.locator('[data-slide-animation-state="finished"]'),
  ).toHaveCount(3);
  await expect(
    dialog.locator(
      '[data-slide-animation-effect="zoom"][data-slide-animation-state="playing"]',
    ),
  ).toHaveCount(1);

  await page.keyboard.press('Escape');
  await expect(dialog).toBeHidden();
  await expect(preview).toBeFocused();
});
