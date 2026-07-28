import { expect, test, type Page } from '@playwright/test';

test('Presentation replaces task panes and closes them from the toolbar with Escape', async ({
  page,
}) => {
  await openPresentationFixture(page);

  await page.getByRole('tab', { name: '设计', exact: true }).click();
  await page.getByRole('button', { name: '母版和版式' }).click();
  const designPane = page.getByRole('region', { name: '母版与布局' });
  await expect(designPane).toBeVisible();

  await page.getByRole('tab', { name: '审阅', exact: true }).click();
  await page.getByRole('button', { name: '查看批注' }).click();
  await expect(designPane).toBeHidden();
  const commentsPane = page.getByRole('region', { name: '演示批注审阅' });
  await expect(commentsPane).toBeVisible();

  await page.keyboard.press('Escape');
  await expect(commentsPane).toBeHidden();

  await page.getByRole('tab', { name: '设计', exact: true }).click();
  await page.getByRole('button', { name: '母版和版式' }).click();
  await expect(designPane).toBeVisible();
  await page.keyboard.press('Escape');
  await expect(designPane).toBeHidden();
});

test('Presentation keeps empty comments in the dialog until they are valid', async ({
  page,
}) => {
  await openPresentationFixture(page);

  await page.getByRole('tab', { name: '审阅', exact: true }).click();
  await page.getByRole('button', { name: '新建批注' }).click();
  const dialog = page.getByRole('dialog', { name: '批注内容' });
  const input = dialog.getByRole('textbox', { name: '批注内容' });
  const submit = dialog.getByRole('button', { name: '添加批注' });

  await expect(submit).toBeDisabled();
  await input.fill('   ');
  await expect(submit).toBeDisabled();
  await expect(dialog.getByRole('alert')).toHaveText('请输入批注内容。');

  await input.fill('需要补充数据来源');
  await expect(submit).toBeEnabled();
  await page.keyboard.press('Escape');
  await expect(dialog).toBeHidden();
});

test('Presentation closes the contextual chart pane with Escape', async ({
  page,
}) => {
  await openPresentationFixture(page);

  await page.getByRole('tab', { name: '插入', exact: true }).click();
  await page.getByRole('button', { name: '图表', exact: true }).click();
  const chartPane = page.getByRole('region', {
    name: '演示图表数据',
    exact: true,
  });
  await expect(chartPane).toBeVisible();

  await page.keyboard.press('Escape');
  await expect(chartPane).toBeHidden();
});

async function openPresentationFixture(page: Page) {
  await page.goto('/');
  await page
    .getByRole('button', {
      name: '业务策略汇报 PPTX · 本次会话',
    })
    .click();
  await page.locator('.work-slide-canvas.interactive').waitFor();
  await expect(page.locator('.work-presentation-editor')).toHaveAttribute(
    'data-presentation-geometry-state',
    'idle',
  );
}
