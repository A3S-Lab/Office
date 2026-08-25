import { expect, test } from '@playwright/test';

test('Presenter view keeps one navigation strip and aligns its content', async ({
  page,
}) => {
  await page.goto('/playground/');
  await page
    .getByRole('button', { name: '业务策略汇报 PPTX · 本次会话' })
    .click();
  await page.getByRole('tab', { name: '幻灯片放映' }).click();
  await page.getByRole('button', { name: '从头开始放映' }).click();

  const dialog = page.getByRole('dialog', { name: '幻灯片放映' });
  await expect(dialog).toBeVisible();
  await dialog.getByRole('button', { name: '演讲者视图' }).click();

  const presenter = dialog.getByRole('region', { name: '演讲者视图' });
  await expect(presenter).toBeVisible();
  await expect(
    dialog.getByRole('button', { name: '演讲者上一张' }),
  ).toHaveCount(0);
  await expect(
    dialog.getByRole('button', { name: '演讲者下一张' }),
  ).toHaveCount(0);
  await expect(
    dialog.getByRole('button', { name: '上一张', exact: true }),
  ).toHaveCount(1);
  await expect(
    dialog.getByRole('button', { name: '下一张', exact: true }),
  ).toHaveCount(1);

  const currentHeading = presenter
    .getByRole('region', { name: '当前幻灯片' })
    .getByRole('heading');
  const nextHeading = presenter
    .getByRole('region', { name: '下一张幻灯片' })
    .getByRole('heading');
  const [currentBounds, nextBounds] = await Promise.all([
    currentHeading.boundingBox(),
    nextHeading.boundingBox(),
  ]);
  if (!currentBounds || !nextBounds) {
    throw new Error('Presenter heading geometry is unavailable.');
  }
  expect(Math.abs(currentBounds.y - nextBounds.y)).toBeLessThanOrEqual(40);

  const presenterToggle = dialog.getByRole('button', {
    name: '退出演讲者视图',
  });
  await presenterToggle.focus();
  await presenterToggle.press('ArrowRight');
  await expect(
    presenter
      .getByRole('region', { name: '当前幻灯片' })
      .getByRole('heading', { name: '核心判断' }),
  ).toBeVisible();
  await presenterToggle.press('Escape');
  await expect(dialog).toBeHidden();
});

test('Presenter view keeps the current slide, cue, notes, and controls usable on phones', async ({
  page,
}, testInfo) => {
  test.skip(
    testInfo.project.name !== 'desktop-1280',
    'The phone contract only needs one browser project.',
  );
  await page.setViewportSize({ width: 390, height: 700 });
  await page.goto('/playground/');
  await page
    .getByRole('button', { name: '业务策略汇报 PPTX · 本次会话' })
    .click();
  await page.getByRole('tab', { name: '幻灯片放映' }).click();
  await page.getByRole('button', { name: '从头开始放映' }).click();

  const dialog = page.getByRole('dialog', { name: '幻灯片放映' });
  await dialog.getByRole('button', { name: '演讲者视图' }).click();
  const presenter = dialog.getByRole('region', { name: '演讲者视图' });
  const current = presenter.getByRole('region', { name: '当前幻灯片' });
  const next = presenter.getByRole('region', { name: '下一张幻灯片' });
  const notes = presenter.getByRole('complementary', { name: '演讲者备注' });
  const footer = dialog.locator('.work-presentation-player > footer');

  await expect(current.locator('.work-slide-canvas')).toBeVisible();
  await expect(next.locator('.work-slide-canvas')).toBeHidden();
  await expect(next.getByText('核心判断', { exact: true })).toBeVisible();
  await expect(notes).toBeVisible();
  await expect(footer).toBeVisible();
  await expect(
    dialog.getByRole('button', { name: '上一张', exact: true }),
  ).toHaveCount(1);
  await expect(
    dialog.getByRole('button', { name: '下一张', exact: true }),
  ).toHaveCount(1);

  const geometry = await dialog.evaluate((element) => {
    const currentRegion = element.querySelector<HTMLElement>(
      "[aria-label='当前幻灯片']",
    );
    const currentCanvas =
      currentRegion?.querySelector<HTMLElement>('.work-slide-canvas');
    const nextRegion = element.querySelector<HTMLElement>(
      "[aria-label='下一张幻灯片']",
    );
    const notesRegion = element.querySelector<HTMLElement>(
      "[aria-label='演讲者备注']",
    );
    const controls = element.querySelector<HTMLElement>(
      '.work-presentation-player > footer',
    );
    const player = element.querySelector<HTMLElement>(
      '.work-presentation-player',
    );
    if (
      !currentRegion ||
      !currentCanvas ||
      !nextRegion ||
      !notesRegion ||
      !controls ||
      !player
    ) {
      throw new Error('Phone presenter geometry is unavailable.');
    }
    const currentBounds = currentRegion.getBoundingClientRect();
    const canvasBounds = currentCanvas.getBoundingClientRect();
    const nextBounds = nextRegion.getBoundingClientRect();
    const notesBounds = notesRegion.getBoundingClientRect();
    const controlsBounds = controls.getBoundingClientRect();
    return {
      canvas: {
        bottom: canvasBounds.bottom,
        left: canvasBounds.left,
        right: canvasBounds.right,
        top: canvasBounds.top,
        width: canvasBounds.width,
      },
      controls: {
        bottom: controlsBounds.bottom,
        top: controlsBounds.top,
      },
      currentBottom: currentBounds.bottom,
      documentScrollWidth: document.documentElement.scrollWidth,
      nextBottom: nextBounds.bottom,
      nextTop: nextBounds.top,
      notesBottom: notesBounds.bottom,
      notesHeight: notesBounds.height,
      notesTop: notesBounds.top,
      playerClientWidth: player.clientWidth,
      playerScrollWidth: player.scrollWidth,
      viewportHeight: document.documentElement.clientHeight,
      viewportWidth: document.documentElement.clientWidth,
    };
  });

  expect(geometry.documentScrollWidth).toBeLessThanOrEqual(
    geometry.viewportWidth + 1,
  );
  expect(geometry.playerScrollWidth).toBeLessThanOrEqual(
    geometry.playerClientWidth + 1,
  );
  expect(geometry.canvas.left).toBeGreaterThanOrEqual(-1);
  expect(geometry.canvas.right).toBeLessThanOrEqual(geometry.viewportWidth + 1);
  expect(geometry.canvas.width).toBeGreaterThanOrEqual(300);
  expect(geometry.canvas.top).toBeGreaterThanOrEqual(0);
  expect(geometry.canvas.bottom).toBeLessThan(geometry.viewportHeight);
  expect(geometry.currentBottom).toBeLessThanOrEqual(geometry.nextTop + 1);
  expect(geometry.nextBottom).toBeLessThanOrEqual(geometry.notesTop + 1);
  expect(geometry.notesHeight).toBeGreaterThanOrEqual(80);
  expect(geometry.notesBottom).toBeLessThanOrEqual(geometry.controls.top + 1);
  expect(geometry.controls.bottom).toBeLessThanOrEqual(
    geometry.viewportHeight + 1,
  );
});
