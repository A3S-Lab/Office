import { expect, type Locator, type Page, test } from '@playwright/test';

const gradientColors = [
  '#1d4ed8',
  '#b42318',
  '#067647',
  '#7a5af8',
  '#c2410c',
  '#0e7490',
] as const;

for (const viewport of [
  { height: 800, name: 'desktop', width: 1280 },
  { height: 800, name: 'compact', width: 768 },
] as const) {
  test(`Spreadsheet renders native XLSX linear and path gradients at ${viewport.name} width`, async ({
    page,
  }, testInfo) => {
    const browserErrors: string[] = [];
    page.on('console', (message) => {
      if (message.type() === 'error') browserErrors.push(message.text());
    });
    page.on('pageerror', (error) => browserErrors.push(error.message));
    await page.setViewportSize(viewport);
    await page.goto('/?e2e=spreadsheet-gradient-fill');

    const grid = page.locator('.fortune-sheet-overlay');
    const status = page.getByTestId('spreadsheet-gradient-fill-status');
    await expect(grid).toBeVisible();
    await expect(grid).toBeFocused();
    await expect(status).toHaveAttribute('data-gradient-count', '6');
    await expect(status).toHaveAttribute('data-authoring', 'format-cells');
    await expect(status).toHaveAttribute('data-revision', '1');

    await expect
      .poll(async () => {
        const samples = await spreadsheetCanvasColorSamples(
          page,
          gradientColors,
        );
        return samples.every((count) => count > 20);
      })
      .toBe(true);

    await page.screenshot({
      path: testInfo.outputPath(
        `spreadsheet-gradient-fill-${viewport.name}.png`,
      ),
      animations: 'disabled',
    });

    await grid.focus();
    await page.keyboard.press('Control+1');
    const dialog = page.getByRole('dialog', { name: '设置单元格格式' });
    await expect(dialog).toBeVisible();
    await dialog.getByRole('tab', { name: '填充' }).click();
    await expect(dialog.getByRole('radio', { name: '渐变' })).toBeChecked();
    await expect(
      dialog.getByRole('combobox', { name: '渐变类型' }),
    ).toContainText('线性');
    await expect(
      dialog.getByRole('textbox', { name: '线性渐变角度' }),
    ).toHaveValue('0');

    await dialog.getByRole('combobox', { name: '渐变类型' }).click();
    await page
      .getByRole('listbox', { name: '渐变类型' })
      .getByRole('option', { name: '路径' })
      .click();
    await dialog.getByRole('button', { name: '添加色标' }).click();
    await dialog.getByRole('textbox', { name: '路径渐变左边界' }).fill('20');
    await expect(
      dialog.getByRole('button', { name: /色标 \d+ 颜色/ }),
    ).toHaveCount(3);
    await expectDialogInsideViewport(page, dialog);
    await page.screenshot({
      path: testInfo.outputPath(
        `spreadsheet-gradient-authoring-${viewport.name}.png`,
      ),
      animations: 'disabled',
    });
    await dialog.getByRole('button', { name: '应用' }).click();
    await expect(dialog).toHaveCount(0);
    await expect(grid).toBeFocused();

    await page.keyboard.press('Control+1');
    const reopened = page.getByRole('dialog', { name: '设置单元格格式' });
    await reopened.getByRole('tab', { name: '填充' }).click();
    await expect(
      reopened.getByRole('combobox', { name: '渐变类型' }),
    ).toContainText('路径');
    await expect(
      reopened.getByRole('textbox', { name: '路径渐变左边界' }),
    ).toHaveValue('20');
    await expect(
      reopened.getByRole('button', { name: /色标 \d+ 颜色/ }),
    ).toHaveCount(3);
    await page.keyboard.press('Escape');
    await expect(reopened).toHaveCount(0);
    await expect(grid).toBeFocused();

    await page.keyboard.press('Control+z');
    await page.keyboard.press('Control+1');
    const restored = page.getByRole('dialog', { name: '设置单元格格式' });
    await restored.getByRole('tab', { name: '填充' }).click();
    await expect(
      restored.getByRole('combobox', { name: '渐变类型' }),
    ).toContainText('线性');
    await expect(
      restored.getByRole('button', { name: /色标 \d+ 颜色/ }),
    ).toHaveCount(2);
    await page.keyboard.press('Escape');
    await expect(restored).toHaveCount(0);
    await expect(grid).toBeFocused();
    expect(browserErrors).toEqual([]);
  });
}

async function expectDialogInsideViewport(
  page: Page,
  dialog: Locator,
): Promise<void> {
  const bounds = await dialog.boundingBox();
  const viewport = page.viewportSize();
  expect(bounds).not.toBeNull();
  expect(viewport).not.toBeNull();
  expect(bounds?.x ?? -1).toBeGreaterThanOrEqual(0);
  expect(bounds?.y ?? -1).toBeGreaterThanOrEqual(0);
  expect((bounds?.x ?? 0) + (bounds?.width ?? 0)).toBeLessThanOrEqual(
    viewport?.width ?? 0,
  );
  expect((bounds?.y ?? 0) + (bounds?.height ?? 0)).toBeLessThanOrEqual(
    viewport?.height ?? 0,
  );
}

async function spreadsheetCanvasColorSamples(
  page: Page,
  colors: readonly string[],
): Promise<number[]> {
  return page.evaluate((expectedColors) => {
    const canvas = document.querySelector<HTMLCanvasElement>(
      '.fortune-sheet-canvas',
    );
    const context = canvas?.getContext('2d');
    if (!canvas || !context || !canvas.width || !canvas.height) {
      return expectedColors.map(() => 0);
    }
    const expected = expectedColors.map((color) => [
      Number.parseInt(color.slice(1, 3), 16),
      Number.parseInt(color.slice(3, 5), 16),
      Number.parseInt(color.slice(5, 7), 16),
    ]);
    const counts = expected.map(() => 0);
    const pixels = context.getImageData(0, 0, canvas.width, canvas.height).data;
    for (let offset = 0; offset < pixels.length; offset += 4) {
      if ((pixels[offset + 3] ?? 0) < 200) continue;
      for (const [index, [red, green, blue]] of expected.entries()) {
        if (
          Math.abs((pixels[offset] ?? 0) - red) <= 10 &&
          Math.abs((pixels[offset + 1] ?? 0) - green) <= 10 &&
          Math.abs((pixels[offset + 2] ?? 0) - blue) <= 10
        ) {
          counts[index] = (counts[index] ?? 0) + 1;
          break;
        }
      }
    }
    return counts;
  }, colors);
}
