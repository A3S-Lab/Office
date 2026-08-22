import { expect, type Page, test } from '@playwright/test';

const patternForegroundColors = [
  '#1d4ed8',
  '#b42318',
  '#067647',
  '#7a5af8',
] as const;

for (const viewport of [
  { height: 800, name: 'desktop', width: 1280 },
  { height: 800, name: 'compact', width: 768 },
] as const) {
  test(`Spreadsheet renders all native XLSX pattern fills at ${viewport.name} width`, async ({
    page,
  }, testInfo) => {
    const browserErrors: string[] = [];
    page.on('console', (message) => {
      if (message.type() === 'error') browserErrors.push(message.text());
    });
    page.on('pageerror', (error) => browserErrors.push(error.message));
    await page.setViewportSize(viewport);
    await page.goto('/?e2e=spreadsheet-pattern-fill');

    const grid = page.locator('.fortune-sheet-overlay');
    const status = page.getByTestId('spreadsheet-pattern-fill-status');
    await expect(grid).toBeVisible();
    await expect(grid).toBeFocused();
    await expect(status).toHaveAttribute('data-pattern-count', '17');
    await expect(status).toHaveAttribute('data-authoring', 'format-cells');
    await expect(status).toHaveAttribute('data-revision', '1');

    await expect
      .poll(async () => {
        const samples = await spreadsheetCanvasColorSamples(
          page,
          patternForegroundColors,
        );
        return samples.every((count) => count > 100);
      })
      .toBe(true);

    await page.screenshot({
      path: testInfo.outputPath(
        `spreadsheet-pattern-fill-${viewport.name}.png`,
      ),
      animations: 'disabled',
    });

    await grid.focus();
    await page.keyboard.press('Control+1');
    const dialog = page.getByRole('dialog', { name: '设置单元格格式' });
    await expect(dialog).toBeVisible();
    await dialog.getByRole('tab', { name: '填充' }).click();
    await expect(dialog.getByRole('radio', { name: '图案' })).toBeChecked();
    const patternType = dialog.getByRole('combobox', {
      name: '填充图案样式',
    });
    await expect(patternType).toContainText('深色下斜线');
    await patternType.click();
    await page
      .getByRole('listbox', { name: '填充图案样式' })
      .getByRole('option', { name: '深色菱形网格' })
      .click();
    await expect(patternType).toContainText('深色菱形网格');
    await page.screenshot({
      path: testInfo.outputPath(
        `spreadsheet-pattern-authoring-${viewport.name}.png`,
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
      reopened.getByRole('combobox', { name: '填充图案样式' }),
    ).toContainText('深色菱形网格');
    await page.keyboard.press('Escape');
    await expect(reopened).toHaveCount(0);
    await expect(grid).toBeFocused();

    await page.keyboard.press('Control+z');
    await page.keyboard.press('Control+1');
    const restored = page.getByRole('dialog', { name: '设置单元格格式' });
    await restored.getByRole('tab', { name: '填充' }).click();
    await expect(
      restored.getByRole('combobox', { name: '填充图案样式' }),
    ).toContainText('深色下斜线');
    await page.keyboard.press('Escape');
    await expect(restored).toHaveCount(0);
    await expect(grid).toBeFocused();
    expect(browserErrors).toEqual([]);
  });
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
          Math.abs((pixels[offset] ?? 0) - red) <= 8 &&
          Math.abs((pixels[offset + 1] ?? 0) - green) <= 8 &&
          Math.abs((pixels[offset + 2] ?? 0) - blue) <= 8
        ) {
          counts[index] = (counts[index] ?? 0) + 1;
          break;
        }
      }
    }
    return counts;
  }, colors);
}
