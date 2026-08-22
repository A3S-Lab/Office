import { expect, type Page, test } from '@playwright/test';

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
