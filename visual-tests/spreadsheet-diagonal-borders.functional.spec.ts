import { expect, type Locator, type Page, test } from '@playwright/test';
import { openSpreadsheetFixture } from './visual-test-support';

const BORDER_COLOR = '#d84b4f';

test('Spreadsheet preserves both diagonal border directions through WPS editing and row insertion', async ({
  page,
}, testInfo) => {
  const browserErrors: string[] = [];
  page.on('console', (message) => {
    if (message.type() === 'error') browserErrors.push(message.text());
  });
  page.on('pageerror', (error) => browserErrors.push(error.message));
  await openSpreadsheetFixture(page);

  const grid = page.locator('.fortune-sheet-overlay');
  const nameBox = page.locator('.fortune-name-box');
  const ribbon = page.locator('.work-spreadsheet-ribbon');
  await grid.focus();
  await page.keyboard.press('Shift+F11');
  await page.keyboard.press('Control+Home');
  await expect(nameBox).toHaveText('A1');

  const initialBounds = await selectedCellCanvasBounds(page);
  const moreBorders = ribbon.getByRole('button', { name: '更多框线' });
  await moreBorders.click();
  const borderDialog = page.getByRole('dialog', { name: '框线设置' });
  await borderDialog
    .getByRole('combobox', { name: '框线样式' })
    .selectOption('thick');
  await borderDialog.getByLabel('框线颜色').fill(BORDER_COLOR);
  await borderDialog.getByRole('menuitemradio', { name: '斜下框线' }).click();
  await expect(borderDialog).toHaveCount(0);
  await expect(grid).toBeFocused();
  await expectDiagonalDirections(page, initialBounds, {
    down: true,
    up: false,
  });

  let formatDialog = await openFormatCells(page, grid);
  await formatDialog.getByRole('tab', { name: '边框' }).click();
  const diagonalDown = formatDialog.getByRole('button', {
    name: '斜下框线',
  });
  const diagonalUp = formatDialog.getByRole('button', { name: '斜上框线' });
  await expect(diagonalDown).toHaveAttribute('aria-pressed', 'true');
  await expect(diagonalUp).toHaveAttribute('aria-pressed', 'false');
  await diagonalUp.click();
  await expect(diagonalUp).toHaveAttribute('aria-pressed', 'true');
  await formatDialog.getByRole('button', { name: '应用' }).click();
  await expect(formatDialog).toHaveCount(0);
  await expect(grid).toBeFocused();
  await expectDiagonalDirections(page, initialBounds, {
    down: true,
    up: true,
  });

  formatDialog = await openFormatCells(page, grid);
  await formatDialog.getByRole('tab', { name: '边框' }).click();
  await expect(
    formatDialog.getByRole('button', { name: '斜下框线' }),
  ).toHaveAttribute('aria-pressed', 'true');
  await expect(
    formatDialog.getByRole('button', { name: '斜上框线' }),
  ).toHaveAttribute('aria-pressed', 'true');
  await page.keyboard.press('Escape');
  await expect(formatDialog).toHaveCount(0);

  await page.keyboard.press('Control+z');
  await expectDiagonalDirections(page, initialBounds, {
    down: true,
    up: false,
  });
  await page.keyboard.press('Control+Shift+z');
  await expectDiagonalDirections(page, initialBounds, {
    down: true,
    up: true,
  });

  const rowsAndColumns = ribbon.getByRole('button', { name: '行和列' });
  await rowsAndColumns.click();
  await page
    .getByRole('menu', { name: '行和列选项' })
    .getByRole('menuitem', { name: '在上方插入行' })
    .click();
  await expect(grid).toBeFocused();
  await page.keyboard.press('Control+Home');
  await page.keyboard.press('ArrowDown');
  await expect(nameBox).toHaveText('A2');
  const shiftedBounds = await selectedCellCanvasBounds(page);
  await expectDiagonalDirections(page, shiftedBounds, {
    down: true,
    up: true,
  });

  await page.screenshot({
    path: testInfo.outputPath('spreadsheet-diagonal-borders.png'),
    animations: 'disabled',
  });
  expect(browserErrors).toEqual([]);
});

async function openFormatCells(page: Page, grid: Locator): Promise<Locator> {
  await grid.focus();
  await page.keyboard.press('Control+1');
  const dialog = page.getByRole('dialog', { name: '设置单元格格式' });
  await expect(dialog).toBeVisible();
  return dialog;
}

interface CanvasBounds {
  bottom: number;
  left: number;
  right: number;
  top: number;
}

async function selectedCellCanvasBounds(page: Page): Promise<CanvasBounds> {
  return page
    .locator('.luckysheet-cell-selected')
    .last()
    .evaluate((selection) => {
      const canvas = document.querySelector<HTMLCanvasElement>(
        '.fortune-sheet-canvas',
      );
      if (!canvas) throw new Error('Spreadsheet canvas is unavailable.');
      const selectionBounds = selection.getBoundingClientRect();
      const canvasBounds = canvas.getBoundingClientRect();
      return {
        bottom: selectionBounds.bottom - canvasBounds.top,
        left: selectionBounds.left - canvasBounds.left,
        right: selectionBounds.right - canvasBounds.left,
        top: selectionBounds.top - canvasBounds.top,
      };
    });
}

async function expectDiagonalDirections(
  page: Page,
  bounds: CanvasBounds,
  expected: { down: boolean; up: boolean },
): Promise<void> {
  for (const direction of ['down', 'up'] as const) {
    const assertion = expect.poll(() =>
      spreadsheetCanvasDiagonalColorSamples(
        page,
        bounds,
        BORDER_COLOR,
        direction,
      ),
    );
    if (expected[direction]) await assertion.toBeGreaterThan(8);
    else await assertion.toBeLessThan(8);
  }
}

async function spreadsheetCanvasDiagonalColorSamples(
  page: Page,
  bounds: CanvasBounds,
  color: string,
  direction: 'down' | 'up',
): Promise<number> {
  return page.evaluate(
    ({ cell, expected, slope }) => {
      const canvas = document.querySelector<HTMLCanvasElement>(
        '.fortune-sheet-canvas',
      );
      const context = canvas?.getContext('2d');
      const canvasBounds = canvas?.getBoundingClientRect();
      if (!canvas || !context || !canvasBounds?.width || !canvasBounds.height) {
        return 0;
      }
      const scaleX = canvas.width / canvasBounds.width;
      const scaleY = canvas.height / canvasBounds.height;
      const red = Number.parseInt(expected.slice(1, 3), 16);
      const green = Number.parseInt(expected.slice(3, 5), 16);
      const blue = Number.parseInt(expected.slice(5, 7), 16);
      let matches = 0;
      for (let index = 2; index <= 18; index += 1) {
        const ratio = index / 20;
        const x = cell.left + (cell.right - cell.left) * ratio;
        const y =
          slope === 'down'
            ? cell.top + (cell.bottom - cell.top) * ratio
            : cell.bottom - (cell.bottom - cell.top) * ratio;
        let matched = false;
        for (let offsetX = -3; offsetX <= 3 && !matched; offsetX += 1) {
          for (let offsetY = -3; offsetY <= 3; offsetY += 1) {
            const pixel = context.getImageData(
              Math.max(0, Math.round((x + offsetX) * scaleX)),
              Math.max(0, Math.round((y + offsetY) * scaleY)),
              1,
              1,
            ).data;
            if (
              Math.abs((pixel[0] ?? 0) - red) <= 12 &&
              Math.abs((pixel[1] ?? 0) - green) <= 12 &&
              Math.abs((pixel[2] ?? 0) - blue) <= 12 &&
              (pixel[3] ?? 0) > 200
            ) {
              matched = true;
              matches += 1;
              break;
            }
          }
        }
      }
      return matches;
    },
    { cell: bounds, expected: color, slope: direction },
  );
}
