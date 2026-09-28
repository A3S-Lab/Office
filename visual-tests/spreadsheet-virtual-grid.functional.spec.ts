import AxeBuilder from '@axe-core/playwright';
import { expect, test } from '@playwright/test';

/**
 * Phase 4 opt-in A3S virtual grid: prove grid ARIA ownership and axe on the
 * owned surface (Fortune host stays aria-hidden underneath).
 */
test.describe('Spreadsheet A3S virtual grid', () => {
  test('owns role=grid focus and passes axe on the grid root', async ({
    page,
  }) => {
    await page.goto('/playground/?virtualGrid=1');
    await page
      .getByRole('button', {
        name: '季度执行计划 XLSX · 本次会话',
      })
      .click();
    const grid = page.locator('.work-spreadsheet-virtual-grid');
    await grid.waitFor({ state: 'visible' });
    await expect(grid).toHaveAttribute('role', 'grid');
    await expect(grid).toHaveAttribute('aria-activedescendant', /a3s-ss-cell-/);

    await grid.focus();
    await page.keyboard.press('ArrowDown');
    await expect(grid).toHaveAttribute(
      'aria-activedescendant',
      'a3s-ss-cell-1-0',
    );

    const results = await new AxeBuilder({ page })
      .include('.work-spreadsheet-virtual-grid')
      .withTags(['wcag2a', 'wcag2aa', 'wcag21a', 'wcag21aa'])
      .analyze();
    const blocking = results.violations.filter((violation) =>
      ['serious', 'critical'].includes(violation.impact ?? ''),
    );
    expect(
      blocking,
      blocking
        .map(
          (violation) =>
            `${violation.id} (${violation.impact}): ${violation.help}`,
        )
        .join('\n'),
    ).toEqual([]);
  });
});
